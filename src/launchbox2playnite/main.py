"""Export LaunchBox data to Playnite-compatible YAML files."""

from __future__ import annotations

import os
from collections import defaultdict
from itertools import product
from multiprocessing import get_context
from pathlib import Path
from threading import Thread
from typing import TYPE_CHECKING
from typing import Any
from typing import cast
from uuid import uuid4

import yaml
from defusedxml import ElementTree
from naay import loads as naay_loads
from PIL import Image
from PIL import UnidentifiedImageError
from tqdm import tqdm

if TYPE_CHECKING:  # pragma: no cover - import-time typing helpers only
    from collections.abc import Callable
    from collections.abc import Iterator
    from collections.abc import Sequence
    from typing import Protocol

    class ProgressQueue(Protocol):
        """Protocol describing the minimal queue API used for progress."""

        def put(self, value: int | None) -> None:
            """Enqueue a progress delta or sentinel."""

        def get(self) -> int | None:
            """Return the next queued progress delta or sentinel."""

    from xml.etree.ElementTree import Element as XmlElement  # noqa: S405

ConfigDict = dict[str, Any]
GameDict = dict[str, Any]
PlaylistDict = dict[str, Any]
FolderDict = dict[str, Any]
MediaDict = dict[str, Any]
StrSetDefaultDict = defaultdict[str, set[str]]


class _QuotedDumper(yaml.SafeDumper):
    """YAML dumper that forces every string to be double-quoted."""


def _quoted_str_representer(dumper: yaml.SafeDumper, data: str) -> yaml.nodes.Node:
    """Serialize a Python string as a quoted YAML scalar.

    Returns:
        yaml.nodes.Node: The double-quoted YAML node.
    """
    represent_scalar = cast("Any", dumper.represent_scalar)
    return represent_scalar("tag:yaml.org,2002:str", data, style='"')


_QuotedDumper.add_representer(str, _quoted_str_representer)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_PATH = PROJECT_ROOT / "config.yaml"


def load_config() -> ConfigDict:
    """Load and validate ``config.yaml`` using the strict naay parser.

    Returns:
        ConfigDict: The parsed configuration mapping.

    Raises:
        SystemExit: If the file is missing or malformed.
    """
    try:
        text = CONFIG_PATH.read_text(encoding="utf-8")
    except FileNotFoundError as exc:
        msg = f"Config file not found: {CONFIG_PATH}"
        raise SystemExit(msg) from exc

    data = naay_loads(text)
    if not isinstance(data, dict):
        msg = "config.yaml must contain a top-level mapping"
        raise SystemExit(msg)
    return cast("ConfigDict", data)


def require_path(config: ConfigDict, key: str) -> Path:
    """Return an absolute ``Path`` for the provided config key.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a filesystem path.

    Returns:
        Path: The resolved absolute path.

    Raises:
        SystemExit: If the key is missing or not a string.
    """
    if key not in config:
        msg = f"Missing '{key}' in config.yaml"
        raise SystemExit(msg)
    raw = config[key]
    if not isinstance(raw, str):
        msg = f"Config value '{key}' must be a string path"
        raise SystemExit(msg)
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = (CONFIG_PATH.parent / path).resolve()
    return path


CONFIG: ConfigDict = load_config()


def require_mapping(config: ConfigDict, key: str) -> ConfigDict:
    """Fetch a nested mapping from the config file.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a nested mapping.

    Returns:
        ConfigDict: The nested mapping value.

    Raises:
        SystemExit: If the key is missing or not a mapping.
    """
    if key not in config:
        msg = f"Missing '{key}' mapping in config.yaml"
        raise SystemExit(msg)
    value = config[key]
    if not isinstance(value, dict):
        msg = f"Config value '{key}' must be a mapping"
        raise SystemExit(msg)
    return cast("ConfigDict", value)


def require_str(config: ConfigDict, key: str) -> str:
    """Return a string value for ``key`` from the config.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a string value.

    Returns:
        str: The stored configuration value.

    Raises:
        SystemExit: If the key is missing or the value is not a string.
    """
    if key not in config:
        msg = f"Missing '{key}' in config.yaml"
        raise SystemExit(msg)
    value = config[key]
    if not isinstance(value, str):
        msg = f"Config value '{key}' must be a string"
        raise SystemExit(msg)
    return value


def require_int(config: ConfigDict, key: str) -> int:
    """Return an integer from the config, coercing numeric strings when needed.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold an integer or integer string.

    Returns:
        int: The normalized integer value.

    Raises:
        SystemExit: If the key is missing or not an integer-like value.
    """
    if key not in config:
        msg = f"Missing '{key}' in config.yaml"
        raise SystemExit(msg)
    value = config[key]
    if isinstance(value, bool):
        msg = f"Config value '{key}' must be an integer"
        raise SystemExit(msg)
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        try:
            return int(value.strip())
        except ValueError as exc:
            msg = f"Config value '{key}' must be an integer string"
            raise SystemExit(msg) from exc
    msg = f"Config value '{key}' must be an integer"
    raise SystemExit(msg)


def require_float(config: ConfigDict, key: str) -> float:
    """Return a float from the config, coercing numeric strings when needed.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a numeric value.

    Returns:
        float: The normalized floating-point value.

    Raises:
        SystemExit: If the key is missing or not numeric.
    """
    if key not in config:
        msg = f"Missing '{key}' in config.yaml"
        raise SystemExit(msg)
    value = config[key]
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            msg = f"Config value '{key}' must be a number"
            return float(value)
        except ValueError as exc:
            msg = f"Config value '{key}' must be a number"
            raise SystemExit(msg) from exc
    msg = f"Config value '{key}' must be a number"
    raise SystemExit(msg)


def require_str_list(config: ConfigDict, key: str) -> tuple[str, ...]:
    """Return a tuple of strings for iterable config values.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a list of strings.

    Returns:
        tuple[str, ...]: The validated string values.

    Raises:
        SystemExit: If the key is missing or contains non-strings.
    """
    if key not in config:
        msg = f"Missing '{key}' in config.yaml"
        raise SystemExit(msg)
    value = config[key]
    if not isinstance(value, list):
        msg = f"Config value '{key}' must be a list"
        raise SystemExit(msg)
    value_list = cast("list[Any]", value)
    cleaned: list[str] = []
    for idx, item in enumerate(value_list):
        if not isinstance(item, str):
            msg = f"Config list '{key}' must contain only strings (index {idx})"
            raise SystemExit(msg)
        cleaned.append(item)
    return tuple(cleaned)


def require_project_path(config: ConfigDict, key: str) -> Path:
    """Resolve a path relative to the project root when needed.

    Args:
        config: The loaded configuration mapping.
        key: Key expected to hold a relative or absolute path.

    Returns:
        Path: The absolute path anchored at the project root.
    """
    raw = require_str(config, key)
    path = Path(raw)
    if not path.is_absolute():
        path = (PROJECT_ROOT / path).resolve()
    return path


def dump_yaml(data: Any, path: Path) -> None:
    """Serialize ``data`` to ``path`` using the custom quoted dumper."""
    with Path(path).open("w", encoding="utf-8") as f:
        yaml.dump(
            data,
            f,
            sort_keys=False,
            allow_unicode=True,
            width=YAML_WIDTH,
            Dumper=_QuotedDumper,
        )


# ----------------------------------------------------
# CONFIG - driven by config.yaml
# ----------------------------------------------------
LAUNCHBOX_ROOT = require_path(CONFIG, "exo_root")

paths_cfg = require_mapping(CONFIG, "paths")
PLATFORMS_DIR = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "platforms"))
PLAYLIST_DIR = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "playlists"))
PARENTS_FILE = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "parents"))

IMAGES_DIR = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "images"))
VIDEOS_DIR = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "videos"))
MANUALS_DIR = LAUNCHBOX_ROOT / Path(require_str(paths_cfg, "manuals"))

outputs_cfg = require_mapping(CONFIG, "outputs")
OUTPUT_GAMES = require_project_path(outputs_cfg, "games")
OUTPUT_PLAYLISTS = require_project_path(outputs_cfg, "playlists")
OUTPUT_FOLDERS = require_project_path(outputs_cfg, "folders")

YAML_WIDTH = require_int(CONFIG, "yaml_width")

ROOT_CATEGORY_NAME = require_str(CONFIG, "root_category_name")
IMAGE_EXTENSIONS = require_str_list(CONFIG, "image_extensions")

generated_icon_cfg = require_mapping(CONFIG, "generated_icon")
GENERATED_ICON_DIR = require_project_path(generated_icon_cfg, "dir")
ICON_SIZE = require_int(generated_icon_cfg, "size")

normalized_cover_cfg = require_mapping(CONFIG, "normalized_cover")
COVER_DIR = require_project_path(normalized_cover_cfg, "dir")
COVER_WIDTH = require_int(normalized_cover_cfg, "min_width")
COVER_HEIGHT = require_int(normalized_cover_cfg, "min_height")
COVER_ASPECT = COVER_WIDTH / COVER_HEIGHT
COVER_MAX_STRETCH = require_float(normalized_cover_cfg, "max_stretch")
PROGRESS_LOG_INTERVAL = 200


# ----------------------------------------------------
# Helpers
# ----------------------------------------------------
def norm_key(s: str) -> str:
    """Return a lowercase alphanumeric key for fuzzy comparisons.

    Args:
        s: Original platform or playlist name.

    Returns:
        str: A normalized token suitable for dictionary lookups.
    """
    return "".join(ch for ch in s.lower() if ch.isalnum())


def normalize_title(title: str) -> str:
    """Strip punctuation that confuses Windows filenames.

    Args:
        title: Original LaunchBox title.

    Returns:
        str: The sanitized title.
    """
    return title.replace(":", "").replace("?", "").strip()


def _add_candidate(candidates: list[str], value: str | None) -> None:
    if not value:
        return
    cleaned = value.strip()
    if cleaned and cleaned not in candidates:
        candidates.append(cleaned)


def _raw_title_variants(title: str) -> set[str]:
    variants = {title}
    replacement_chars = (":", "!", "?", "*", '"', "<", ">", "|")
    replacements = ("_", " ", "")
    for ch in replacement_chars:
        updates = {
            variant.replace(ch, repl)
            for variant in variants
            if ch in variant
            for repl in replacements
        }
        variants.update(updates)
    return variants


def _base_variants(base: str) -> set[str]:
    if not base:
        return set()
    variants = {base}
    apostrophes = ("'", "\u2019")
    for replacement in ("_", "", " "):
        variant = base
        for mark in apostrophes:
            variant = variant.replace(mark, replacement)
        variants.add(variant)
    variants.add(base.replace(" ", "_"))
    variants.add(base.replace(" ", ""))
    return {value for value in variants if value}


def media_name_candidates(title: str) -> list[str]:
    """Generate plausible filename bases for the provided title.

    Args:
        title: The LaunchBox title to expand.

    Returns:
        list[str]: Ordered candidates, most likely first.
    """
    candidates: list[str] = []
    for raw in _raw_title_variants(title):
        base = normalize_title(raw)
        for variant in _base_variants(base):
            _add_candidate(candidates, variant)
    return candidates


def normalized_media_key(value: str) -> str:
    """Lowercase and strip punctuation to build a matching key.

    Args:
        value: The raw filename component to normalize.

    Returns:
        str: A simplified key suited for fuzzy comparisons.
    """
    cleaned = value.replace("_", " ").replace("-", " ")
    cleaned = cleaned.replace("\u2019", "'")
    return normalize_title(cleaned).lower()


def safe_filename(title: str, suffix: str) -> str:
    """Convert a title into a filesystem-friendly filename.

    Args:
        title: Game title used as the base.
        suffix: File extension without the dot.

    Returns:
        str: A sanitized filename including the suffix.
    """
    base = normalize_title(title)
    if not base:
        base = "icon"
    safe = "".join(ch if ch.isalnum() else "-" for ch in base.lower())
    while "--" in safe:
        safe = safe.replace("--", "-")
    safe = safe.strip("-") or "icon"
    return f"{safe}.{suffix}"


def get_image_dimensions(path: Path) -> tuple[int, int] | None:
    """Measure an image while swallowing unreadable files.

    Args:
        path: Image file to inspect.

    Returns:
        tuple[int, int] | None: Image width/height or ``None`` on failure.
    """
    try:
        with Image.open(path) as img_obj:
            img: Image.Image = img_obj
            return img.width, img.height
    except (OSError, UnidentifiedImageError, ValueError):
        return None


def generate_icon_from_cover(title: str, cover_path: str | None) -> str | None:
    """Build a square icon derived from a cover image if possible.

    Args:
        title: Title used to label generated assets.
        cover_path: Source cover image.

    Returns:
        str | None: Absolute path to the generated icon, if any.
    """
    if not cover_path:
        return None

    cover_file = Path(cover_path)
    if not cover_file.exists():
        return None

    GENERATED_ICON_DIR.mkdir(parents=True, exist_ok=True)
    filename = safe_filename(cover_file.stem or title, "png")
    dest = GENERATED_ICON_DIR / filename

    try:
        if dest.exists() and dest.stat().st_mtime >= cover_file.stat().st_mtime:
            return str(dest.resolve())

        with Image.open(cover_file) as img_obj:
            img: Image.Image = img_obj.convert("RGBA")
            img.thumbnail((ICON_SIZE, ICON_SIZE))
            background = Image.new("RGBA", (ICON_SIZE, ICON_SIZE), (0, 0, 0, 0))
            offset = (
                (ICON_SIZE - img.width) // 2,
                (ICON_SIZE - img.height) // 2,
            )
            background.paste(img, offset)
            background.save(dest, format="PNG")

        return str(dest.resolve())
    except (OSError, UnidentifiedImageError, ValueError) as exc:
        print(f"⚠ Failed to generate icon for '{title}': {exc}")
        return None


def normalize_cover_image(
    title: str, cover_path: str | None, target_width: int | None = None
) -> str | None:
    """Resize and pad cover art to the configured aspect ratio.

    Args:
        title: Title used to name generated assets.
        cover_path: Source cover image path.
        target_width: Optional override for the minimum width.

    Returns:
        str | None: Absolute path to the normalized cover, if produced.
    """
    if not cover_path:
        return None

    source = Path(cover_path)
    if not source.exists():
        return None

    COVER_DIR.mkdir(parents=True, exist_ok=True)
    filename = safe_filename(source.stem or title, "png")
    dest = COVER_DIR / filename

    try:
        if dest.exists() and dest.stat().st_mtime >= source.stat().st_mtime:
            return str(dest.resolve())

        with Image.open(source) as img_obj:
            img: Image.Image = img_obj.convert("RGBA")
            src_aspect = img.width / img.height if img.height else COVER_ASPECT

            target_width_px = (
                max(COVER_WIDTH, int(target_width)) if target_width else COVER_WIDTH
            )
            target_height_px = max(1, round(target_width_px / COVER_ASPECT))

            scale_by_height = target_height_px / img.height if img.height else 1.0
            scale_by_width = target_width_px / img.width if img.width else 1.0

            scale = scale_by_height if src_aspect > COVER_ASPECT else scale_by_width

            scale = min(scale, COVER_MAX_STRETCH)

            new_size = (
                max(1, int(img.width * scale)),
                max(1, int(img.height * scale)),
            )

            img_any = cast("Any", img)
            resized = cast(
                "Image.Image",
                img_any.resize(new_size, Image.Resampling.LANCZOS),
            )

            canvas = Image.new(
                "RGBA", (target_width_px, target_height_px), (0, 0, 0, 255)
            )
            offset = (
                (target_width_px - resized.width) // 2,
                (target_height_px - resized.height) // 2,
            )
            canvas.paste(resized, offset)
            canvas.save(dest, format="PNG")

        return str(dest.resolve())
    except (OSError, UnidentifiedImageError, ValueError) as exc:
        print(f"⚠ Failed to normalize cover for '{title}': {exc}")
        return None


def damerau_levenshtein(a: str, b: str, max_distance: int) -> int:
    """Compute a bounded Damerau-Levenshtein distance.

    Args:
        a: First string being compared.
        b: Second string being compared.
        max_distance: Early-exit threshold for the distance.

    Returns:
        int: The edit distance, capped at ``max_distance + 1``.
    """
    if abs(len(a) - len(b)) > max_distance:
        return max_distance + 1

    len_a = len(a)
    len_b = len(b)

    if len_a == 0:
        return len_b
    if len_b == 0:
        return len_a

    d = [[0] * (len_b + 1) for _ in range(len_a + 1)]

    for i in range(len_a + 1):
        d[i][0] = i
    for j in range(len_b + 1):
        d[0][j] = j

    for i in range(1, len_a + 1):
        best_in_row = max_distance + 1
        for j in range(1, len_b + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            d[i][j] = min(
                d[i - 1][j] + 1,  # deletion
                d[i][j - 1] + 1,  # insertion
                d[i - 1][j - 1] + cost,  # substitution
            )

            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + cost)

            best_in_row = min(best_in_row, d[i][j])

        if best_in_row > max_distance:
            return max_distance + 1

    return d[len_a][len_b]


def find_fuzzy_image(
    folder: Path, title: str, extensions: Sequence[str], max_distance: int = 2
) -> str | None:
    """Search ``folder`` recursively for an approximate media match.

    Args:
        folder: Root folder containing potential matches.
        title: Reference title to match against.
        extensions: Allowed file extensions.
        max_distance: Maximum edit distance before giving up.

    Returns:
        str | None: The best-matching file path or ``None`` if missing.
    """
    if not folder.exists():
        return None

    target = normalized_media_key(title)
    best_path: str | None = None
    best_distance = max_distance + 1

    for ext in extensions:
        for file in folder.glob(f"**/*.{ext}"):
            candidate = normalized_media_key(file.stem)
            distance = damerau_levenshtein(target, candidate, max_distance)
            if distance >= best_distance:
                continue
            best_distance = distance
            best_path = str(file)
            if distance == 0:
                return best_path

    if best_distance <= max_distance:
        return best_path
    return None


def find_platform_dir(root: Path, platform_name: str | None) -> Path | None:
    """Locate a platform-specific subdirectory, if it exists.

    Args:
        root: Root folder to inspect.
        platform_name: LaunchBox platform name.

    Returns:
        Path | None: Matching directory, if discovered.
    """
    if not platform_name or not root.exists():
        return None

    def try_match(base: Path) -> Path | None:
        direct = base / platform_name
        if direct.exists():
            return direct

        target = norm_key(platform_name)
        for child in base.iterdir():
            if child.is_dir() and norm_key(child.name) == target:
                return child
        return None

    for candidate_base in (root, root / "Platforms", root / "Platform Categories"):
        if candidate_base.exists():
            match = try_match(candidate_base)
            if match:
                return match
    return None


def _iter_media_paths(
    folder: Path, candidates: Sequence[str], extensions: Sequence[str]
) -> Iterator[str]:
    for ext, candidate in product(extensions, candidates):
        pattern = f"**/{candidate}*.{ext}"
        for file_path in folder.glob(pattern):
            yield str(file_path)


def gather_media_files(
    title: str, folders: Sequence[Path | None], extensions: Sequence[str]
) -> list[str]:
    """Gather all matching media files, deduplicating by absolute path.

    Args:
        title: Title used to generate filename candidates.
        folders: Candidate directories to search.
        extensions: File extensions to include.

    Returns:
        list[str]: Sorted by discovery order without duplicates.
    """
    matches: list[str] = []
    seen: set[str] = set()
    candidates = media_name_candidates(title)

    for folder in folders:
        if folder is None or not folder.exists():
            continue
        for path in _iter_media_paths(folder, candidates, extensions):
            if path in seen:
                continue
            seen.add(path)
            matches.append(path)

    return matches


def find_first_image(folder: Path, title: str) -> str | None:
    """Return the first matching image path, falling back to fuzzy search."""
    if not folder.exists():
        return None
    candidates = media_name_candidates(title)
    for ext in IMAGE_EXTENSIONS:
        for candidate in candidates:
            pattern = f"**/{candidate}*.{ext}"
            for file in folder.glob(pattern):
                return str(file)
    return find_fuzzy_image(folder, title, IMAGE_EXTENSIONS)


def collect_images(title: str, platform_name: str | None) -> MediaDict:  # noqa: C901
    """Aggregate cover, icon, and supporting imagery for a game.

    Args:
        title: Game title used for filename guesses.
        platform_name: LaunchBox platform, if known.

    Returns:
        MediaDict: Paths to collected assets keyed by media type.
    """
    media: MediaDict = {}
    platform_dir = find_platform_dir(IMAGES_DIR, platform_name)

    platform_candidates: list[Path | None] = [platform_dir]
    for extra_base in (IMAGES_DIR / "Platforms", IMAGES_DIR / "Platform Categories"):
        extra = find_platform_dir(extra_base, platform_name)
        if extra and extra not in platform_candidates:
            platform_candidates.append(extra)

    def iter_folders(subfolder: str) -> Iterator[Path]:
        if platform_candidates:
            for base in platform_candidates:
                if base:
                    yield base / subfolder
        yield IMAGES_DIR / subfolder

    def first_from(*subfolders: str) -> str | None:
        for subfolder in subfolders:
            if not subfolder:
                continue
            for folder in iter_folders(subfolder):
                path = find_first_image(folder, title)
                if path:
                    return path
        return None

    cover_path = first_from(
        "Box - Front",
        "Fanart - Box - Front",
        "Box - Front - Reconstructed",
        "Fanart - Box - Front - Reconstructed",
        "Box - 3D",
        "Box - Full",
    )
    media["cover"] = normalize_cover_image(title, cover_path) or cover_path
    media["icon"] = first_from("Icon", "Clear Logo")
    if not media["icon"] and media.get("cover"):
        generated_icon = generate_icon_from_cover(title, media["cover"])
        if generated_icon:
            media["icon"] = generated_icon
    media["background"] = first_from(
        "Fanart - Background",
        "Background",
        "Banner",
        "Steam Banner",
        "Amazon Background",
        "Epic Games Background",
        "Origin Background",
        "Uplay Background",
    )

    screenshot_folders = [
        base / "Screenshot - Gameplay"
        for base in platform_candidates
        if base is not None
    ]
    screenshot_folders.append(IMAGES_DIR / "Screenshot - Gameplay")
    media["screenshots"] = gather_media_files(
        title, screenshot_folders, ("png", "jpg", "jpeg", "webp")
    )
    return media


def collect_videos(title: str, platform_name: str | None) -> list[str]:
    """Locate gameplay videos for the provided title and platform.

    Args:
        title: Game title used for filename guesses.
        platform_name: LaunchBox platform, if known.

    Returns:
        list[str]: Ordered list of matching video files.
    """
    dirs: list[Path] = []
    platform_dir = find_platform_dir(VIDEOS_DIR, platform_name)
    if platform_dir:
        dirs.append(platform_dir)
    dirs.append(VIDEOS_DIR)
    return gather_media_files(title, dirs, ("mp4", "avi", "mkv", "webm"))


def collect_manual(title: str, platform_name: str | None) -> str | None:
    """Return the first matching manual or walkthrough document.

    Args:
        title: Game title used for filename guesses.
        platform_name: LaunchBox platform, if known.

    Returns:
        str | None: The first matching manual path, if present.
    """
    dirs: list[Path] = []
    platform_dir = find_platform_dir(MANUALS_DIR, platform_name)
    if platform_dir:
        dirs.append(platform_dir)
    dirs.append(MANUALS_DIR)
    matches = gather_media_files(title, dirs, ("pdf", "txt", "cbz", "cbr"))
    return matches[0] if matches else None


# ----------------------------------------------------
# Game parsing
# ----------------------------------------------------
def parse_launchbox_game(  # noqa: C901, PLR0914
    game_elem: XmlElement, platform_name: str
) -> GameDict:
    """Convert a LaunchBox ``Game`` XML node to a Playnite dictionary.

    Args:
        game_elem: XML element describing the LaunchBox game.
        platform_name: Name of the platform containing the game.

    Returns:
        GameDict: The normalized Playnite representation.
    """

    def get(tag: str) -> str | None:
        elem = game_elem.find(tag)
        if elem is not None and elem.text:
            return elem.text.strip()
        return None

    title = get("Title") or "Unknown Game"
    sort_title = get("SortTitle") or title
    release_date = get("ReleaseDate")
    developer = get("Developer")
    publisher = get("Publisher")
    genre_raw = get("Genre")
    description = get("Notes")

    images = collect_images(title, platform_name)
    videos = collect_videos(title, platform_name)
    manual = collect_manual(title, platform_name)

    game: GameDict = {
        "Id": str(uuid4()),
        "Name": title,
        "SortingName": sort_title,
        "Platform": platform_name,
        "ReleaseYear": release_date[:4] if release_date else None,
        "Developers": [developer] if developer else [],
        "Publishers": [publisher] if publisher else [],
        "Genres": [g.strip() for g in genre_raw.split(";") if g.strip()]
        if genre_raw
        else [],
        "Description": description,
    }

    rom_path = get("RomPath")
    if rom_path:
        game["Roms"] = [{"Path": rom_path}]

    exe_path = get("ApplicationPath")
    if exe_path:
        game["PlayAction"] = {
            "Path": exe_path,
            "Arguments": get("CommandLine") or "",
            "WorkingDir": str(Path(exe_path).parent),
        }

    cover = images.get("cover")
    if cover:
        game["Image"] = cover
    background = images.get("background")
    if background:
        game["BackgroundImage"] = background
    icon = images.get("icon")
    if icon:
        game["Icon"] = icon
    screenshots = images.get("screenshots")
    if screenshots:
        game["Screenshots"] = screenshots
    if videos:
        game["Videos"] = videos
    if manual:
        game["Manual"] = manual

    return {k: v for k, v in game.items() if v not in ("", None, [], {})}


def _process_platform_file(
    platform_path: str,
    progress_cb: Callable[[int], None] | None = None,
) -> tuple[str, list[GameDict], dict[str, GameDict]]:
    """Parse a single LaunchBox platform XML file.

    Args:
        platform_path: Absolute path to the platform XML file.
        progress_cb: Optional callback invoked after each parsed game.

    Returns:
        tuple[str, list[GameDict], dict[str, GameDict]]: Platform name, list of
        games, and LaunchBox ID lookup table.

    Raises:
        ValueError: If the XML file lacks a root element.
    """
    platform_file = Path(platform_path)
    platform_name = platform_file.stem

    tree = ElementTree.parse(platform_file)
    root = tree.getroot()
    if root is None:
        msg = f"Platform file '{platform_file}' has no root element"
        raise ValueError(msg)
    game_nodes = list(root.findall("Game"))
    total_games = len(game_nodes)
    tqdm.write(f"Found {total_games} games in {platform_name}")

    games: list[GameDict] = []
    games_by_lb_id: dict[str, GameDict] = {}

    for idx, game_elem in enumerate(game_nodes, start=1):
        game = parse_launchbox_game(game_elem, platform_name)

        lb_id_elem = game_elem.find("ID")
        if lb_id_elem is None:
            lb_id_elem = game_elem.find("Id")
        if lb_id_elem is not None and lb_id_elem.text:
            lb_id = lb_id_elem.text.strip()
            game["LaunchBoxId"] = lb_id
            games_by_lb_id[lb_id] = game

        games.append(game)

        if total_games >= PROGRESS_LOG_INTERVAL and idx % PROGRESS_LOG_INTERVAL == 0:
            tqdm.write(f"… {platform_name}: processed {idx}/{total_games} games")
        if progress_cb is not None:
            progress_cb(1)

    return platform_name, games, games_by_lb_id


def _process_platform_file_worker(
    args: tuple[str, ProgressQueue],
) -> tuple[str, str, list[GameDict], dict[str, GameDict]]:
    """Parse a platform file inside a worker process and push progress.

    Returns:
        tuple[str, str, list[GameDict], dict[str, GameDict]]: The original
        path plus parsed platform outputs.
    """
    platform_path, progress_queue = args

    def queue_progress(count: int) -> None:
        progress_queue.put(count)

    platform_name, games, lb_map = _process_platform_file(
        platform_path,
        progress_cb=queue_progress,
    )
    return platform_path, platform_name, games, lb_map


def _drain_progress_queue(
    progress_queue: ProgressQueue,
    game_bar: tqdm[Any],
) -> None:
    """Drain the multiprocessing progress queue into the tqdm bar."""
    while True:
        count = progress_queue.get()
        if count is None:
            break
        game_bar.update(count)


# ----------------------------------------------------
# Playlists: Data\Playlists\*.xml
# ----------------------------------------------------
def parse_playlists(
    games_by_lb_id: dict[str, GameDict],
) -> tuple[list[PlaylistDict], dict[str, PlaylistDict]]:
    """Build playlists with a lookup table keyed by LaunchBox IDs.

    Args:
        games_by_lb_id: Map of LaunchBox game GUIDs to Playnite game dicts.

    Returns:
        tuple[list[PlaylistDict], dict[str, PlaylistDict]]: A list of playlists
        plus a helper dict for quick ID to playlist resolution.
    """
    playlists: list[PlaylistDict] = []
    playlists_by_lb_id: dict[str, PlaylistDict] = {}

    if not PLAYLIST_DIR.exists():
        return playlists, playlists_by_lb_id

    for file in PLAYLIST_DIR.glob("*.xml"):
        tree = ElementTree.parse(file)
        root = tree.getroot()
        if root is None:
            continue

        name_elem = root.find("Name")
        pl_name = (
            name_elem.text.strip()
            if name_elem is not None and name_elem.text
            else file.stem
        )

        # LaunchBox playlist GUID
        id_elem = root.find("Id") or root.find("ID")
        if id_elem is None or not id_elem.text:
            # some older lists may not have Id; skip if so
            continue
        lb_id = id_elem.text.strip()

        game_ids: list[str] = []
        for pg in root.findall("PlaylistGame"):
            gid_elem = pg.find("GameId")
            if gid_elem is None or not gid_elem.text:
                continue
            lb_game_id = gid_elem.text.strip()
            g = games_by_lb_id.get(lb_game_id)
            if g:
                game_ids.append(g["Id"])

        playlist: PlaylistDict = {
            "Id": str(uuid4()),  # Playnite internal ID
            "Name": pl_name,
            "Description": None,
            "GameIds": game_ids,
            "LaunchBoxId": lb_id,
        }

        playlists.append(playlist)
        playlists_by_lb_id[lb_id] = playlist

    return playlists, playlists_by_lb_id


# ----------------------------------------------------
# Parents.xml → hierarchy
# ----------------------------------------------------
def parse_parents(
    playlists_by_lb_id: dict[str, PlaylistDict],
) -> tuple[
    StrSetDefaultDict,
    StrSetDefaultDict,
    StrSetDefaultDict,
    StrSetDefaultDict,
    StrSetDefaultDict,
    set[str],
]:
    """Parse ``Parents.xml`` and describe LaunchBox relationships.

    Args:
        playlists_by_lb_id: Helper map for validating playlist IDs.

    Returns:
        tuple: Five ``defaultdict(set)`` instances describing category/platform
        relationships plus the set of root category names.
    """
    category_children_categories: StrSetDefaultDict = defaultdict(set)
    category_children_platforms: StrSetDefaultDict = defaultdict(set)
    category_children_playlists: StrSetDefaultDict = defaultdict(set)
    platform_children_platforms: StrSetDefaultDict = defaultdict(set)
    platform_children_playlists: StrSetDefaultDict = defaultdict(set)

    root_categories: set[str] = set()

    if not PARENTS_FILE.exists():
        print(f"⚠ Parents.xml not found at {PARENTS_FILE}")
        return (
            category_children_categories,
            category_children_platforms,
            category_children_playlists,
            platform_children_platforms,
            platform_children_playlists,
            root_categories,
        )

    tree = ElementTree.parse(PARENTS_FILE)
    root = tree.getroot()
    if root is None:
        return (
            category_children_categories,
            category_children_platforms,
            category_children_playlists,
            platform_children_platforms,
            platform_children_playlists,
            root_categories,
        )

    for parent in root.findall("Parent"):
        platform_name = (parent.findtext("PlatformName") or "").strip()
        playlist_id = (parent.findtext("PlaylistId") or "").strip()
        cat_name = (parent.findtext("PlatformCategoryName") or "").strip()

        parent_platform = (parent.findtext("ParentPlatformName") or "").strip()
        parent_cat = (parent.findtext("ParentPlatformCategoryName") or "").strip()

        # detect root categories
        if cat_name and not parent_cat and not platform_name and not playlist_id:
            root_categories.add(cat_name)

        # category -> subcategory
        if cat_name and parent_cat:
            category_children_categories[parent_cat].add(cat_name)

        # category -> platform
        if platform_name and parent_cat:
            category_children_platforms[parent_cat].add(platform_name)

        # category -> playlist
        if playlist_id and parent_cat and playlist_id in playlists_by_lb_id:
            category_children_playlists[parent_cat].add(playlist_id)

        # platform -> playlist
        if playlist_id and parent_platform and playlist_id in playlists_by_lb_id:
            platform_children_playlists[parent_platform].add(playlist_id)

        # platform -> sub-platform
        if platform_name and parent_platform:
            platform_children_platforms[parent_platform].add(platform_name)

    return (
        category_children_categories,
        category_children_platforms,
        category_children_playlists,
        platform_children_platforms,
        platform_children_playlists,
        root_categories,
    )


# ----------------------------------------------------
# Build Playnite folder tree from parents
# ----------------------------------------------------
def build_folder_tree_from_parents(  # noqa: PLR0913, PLR0917
    root_category_name: str,
    category_children_categories: StrSetDefaultDict,
    category_children_platforms: StrSetDefaultDict,
    category_children_playlists: StrSetDefaultDict,
    platform_children_platforms: StrSetDefaultDict,
    platform_children_playlists: StrSetDefaultDict,
    games_by_platform_norm: dict[str, list[str]],
    playlists_by_lb_id: dict[str, PlaylistDict],
) -> list[FolderDict]:
    """Expand parsed parent relationships into Playnite folders.

    Args:
        root_category_name: Top-level category to expand.
        category_children_categories: Category-to-subcategory mapping.
        category_children_platforms: Category-to-platform mapping.
        category_children_playlists: Category-to-playlist mapping.
        platform_children_platforms: Platform-to-platform mapping.
        platform_children_playlists: Platform-to-playlist mapping.
        games_by_platform_norm: Normalized platform names mapped to game IDs.
        playlists_by_lb_id: Lookup of LaunchBox playlist metadata.

    Returns:
        list[FolderDict]: Playnite folder definitions rooted at the category.
    """

    def make_playlist_folder(lb_playlist_id: str) -> FolderDict | None:
        pl = playlists_by_lb_id.get(lb_playlist_id)
        if not pl:
            return None
        return {
            "Id": pl["Id"],
            "Name": pl["Name"],
            "GameIds": sorted(pl["GameIds"]),
        }

    def make_platform_folder(platform_name: str) -> FolderDict:
        folder: FolderDict = {
            "Id": str(uuid4()),
            "Name": platform_name,
        }

        nk = norm_key(platform_name)
        if nk in games_by_platform_norm:
            folder["GameIds"] = sorted(games_by_platform_norm[nk])

        children: list[FolderDict] = [
            *(
                pf
                for lb_pid in platform_children_playlists.get(platform_name, [])
                if (pf := make_playlist_folder(lb_pid)) is not None
            ),
            *(
                make_platform_folder(child_platform)
                for child_platform in platform_children_platforms.get(platform_name, [])
            ),
        ]

        if children:
            folder["Children"] = children

        return folder

    def make_category_folder(category_name: str) -> FolderDict:
        folder: FolderDict = {
            "Id": str(uuid4()),
            "Name": category_name,
        }
        children: list[FolderDict] = [
            *(
                make_category_folder(child_cat)
                for child_cat in sorted(
                    category_children_categories.get(category_name, [])
                )
            ),
            *(
                make_platform_folder(plat)
                for plat in sorted(category_children_platforms.get(category_name, []))
            ),
            *(
                pf
                for lb_pid in category_children_playlists.get(category_name, [])
                if (pf := make_playlist_folder(lb_pid)) is not None
            ),
        ]

        if children:
            folder["Children"] = children

        return folder

    root_folder = make_category_folder(root_category_name)
    return [root_folder]


# ----------------------------------------------------
# Main
# ----------------------------------------------------
def import_launchbox_exo() -> None:  # noqa: PLR0914, PLR0915
    """Parse LaunchBox XML and emit Playnite-friendly YAML outputs.

    Raises:
        SystemExit: If required LaunchBox inputs are missing.
        RuntimeError: If a worker fails while parsing a platform file.
    """
    playnite_games: list[GameDict] = []
    games_by_lb_id: dict[str, GameDict] = {}
    games_by_platform_norm: defaultdict[str, list[str]] = defaultdict(list)

    # --- parse all platform game lists ---
    if not PLATFORMS_DIR.exists():
        msg = f"Platforms dir not found: {PLATFORMS_DIR}"
        raise SystemExit(msg)

    platform_files = sorted(PLATFORMS_DIR.glob("*.xml"))
    if not platform_files:
        msg = f"No platform XML fil  es found in {PLATFORMS_DIR}"
        raise SystemExit(msg)

    platform_results: dict[Path, tuple[str, list[GameDict], dict[str, GameDict]]] = {}

    platform_bar = tqdm(total=len(platform_files), desc="Platforms", unit="platform")
    game_bar = tqdm(desc="Games", unit="game")

    ctx = get_context("spawn")
    manager = ctx.Manager()
    progress_queue: ProgressQueue = cast("ProgressQueue", manager.Queue())
    progress_thread = Thread(
        target=_drain_progress_queue,
        args=(progress_queue, game_bar),
        daemon=True,
    )
    progress_thread.start()
    worker_count = max(1, min(len(platform_files), os.cpu_count() or 1))

    try:
        for platform_file in platform_files:
            tqdm.write(f"Parsing {platform_file.name}…")

        worker_inputs = [(str(path), progress_queue) for path in platform_files]
        with ctx.Pool(processes=worker_count) as pool:
            for (
                platform_path,
                platform_name,
                games,
                lb_map,
            ) in pool.imap_unordered(_process_platform_file_worker, worker_inputs):
                path_obj = Path(platform_path)
                platform_results[path_obj] = (platform_name, games, lb_map)
                platform_bar.update(1)
                tqdm.write(
                    f"✔ Parsed {platform_name} ({len(games)} games, {len(lb_map)} IDs)"
                )
    except Exception as exc:  # pragma: no cover - defensive
        msg = f"Failed to process platform files: {exc}"
        raise RuntimeError(msg) from exc
    finally:
        progress_queue.put(None)
        progress_thread.join()
        manager.shutdown()
        platform_bar.close()
        game_bar.close()

    for platform_file in platform_files:
        platform_name, games, lb_map = platform_results[platform_file]
        playnite_games.extend(games)
        games_by_lb_id.update(lb_map)
        games_by_platform_norm[norm_key(platform_name)].extend(
            game["Id"] for game in games
        )

    # --- write games yaml ---
    tqdm.write(f"Saving games to {OUTPUT_GAMES} ({len(playnite_games)} entries)…")
    dump_yaml(playnite_games, OUTPUT_GAMES)

    # --- playlists ---
    playlists, playlists_by_lb_id = parse_playlists(games_by_lb_id)

    tqdm.write(f"Saving playlists to {OUTPUT_PLAYLISTS} ({len(playlists)} entries)…")
    dump_yaml(playlists, OUTPUT_PLAYLISTS)

    # --- parents.xml → folder tree ---
    (
        category_children_categories,
        category_children_platforms,
        category_children_playlists,
        platform_children_platforms,
        platform_children_playlists,
        root_categories,
    ) = parse_parents(playlists_by_lb_id)

    if (
        ROOT_CATEGORY_NAME not in root_categories
        and ROOT_CATEGORY_NAME not in category_children_categories
    ):
        tqdm.write(
            f"⚠ Root category '{ROOT_CATEGORY_NAME}' not found as a top-level category in Parents.xml"
        )
    else:
        folders = build_folder_tree_from_parents(
            ROOT_CATEGORY_NAME,
            category_children_categories,
            category_children_platforms,
            category_children_playlists,
            platform_children_platforms,
            platform_children_playlists,
            games_by_platform_norm,
            playlists_by_lb_id,
        )

        tqdm.write(
            f"Saving folders to {OUTPUT_FOLDERS} rooted at '{ROOT_CATEGORY_NAME}'"
        )
        dump_yaml(folders, OUTPUT_FOLDERS)

        tqdm.write(
            f"✔ Exported folder tree rooted at '{ROOT_CATEGORY_NAME}' to {OUTPUT_FOLDERS}"
        )

    tqdm.write(f"✔ Exported {len(playnite_games)} games → {OUTPUT_GAMES}")
    tqdm.write(f"✔ Exported {len(playlists)} playlists → {OUTPUT_PLAYLISTS}")


if __name__ == "__main__":
    import_launchbox_exo()
