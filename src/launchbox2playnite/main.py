from __future__ import annotations

import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from uuid import uuid4

import yaml

# ----------------------------------------------------
# CONFIG – change this to your eXo root
# ----------------------------------------------------
LAUNCHBOX_ROOT = Path(r"C:\Users\richa\Downloads\eXoWin9x")

PLATFORMS_DIR = LAUNCHBOX_ROOT / "Data" / "Platforms"
PLAYLIST_DIR = LAUNCHBOX_ROOT / "Data" / "Playlists"
PARENTS_FILE = LAUNCHBOX_ROOT / "Data" / "Parents.xml"

IMAGES_DIR = LAUNCHBOX_ROOT / "Images"  # adjust if needed
VIDEOS_DIR = LAUNCHBOX_ROOT / "Videos"
MANUALS_DIR = LAUNCHBOX_ROOT / "Manuals"

OUTPUT_GAMES = Path("playnite_import_games.yaml")
OUTPUT_PLAYLISTS = Path("playnite_import_playlists.yaml")
OUTPUT_FOLDERS = Path("playnite_import_folders.yaml")

ROOT_CATEGORY_NAME = "Computers"  # what to root the tree at


# ----------------------------------------------------
# Helpers
# ----------------------------------------------------
def norm_key(s: str) -> str:
    """Normalize names for matching (Windows 3x vs Windows 3.X)."""
    return "".join(ch for ch in s.lower() if ch.isalnum())


def normalize_title(title: str) -> str:
    return title.replace(":", "").replace("?", "").strip()


def find_first_image(folder: Path, title: str):
    if not folder.exists():
        return None
    base = normalize_title(title)
    for ext in ("png", "jpg", "jpeg", "webp"):
        for file in folder.glob(f"{base}*.{ext}"):
            return str(file)
    return None


def collect_images(title: str):
    media = {}
    media["cover"] = find_first_image(IMAGES_DIR / "Box - Front", title)
    media["icon"] = find_first_image(IMAGES_DIR / "Clear Logo", title)
    media["background"] = find_first_image(IMAGES_DIR / "Fanart - Background", title)

    screenshots = []
    base = normalize_title(title)
    screenshot_dir = IMAGES_DIR / "Screenshot - Gameplay"
    if screenshot_dir.exists():
        for ext in ("png", "jpg", "jpeg", "webp"):
            for f in screenshot_dir.glob(f"{base}*.{ext}"):
                screenshots.append(str(f))
    media["screenshots"] = screenshots
    return media


def collect_videos(title: str):
    vids = []
    if VIDEOS_DIR.exists():
        base = normalize_title(title)
        for ext in ("mp4", "avi", "mkv", "webm"):
            for f in VIDEOS_DIR.glob(f"{base}*.{ext}"):
                vids.append(str(f))
    return vids


def collect_manual(title: str):
    if not MANUALS_DIR.exists():
        return None
    base = normalize_title(title)
    for ext in ("pdf", "txt"):
        for f in MANUALS_DIR.glob(f"{base}*.{ext}"):
            return str(f)
    return None


# ----------------------------------------------------
# Game parsing
# ----------------------------------------------------
def parse_launchbox_game(game_elem, platform_name: str):
    def get(tag):
        elem = game_elem.find(tag)
        return elem.text.strip() if elem is not None and elem.text else None

    title = get("Title") or "Unknown Game"
    game_id = str(uuid4())

    images = collect_images(title)
    videos = collect_videos(title)
    manual = collect_manual(title)

    game = {
        "Id": game_id,
        "Name": title,
        "SortingName": get("SortTitle") or title,
        "Platform": platform_name,
        "ReleaseYear": get("ReleaseDate")[:4] if get("ReleaseDate") else None,
        "Developers": [get("Developer")] if get("Developer") else [],
        "Publishers": [get("Publisher")] if get("Publisher") else [],
        "Genres": [g.strip() for g in get("Genre").split(";")] if get("Genre") else [],
        "Description": get("Notes"),
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

    if images.get("cover"):
        game["Image"] = images["cover"]
    if images.get("background"):
        game["BackgroundImage"] = images["background"]
    if images.get("icon"):
        game["Icon"] = images["icon"]
    if images.get("screenshots"):
        game["Screenshots"] = images["screenshots"]
    if videos:
        game["Videos"] = videos
    if manual:
        game["Manual"] = manual

    # Strip empties
    return {k: v for k, v in game.items() if v not in ("", None, [], {})}


# ----------------------------------------------------
# Playlists: Data\Playlists\*.xml
# ----------------------------------------------------
def parse_playlists(games_by_lb_id):
    """
    Returns:
      playlists: list of Playnite playlist dicts
      playlists_by_lb_id: dict[LaunchBoxPlaylistGUID] -> playlist_dict
    """
    playlists = []
    playlists_by_lb_id = {}

    if not PLAYLIST_DIR.exists():
        return playlists, playlists_by_lb_id

    for file in PLAYLIST_DIR.glob("*.xml"):
        tree = ET.parse(file)
        root = tree.getroot()

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

        game_ids = []
        for pg in root.findall("PlaylistGame"):
            gid_elem = pg.find("GameId")
            if gid_elem is None or not gid_elem.text:
                continue
            lb_game_id = gid_elem.text.strip()
            g = games_by_lb_id.get(lb_game_id)
            if g:
                game_ids.append(g["Id"])

        playlist = {
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
def parse_parents(playlists_by_lb_id):
    """
    Parses Data/Parents.xml and returns relationship dicts.

    We care about:
      - categories
      - platforms
      - playlists

    Relationships:
      - category -> subcategory       (PlatformCategoryName + ParentPlatformCategoryName)
      - category -> platform          (PlatformName + ParentPlatformCategoryName)
      - category -> playlist          (PlaylistId + ParentPlatformCategoryName)
      - platform -> playlist          (PlaylistId + ParentPlatformName)
      - platform -> sub-platform      (PlatformName + ParentPlatformName)
    """

    category_children_categories = defaultdict(set)
    category_children_platforms = defaultdict(set)
    category_children_playlists = defaultdict(set)
    platform_children_platforms = defaultdict(set)
    platform_children_playlists = defaultdict(set)

    root_categories = set()

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

    tree = ET.parse(PARENTS_FILE)
    root = tree.getroot()

    for parent in root.findall("Parent"):
        platform_name = (parent.findtext("PlatformName") or "").strip()
        playlist_id = (parent.findtext("PlaylistId") or "").strip()
        cat_name = (parent.findtext("PlatformCategoryName") or "").strip()

        parent_platform = (parent.findtext("ParentPlatformName") or "").strip()
        parent_playlist_id = (parent.findtext("ParentPlaylistId") or "").strip()
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
        if playlist_id and parent_cat:
            if playlist_id in playlists_by_lb_id:
                category_children_playlists[parent_cat].add(playlist_id)

        # platform -> playlist
        if playlist_id and parent_platform:
            if playlist_id in playlists_by_lb_id:
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
def build_folder_tree_from_parents(
    root_category_name: str,
    category_children_categories,
    category_children_platforms,
    category_children_playlists,
    platform_children_platforms,
    platform_children_playlists,
    games_by_platform_norm,
    playlists_by_lb_id,
):
    def make_playlist_folder(lb_playlist_id: str):
        pl = playlists_by_lb_id.get(lb_playlist_id)
        if not pl:
            return None
        return {
            "Id": pl["Id"],
            "Name": pl["Name"],
            "GameIds": sorted(pl["GameIds"]),
        }

    def make_platform_folder(platform_name: str):
        folder = {
            "Id": str(uuid4()),
            "Name": platform_name,
        }

        # attach games by normalized platform name
        nk = norm_key(platform_name)
        if nk in games_by_platform_norm:
            folder["GameIds"] = sorted(games_by_platform_norm[nk])

        # child playlists
        children = []

        for lb_pid in platform_children_playlists.get(platform_name, []):
            pf = make_playlist_folder(lb_pid)
            if pf:
                children.append(pf)

        # child platforms (sub-platforms)
        for child_platform in platform_children_platforms.get(platform_name, []):
            children.append(make_platform_folder(child_platform))

        if children:
            folder["Children"] = children

        return folder

    def make_category_folder(category_name: str):
        folder = {
            "Id": str(uuid4()),
            "Name": category_name,
        }
        children = []

        # subcategories
        for child_cat in sorted(category_children_categories.get(category_name, [])):
            children.append(make_category_folder(child_cat))

        # platforms under this category
        for plat in sorted(category_children_platforms.get(category_name, [])):
            children.append(make_platform_folder(plat))

        # playlists directly under category
        for lb_pid in category_children_playlists.get(category_name, []):
            pf = make_playlist_folder(lb_pid)
            if pf:
                children.append(pf)

        if children:
            folder["Children"] = children

        return folder

    # Root is a category: "Computers"
    root_folder = make_category_folder(root_category_name)
    return [root_folder]


# ----------------------------------------------------
# Main
# ----------------------------------------------------
def import_launchbox_exo():
    playnite_games = []
    games_by_lb_id = {}
    games_by_platform_norm = defaultdict(list)

    # --- parse all platform game lists ---
    if not PLATFORMS_DIR.exists():
        raise SystemExit(f"Platforms dir not found: {PLATFORMS_DIR}")

    for platform_file in PLATFORMS_DIR.glob("*.xml"):
        platform_name = platform_file.stem  # e.g. "Windows 3x", "Windows 9x"

        tree = ET.parse(platform_file)
        root = tree.getroot()

        for game_elem in root.findall("Game"):
            game = parse_launchbox_game(game_elem, platform_name)

            # LaunchBox game ID
            lb_id_elem = game_elem.find("ID")
            if lb_id_elem is None:
                lb_id_elem = game_elem.find("Id")
            if lb_id_elem is not None and lb_id_elem.text:
                lb_id = lb_id_elem.text.strip()
                game["LaunchBoxId"] = lb_id
                games_by_lb_id[lb_id] = game

            playnite_games.append(game)
            games_by_platform_norm[norm_key(platform_name)].append(game["Id"])

    # --- write games yaml ---
    with open(OUTPUT_GAMES, "w", encoding="utf-8") as f:
        yaml.safe_dump(playnite_games, f, sort_keys=False, allow_unicode=True)

    # --- playlists ---
    playlists, playlists_by_lb_id = parse_playlists(games_by_lb_id)

    with open(OUTPUT_PLAYLISTS, "w", encoding="utf-8") as f:
        yaml.safe_dump(playlists, f, sort_keys=False, allow_unicode=True)

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
        print(
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

        with open(OUTPUT_FOLDERS, "w", encoding="utf-8") as f:
            yaml.safe_dump(folders, f, sort_keys=False, allow_unicode=True)

        print(
            f"✔ Exported folder tree rooted at '{ROOT_CATEGORY_NAME}' to {OUTPUT_FOLDERS}"
        )

    print(f"✔ Exported {len(playnite_games)} games → {OUTPUT_GAMES}")
    print(f"✔ Exported {len(playlists)} playlists → {OUTPUT_PLAYLISTS}")


if __name__ == "__main__":
    import_launchbox_exo()
