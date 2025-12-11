from __future__ import annotations

import argparse
from collections import defaultdict
from collections.abc import Iterable
from pathlib import Path
from typing import Any

import yaml


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract duplicate records from a Playnite YAML export based on one or "
            "more string fields."
        )
    )
    parser.add_argument(
        "--input",
        default="playnite_import_games.yaml",
        help="Path to the source YAML file (defaults to playnite_import_games.yaml).",
    )
    parser.add_argument(
        "--output",
        default="duplicate_games.yaml",
        help="Destination file for duplicate groups (defaults to duplicate_games.yaml).",
    )
    parser.add_argument(
        "--fields",
        nargs="+",
        default=["Title", "Name", "SortingName"],
        help=(
            "Ordered list of fields to use when building the duplicate key. "
            "The first populated field wins."
        ),
    )
    return parser.parse_args()


def load_yaml(path: Path) -> list[dict[str, Any]]:
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise SystemExit(f"Input file not found: {path}") from exc
    except yaml.YAMLError as exc:
        raise SystemExit(f"Failed to parse YAML: {exc}") from exc

    if not isinstance(data, list):
        raise SystemExit("Expected the YAML root to be a list of game records")

    typed_data: list[dict[str, Any]] = []
    for idx, item in enumerate(data, start=1):
        if not isinstance(item, dict):
            raise SystemExit(f"Record #{idx} is not a mapping; aborting")
        typed_data.append(item)
    return typed_data


def pick_value(record: dict[str, Any], fields: Iterable[str]) -> str | None:
    for field in fields:
        raw = record.get(field)
        if raw is None:
            continue
        if isinstance(raw, str):
            cleaned = raw.strip()
            if cleaned:
                return cleaned
            continue
        return str(raw)
    return None


def normalize_key(value: str) -> str:
    return value.casefold()


def collect_duplicates(
    records: list[dict[str, Any]], fields: Iterable[str]
) -> list[dict[str, Any]]:
    grouped: dict[str, list[tuple[str, dict[str, Any]]]] = defaultdict(list)

    for record in records:
        value = pick_value(record, fields)
        if not value:
            continue
        grouped[normalize_key(value)].append((value, record))

    duplicates: list[dict[str, Any]] = []
    for key in sorted(grouped):
        entries = grouped[key]
        if len(entries) <= 1:
            continue
        first_value = entries[0][0]
        duplicates.append({
            "duplicate_value": first_value,
            "count": len(entries),
            "items": [entry for _, entry in entries],
        })
    return duplicates


def dump_yaml(path: Path, data: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(data, handle, sort_keys=False, allow_unicode=True)


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    records = load_yaml(input_path)
    duplicates = collect_duplicates(records, args.fields)
    dump_yaml(output_path, duplicates)

    print(
        f"Found {len(duplicates)} duplicate group(s) using fields {args.fields}. "
        f"Details written to {output_path}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
