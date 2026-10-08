from __future__ import annotations

import argparse
import sys
from pathlib import Path

from table_gen.attendance import write_attendance_tables
from table_gen.config import DEFAULT_CONFIG_PATH, load_config
from table_gen.schedule import write_weekly_schedules


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Generate seminar attendance tables and weekly open-lecture schedules "
            "from an Excel workbook."
        )
    )
    parser.add_argument(
        "command",
        nargs="?",
        default="all",
        choices=("attendance", "schedule", "all"),
        help="What to generate (default: all)",
    )
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help=f"Path to config.yaml (default: {DEFAULT_CONFIG_PATH})",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        config = load_config(args.config)
    except (OSError, KeyError, ValueError) as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 1

    enabled = [s.name for s in config.enabled_seminars]
    if not enabled:
        print("No seminars enabled in config.", file=sys.stderr)
        return 1

    print(f"Input: {config.input_file}")
    print(f"Output: {config.output_dir}")
    print(f"Seminars: {', '.join(enabled)}")

    try:
        if args.command in ("attendance", "all"):
            write_attendance_tables(config)
        if args.command in ("schedule", "all"):
            write_weekly_schedules(config)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
