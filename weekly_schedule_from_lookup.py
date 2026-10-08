"""Backward-compatible entry point for weekly schedules. Prefer: python -m table_gen schedule """

from table_gen.__main__ import main

if __name__ == "__main__":
    raise SystemExit(main(["schedule"]))
