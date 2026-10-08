"""Backward-compatible entry point for attendance tables. Prefer: python -m table_gen attendance """

from table_gen.__main__ import main

if __name__ == "__main__":
    raise SystemExit(main(["attendance"]))
