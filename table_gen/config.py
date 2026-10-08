from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.yaml"


@dataclass(frozen=True)
class StudentColumns:
    name: int
    surname: int
    index: int
    email: int
    english: int


@dataclass(frozen=True)
class LookupColumns:
    email: int
    dates_value: int
    titles_value: int


@dataclass(frozen=True)
class ScheduleConfig:
    start_date: str
    end_date: str
    week_interval_days: int


@dataclass(frozen=True)
class Seminar:
    name: str
    enabled: bool


@dataclass(frozen=True)
class Config:
    input_file: Path
    output_dir: Path
    seminars: list[Seminar]
    schedule: ScheduleConfig
    student_columns: StudentColumns
    lookup_columns: LookupColumns
    dates_suffix: str
    titles_suffix: str

    @property
    def enabled_seminars(self) -> list[Seminar]:
        return [s for s in self.seminars if s.enabled]


def _require(data: dict[str, Any], key: str) -> Any:
    if key not in data:
        raise KeyError(f"Missing required config key: {key}")
    return data[key]


def load_config(path: Path | None = None) -> Config:
    config_path = path or DEFAULT_CONFIG_PATH
    if not config_path.is_file():
        raise FileNotFoundError(f"Config not found: {config_path}")

    with config_path.open(encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    root = config_path.parent
    student = _require(_require(raw, "columns"), "student_sheet")
    lookup = _require(_require(raw, "columns"), "lookup_sheets")
    schedule = _require(raw, "schedule")
    suffixes = raw.get("suffixes") or {}

    seminars = [
        Seminar(name=str(item["name"]), enabled=bool(item.get("enabled", True)))
        for item in _require(raw, "seminars")
    ]

    return Config(
        input_file=(root / str(_require(raw, "input_file"))).resolve(),
        output_dir=(root / str(_require(raw, "output_dir"))).resolve(),
        seminars=seminars,
        schedule=ScheduleConfig(
            start_date=str(_require(schedule, "start_date")),
            end_date=str(_require(schedule, "end_date")),
            week_interval_days=int(schedule.get("week_interval_days", 7)),
        ),
        student_columns=StudentColumns(
            name=int(student["name"]),
            surname=int(student["surname"]),
            index=int(student["index"]),
            email=int(student["email"]),
            english=int(student["english"]),
        ),
        lookup_columns=LookupColumns(
            email=int(lookup["email"]),
            dates_value=int(lookup.get("dates_value", lookup.get("value", 4))),
            titles_value=int(lookup.get("titles_value", lookup.get("value", 4))),
        ),
        dates_suffix=str(suffixes.get("dates", "-daty")),
        titles_suffix=str(suffixes.get("titles", "-tematy")),
    )
