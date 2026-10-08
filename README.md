# Seminar table generator

Generates markdown **attendance tables** and **weekly open-lecture schedules** from the seminar Excel workbook (UWr-style sheets).

Works with **one or more seminars** (e.g. only *Narzędzia* this year, or both *Narzędzia* and *Agile*) via `config.yaml`.

## Layout

```
input/          # put the .xlsx workbook here
output/         # generated .md files
config.yaml     # seminars, dates, column indexes
table_gen/      # package
```

## Setup

```bash
python -m pip install -r requirements.txt
```

Place the workbook in `input/` and set `input_file` in `config.yaml` if the filename differs.

## Config

Edit `config.yaml`:

| Key | Purpose |
|---|---|
| `seminars[].enabled` | Turn a seminar on/off without changing the Excel |
| `schedule.start_date` / `end_date` | Weekly schedule window |
| `columns` | 0-based column indexes if the sheet layout changes |
| `columns.lookup_sheets.dates_value` | Date column on `*-daty` (default: F = 5) |
| `columns.lookup_sheets.titles_value` | Title column on `*-tematy` (default: E = 4) |

Dates should be real Excel dates (or numeric strings). If dates come from formulas, open and save the file in Excel once so cached values are available.

Sheet naming convention per seminar `Name`:

- `Name` — students / attendance
- `Name-daty` — presentation dates (lookup by email)
- `Name-tematy` — presentation titles (lookup by email)

## Usage

```bash
# Both attendance + weekly schedule for all enabled seminars
python -m table_gen

# Only attendance tables
python -m table_gen attendance

# Only weekly schedules
python -m table_gen schedule

# Custom config
python -m table_gen all -c path/to/config.yaml
```

Legacy wrappers (same as the commands above):

```bash
python table_generator.py
python weekly_schedule_from_lookup.py
```

## Output

Files land in `output/`:

- `{Seminar}_attendance.md` — index numbers × meeting dates (`+` = present)
- `{Seminar}_weekly_schedule.md` — week-by-week open lecture list

Disabled seminars, or seminars whose sheets are missing from the workbook, are skipped with a message.

## Example: one seminar this year

```yaml
seminars:
  - name: Narzędzia
    enabled: true
  - name: Agile
    enabled: false
```

## Example: both seminars

```yaml
seminars:
  - name: Narzędzia
    enabled: true
  - name: Agile
    enabled: true
```

Ensure the workbook contains the matching sheet triples for each enabled seminar.
