# timetable-exporter

Convert spreadsheet event data into iCalendar (\*.ics) files, and optionally generate a printable weekly view (\*.xlsx).

## Quickstart (minimal flags)

If you already have your JSON configs:

`timetable-exporter my.xlsx --mapping my.mapping.json --filters my.filters.json`

Add weekly view with one extra flag:

`timetable-exporter my.xlsx --mapping my.mapping.json --filters my.filters.json --week-view`

## Install (end users)

`pipx install timetable-exporter`

## Install (development)

1) Create and activate a virtualenv

- PowerShell:
  - `python -m venv .venv`
  - `.\.venv\Scripts\python -m pip install -U pip`

2) Install this project (editable)

`.\.venv\Scripts\python -m pip install -e ".[dev]"`

Run via the module (always works inside a venv):

`python -m timetable_exporter --help`

## Usage

### 1) Single calendar output (one .ics)

`timetable-exporter my.xlsx --mapping my.mapping.json --output_dir output`

### 2) Split into multiple calendars (filters JSON)

`timetable-exporter my.xlsx --mapping my.mapping.json --filters my.filters.json`

If `output_dir` is present in the filters JSON, it will be used by default (you can still override with `--output_dir`).

### 3) Weekly view export (Excel)

Explicit output + template:

`timetable-exporter my.xlsx --mapping my.mapping.json --filters my.filters.json --week-view-output output/week_view.xlsx --week-view-template data/presets/week_view.template.json`

Or use defaults:

`timetable-exporter my.xlsx --mapping my.mapping.json --filters my.filters.json --week-view`

#### Multi-calendar weekly view behavior

If your filters JSON contains multiple calendars:

- `--week-view-output something.xlsx` (or the default output path) produces **one workbook** with **one sheet per calendar**.
- `--week-view-calendar NAME` limits the weekly view to a single calendar (one sheet).
- `--week-view-output some_folder` writes **one workbook per calendar** into that folder.

### Defaults and override order (important)

The CLI is designed so you can put defaults into JSON and only override when needed.

- `output_dir`
    1) `--output_dir`
    2) `filters.json` → `output_dir`
    3) `.`

- Weekly view generation
    - Enabled when you pass `--week-view` or `--week-view-output`.

- `week_view_output`
    1) `--week-view-output`
    2) `filters.json` → `week_view_output`
    3) `<output_dir>/week_view.xlsx`

- `week_view_template`
    1) `--week-view-template`
    2) `filters.json` → `week_view_template`
    3) `data/presets/week_view.template.json`

### Common flags (when you actually need them)

- `--exact`: use exact (not contains) matching in filters.
- `--timezone`: set event timezone when your data is timezone-naive.
- `--skip-extensions expand_dates,combine_date_time`: debug a mapping by temporarily bypassing one or more user extensions.
- `--week-view-calendar NAME`: when you have multiple calendars but only want one weekly view sheet.

## Presets (JSON)

Presets and templates live under data/presets and can be grouped in subfolders.

- List available presets:
   ```
   timetable-exporter --list-presets
   ```

Templates live in [data/presets](data/presets). For local-only presets, use [data/proprietary](data/proprietary) (gitignored).

## Weekly view export (printable Excel)

Generate a single-week grid from your data using a template JSON:

```
timetable-exporter /path/to/your/timetable.xlsx \
   --week-view-output /path/to/output/weekly_view.xlsx \
   --week-view-template /path/to/week_view.template.json
```

You can use the default template at [data/presets/week_view.template.json](data/presets/week_view.template.json) as a starting point.

## Command-Line Options

You can use various command-line options to customize the output. For example:

```text
options:
    -h, --help            show this help message and exit
    --list-presets        List built-in mapping/filter presets and exit.
    --mapping MAPPING     Path to the JSON file containing mapping/config.
    --mapping-preset MAPPING_PRESET Built-in mapping preset name (see --list-presets).
    --filters FILTERS_FILE Path to the JSON file containing filters.
    --filters-preset FILTERS_PRESET Built-in filters preset name (see --list-presets).
    --exact               Use exact matching for filters.
    --timezone TIMEZONE   Timezone for the events (default: Australia/Sydney).
    --output_dir OUTPUT_DIR
    --week-view-output WEEK_VIEW_OUTPUT
    --week-view-template WEEK_VIEW_TEMPLATE
    --week-view-calendar WEEK_VIEW_CALENDAR
    --skip-extensions SKIP_EXTENSIONS  Comma-separated list of user extension functions to skip.
```

## Custom user extensions

You can add custom DataFrame helpers by dropping a Python file into [timetable_exporter/user_extensions](timetable_exporter/user_extensions). Every public function in those files is automatically attached to the `timetable` accessor (e.g., `df.timetable.my_func(...)`).

How it works:
- The package auto-imports every `.py` file in that folder at runtime.
- Every top-level name that doesn’t start with `__` is added to `TimetableAccessor`.


## Data layout

See [data/README.md](data/README.md) for the recommended structure.

## Config JSON Structure
The config JSON file maps columns in your Excel file to the fields required for generating the iCal file.

```json
{
    "columns": {
   "summary": "Summary Column",
   "dtstart": "Start DateTime Column",
   "duration": "Duration Column",
   "location": "Location Column",
   "description": "Description Column"
    },
    "additional_parsing": {
        "summary":
        {
            "func": "replace",
            "args": ["/", ", "],
            "kwargs": {}
        },
        "duration":
        {
            "func": "__truediv__",
            "args": [2],
            "kwargs": {}
        },
        "attendees": {
            "func": "split",
            "args": [";"],
            "kwargs": {}
        },
        "categories": {
            "func": "split",
            "args": [";"],
            "kwargs": {}
        }

    }
}
```
### Explanation of columns

The columns section maps the named Excel columns to the properties of the iCal event. Below is an explanation of each column, including whether it is **required**, *optional*, or **conditionally required**. 

- **summary** (**Required**): This column maps to the event name or summary in the iCal file. It should contain a brief description or title of the event.
- **start** (**Required if `date_start` and `time_start` are not provided**): This column maps to the event start datetime. It should contain both the date and the time, such as (2025-03-18 13:30:00).
- **end** (**Required if `duration` is not provided**): This column maps to the event start datetime. It should contain both the date and the time, such as (2025-03-18 14:00:00).
- **date_start** (**Required if `start` is not provided**): This column maps to the start date of the event. It should contain the date when the event begins.
- **time_start** (**Required if `start` is not provided**): This column maps to the start time of the event. It should contain the time when the event begins.
- **duration** (**Required if `end` is not provided**): This column maps to the duration of the event. It should contain the length of time the event lasts.
- **location** (*Optional*): This column maps to the location of the event. It should contain the place where the event will be held.
- **description** (*Optional*): This column maps to the detailed description of the event. 

- **attendees** (*Optional*): This column maps to the attendees of the event, with splitting on the delimiter specified in the `additional_parsing` section below.
- **categories** (*Optional*): This column maps to the categories of the event. It should contain a list of categories, with splitting on the delimiter specified in the `additional_parsing` section below.

### Additional Parsing
The `additional_parsing` section allows for limited preprocessing of data before exporting. Each key in the `additional_parsing` dictionary corresponds to a column name and specifies a method of the object to be applied to the data in that column. Below is an explanation of each key and its associated function:

- **summary**: Uses the `replace` function to replace all occurrences of "/" with ", ".
  ```json
  "summary": {
      "func": "replace",
      "args": ["/", ", "],
      "kwargs": {}
  }
   ```
- **attendees**: Uses the `split`
 function to split the string of categories by ";" delimiter.
   ```json
   "attendees": {
      "func": "split",
      "args": [";"],
      "kwargs": {}
   }
   ```
- **categories**: Uses the `split` function to split the string of categories by ";" delimiter.
   ```json
   "categories": {
    "func": "split",
    "args": [";"],
    "kwargs": {}
   }
   ```
## Contributing
Feel free to submit issues or pull requests if you have suggestions or improvements for the project.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.