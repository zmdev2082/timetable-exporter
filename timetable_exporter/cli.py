import argparse
import sys
import os
import json
import traceback
import re
from .ical_generator import IcalGenerator
from .user_extensions import TimetableAccessor
from .argparse_utils import LoadExcelAction, LoadJSONAction, ValidateDirectoryAction  # Import custom actions
from openpyxl import Workbook

from .week_view_exporter import build_week_view_workbook, render_week_view_worksheet


def _root_data_path(*parts: str) -> str:
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
    return os.path.join(base, *parts)


def _default_mapping_path() -> str:
    return _root_data_path("presets", "mapping.template.json")


def _preset_roots() -> list[str]:
    """Search paths for presets.

    - data/presets: shipped, public templates
    - data/proprietary/presets: optional local-only presets (gitignored)
    """
    roots = [_root_data_path("presets"), _root_data_path("proprietary", "presets")]
    return [p for p in roots if os.path.isdir(p)]


def _walk_preset_files(base: str):
    """Yield (root, filename) for preset files, skipping public 'proprietary' subtrees."""
    for root, dirs, files in os.walk(base):
        # Never ship or expose presets from a public 'proprietary' subtree.
        if os.path.basename(base) == "presets":
            dirs[:] = [d for d in dirs if d.lower() != "proprietary"]
        for name in files:
            yield root, name


def _load_preset_json(filename: str) -> dict:
    bases = _preset_roots()

    # Allow subfolder-qualified names like "team/foo.mapping.json" relative to any preset root
    for base in bases:
        candidate = os.path.join(base, filename)
        if os.path.isfile(candidate):
            with open(candidate, "r", encoding="utf-8") as f:
                return json.load(f)

    matches: list[tuple[str, str]] = []  # (base, abs_path)
    for base in bases:
        for root, name in _walk_preset_files(base):
            if name == filename:
                matches.append((base, os.path.join(root, name)))

    if len(matches) == 1:
        with open(matches[0][1], "r", encoding="utf-8") as f:
            return json.load(f)
    if len(matches) > 1:
        rels = [os.path.relpath(p, base).replace("\\", "/") for base, p in matches]
        raise ValueError(f"Preset name is ambiguous: {filename}. Matches: {rels}")

    raise FileNotFoundError(f"Preset not found: {filename}")


def _list_presets() -> tuple[list[str], list[str]]:
    mappings: list[str] = []
    filters: list[str] = []
    for base in _preset_roots():
        for root, name in _walk_preset_files(base):
            rel = os.path.relpath(os.path.join(root, name), base)
            rel = rel.replace("\\", "/")
            if name.endswith(".mapping.json"):
                mappings.append(rel[:-len(".mapping.json")])
            elif name.endswith(".filters.json"):
                filters.append(rel[:-len(".filters.json")])
    return sorted(mappings), sorted(filters)


_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


def _safe_sheet_title(title: str, used: set[str]) -> str:
    # Excel: max 31 chars, cannot contain: \ / * ? : [ ]
    base = _INVALID_SHEET_CHARS.sub("_", (title or "").strip())
    base = base[:31] if base else "Sheet"

    candidate = base
    i = 2
    while candidate in used:
        suffix = f"_{i}"
        candidate = (base[: max(0, 31 - len(suffix))] + suffix) or f"Sheet{i}"
        i += 1
    used.add(candidate)
    return candidate


def setup_argparse():
    parser = argparse.ArgumentParser(description='Generate an iCal file from a timetabling Excel sheet.')
    parser.add_argument('excel_file', nargs='?', type=str, help='The path of the local Excel file.')

    parser.add_argument('--list-presets', action='store_true', help='List built-in mapping/filter presets and exit.')

    parser.add_argument('filters_file', nargs='?', type=str, action=LoadJSONAction, help='Path to the JSON file containing filters.')
    parser.add_argument('--filters', dest='filters_file', type=str, action=LoadJSONAction, help='Path to the JSON file containing filters.')
    parser.add_argument('--filters-preset', type=str, help='Built-in filters preset name (see --list-presets).')

    mapping_group = parser.add_mutually_exclusive_group()
    mapping_group.add_argument('--mapping', type=str, action=LoadJSONAction, help='Path to the JSON file containing mapping/config. (default: bundled mapping.json)')
    mapping_group.add_argument('--mapping-preset', type=str, help='Built-in mapping preset name (see --list-presets).')

    parser.add_argument('--exact', action='store_true', help='Use exact matching for filters.')
    parser.add_argument('--timezone', type=str, default='Australia/Sydney', help='Timezone for the events (default: Australia/Sydney).')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode.')
    parser.add_argument('--output_dir', type=str, default=None, action=ValidateDirectoryAction, help='Directory to save output iCal files (defaults to filters output_dir if present).')
    parser.add_argument('--week-view', action='store_true', help='Generate a weekly view workbook using defaults (or values from filters JSON).')
    parser.add_argument('--week-view-output', type=str, default=None, help='Weekly view output path. If filters define multiple calendars: .xlsx => one workbook with one sheet per calendar; otherwise treat as a directory and write one .xlsx per calendar.')
    parser.add_argument('--week-view-template', type=str, action=LoadJSONAction, help='Path to weekly view template JSON (defaults to data/presets/week_view.template.json).')
    parser.add_argument('--week-view-calendar', type=str, default=None, help='When filters define multiple calendars, generate weekly view for a single calendar filename.')
    parser.add_argument('--skip-extensions', type=str, default=None, help='Comma-separated list of user extension functions to skip (e.g. expand_dates).')

    return parser

def timetable_exporter():
    parser = setup_argparse()
    try:
        args = parser.parse_args()
        if args.list_presets:
            mapping_names, filter_names = _list_presets()
            print("Mapping presets:")
            for name in mapping_names:
                print(f"  - {name}")
            print("Filters presets:")
            for name in filter_names:
                print(f"  - {name}")
            sys.exit(0)

        if args.excel_file is None:
            parser.error("Missing excel_file. Provide a path to the Excel file.")

        # Load the Excel file into a DataFrame
        LoadExcelAction(option_strings=['excel_file'], dest='excel_file')(parser, args, args.excel_file)

        if getattr(args, "mapping_preset", None):
            args.mapping = _load_preset_json(f"{args.mapping_preset}.mapping.json")
        elif args.mapping is None:
            LoadJSONAction(option_strings=['--mapping'], dest='mapping')(parser, args, _default_mapping_path())

        if args.filters_preset and args.filters_file is not None:
            parser.error("Provide either filters JSON file OR --filters-preset, not both.")

        if args.filters_preset:
            args.filters_file = _load_preset_json(f"{args.filters_preset}.filters.json")

        # Filters are optional: if none provided, we generate a single calendar for the whole dataset.
        if args.filters_file is None:
            args.filters_file = {}

        df = args.excel_file

        filters_payload = args.filters_file

        if args.output_dir is None:
            output_dir = filters_payload.get("output_dir") or "."
            ValidateDirectoryAction(option_strings=['--output_dir'], dest='output_dir')(parser, args, output_dir)

        # Allow week view defaults to come from filters JSON
        if args.week_view_output is None:
            args.week_view_output = filters_payload.get("week_view_output")
        if args.week_view_template is None and filters_payload.get("week_view_template"):
            LoadJSONAction(option_strings=['--week-view-template'], dest='week_view_template')(parser, args, filters_payload.get("week_view_template"))

        # Apply global filters (optional)
        global_filters = filters_payload.get("global_filters")
        if global_filters:
            df = df.timetable.filter(global_filters, exact_match=args.exact)

        # Check if the DataFrame is empty after filtering
        if df.empty:
            print("No data found after applying global filters. Exiting.")
            sys.exit(0)
        
        # Preprocess the DataFrame using custom user extensions
        # This assumes that the user extensions are properly registered
        # which should be automatic if the file has been added to the user_extensions directory
        # and that the DataFrame has a 'timetable' accessor
        user_extensions = args.mapping.get("user_extensions", {})
        skip_extensions = set()
        if args.skip_extensions:
            skip_extensions = {name.strip() for name in args.skip_extensions.split(',') if name.strip()}

        for func, parameters in user_extensions.items():
            if func in skip_extensions:
                continue
            for call in parameters:
                df = getattr(df.timetable, func)(*call.get("args", []), **(call.get("kwargs") or {}))

        # Load the configuration file
        columns = args.mapping["columns"]
        company = args.mapping.get("company", "timetable-exporter")



        # Initialize the iCal generator
        ical_generator = IcalGenerator(columns, timezone=args.timezone)

        # Optional weekly view export
        if args.week_view or args.week_view_output:
            week_view_cfg = args.week_view_template
            if week_view_cfg is None:
                week_view_cfg = _load_preset_json("week_view.template.json")

            calendars = filters_payload.get("calendars") or []
            output_path = args.week_view_output
            if output_path is None:
                output_path = os.path.join(args.output_dir, "week_view.xlsx")

            if calendars:
                if args.week_view_calendar:
                    selected = next((c for c in calendars if c.get("filename") == args.week_view_calendar), None)
                    if selected is None:
                        raise ValueError(f"Calendar not found for weekly view: {args.week_view_calendar}")
                    week_view_df = df.timetable.filter(selected["filter"], exact_match=args.exact)
                    wb = build_week_view_workbook(week_view_df, week_view_cfg)
                    wb.save(output_path)
                elif len(calendars) == 1:
                    selected = calendars[0]
                    week_view_df = df.timetable.filter(selected["filter"], exact_match=args.exact)
                    wb = build_week_view_workbook(week_view_df, week_view_cfg)
                    wb.save(output_path)
                else:
                    # Multiple calendars:
                    # - if output_path ends with .xlsx => one workbook, one sheet per calendar
                    # - otherwise treat output_path as a directory and write one workbook per calendar
                    if output_path.lower().endswith(".xlsx"):
                        wb = Workbook()
                        # Remove default sheet
                        wb.remove(wb.active)
                        used_titles: set[str] = set()
                        for calendar in calendars:
                            week_view_df = df.timetable.filter(calendar["filter"], exact_match=args.exact)
                            sheet_name = _safe_sheet_title(calendar.get("filename") or "Calendar", used_titles)
                            ws = wb.create_sheet(title=sheet_name)
                            render_week_view_worksheet(ws, week_view_df, week_view_cfg)
                        wb.save(output_path)
                    else:
                        ValidateDirectoryAction(option_strings=['--week-view-output'], dest='week_view_output')(parser, args, output_path)
                        for calendar in calendars:
                            week_view_df = df.timetable.filter(calendar["filter"], exact_match=args.exact)
                            wb = build_week_view_workbook(week_view_df, week_view_cfg)
                            out_file = os.path.join(output_path, f"{calendar['filename']}.xlsx")
                            wb.save(out_file)
            else:
                wb = build_week_view_workbook(df, week_view_cfg)
                wb.save(output_path)

        # Process each calendar in the filters
        calendars = filters_payload.get("calendars")

        if not calendars:
            output_file = os.path.join(args.output_dir, "timetable.ics")
            timetable_data = df.to_dict(orient='records')
            cal = ical_generator.generate_ical(timetable_data, company)
            with open(output_file, 'wb') as f:
                f.write(cal.to_ical())
            return

        for calendar in calendars:
            output_file = os.path.join(args.output_dir, f"{calendar['filename']}.ics")
            calendar_filters = calendar["filter"]
            filtered_df = df.timetable.filter(calendar_filters, exact_match=args.exact)
            timetable_data = filtered_df.to_dict(orient='records')

            cal = ical_generator.generate_ical(timetable_data, company)
            with open(output_file, 'wb') as f:
                f.write(cal.to_ical())

    except Exception as e:
        debug_enabled = bool(os.getenv("TIMETABLE_EXPORTER_DEBUG"))
        if not debug_enabled:
            try:
                debug_enabled = bool(getattr(args, 'debug', False))
            except Exception:
                debug_enabled = False

        if debug_enabled:
            traceback.print_exc()
        print(f"An unexpected error occurred: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    timetable_exporter()

