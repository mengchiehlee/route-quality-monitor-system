from __future__ import annotations

import argparse
import csv
import copy
import io
import json
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell, ILLEGAL_CHARACTERS_RE, MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
DEFAULT_SOURCE_DIR = PROJECT_ROOT / "input"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"
DEFAULT_QUALITY_WORKBOOK = "质量.xlsx"
DEFAULT_CENTER_TIMING_WORKBOOK = "中心签出时效.csv"
DEFAULT_SIGNIN_EXCLUDED_WORKBOOK = "中心签入时效 (剔除).csv"
DEFAULT_SIGNIN_INCLUDED_WORKBOOK = "中心签入时效 (不剔除).csv"
TOTAL_LABEL = "\u5408\u8ba1"

MONITOR_COLUMNS = {
    "date": 1,
    "area": 2,
    "dsp": 3,
    "route": 4,
    "dispatch_qty": 5,
    "delivered_qty": 6,
    "pod_2400": 7,
    "ascan_dscan_72h_weekend_holiday": 8,
    "ascan_dscan_72h_business_day": 9,
    "fake_sign_qty": 10,
    "fake_sign_rate": 11,
    "complaint_qty": 12,
    "complaint_rate": 13,
    "suspected_lost_qty": 14,
    "tracking_break_qty": 15,
}

BLUE_METRICS = [
    "dispatch_qty",
    "delivered_qty",
    "fake_sign_qty",
    "fake_sign_rate",
    "complaint_qty",
    "complaint_rate",
    "suspected_lost_qty",
    "tracking_break_qty",
]

GREEN_METRICS = [
    "pod_2400",
    "ascan_dscan_72h_weekend_holiday",
    "ascan_dscan_72h_business_day",
]

QUALITY_METRICS = [
    "delivered_qty",
    "fake_sign_qty",
    "fake_sign_rate",
    "complaint_qty",
    "complaint_rate",
    "suspected_lost_qty",
    "tracking_break_qty",
]

PERCENT_COLUMNS = [
    MONITOR_COLUMNS["pod_2400"],
    MONITOR_COLUMNS["ascan_dscan_72h_weekend_holiday"],
    MONITOR_COLUMNS["ascan_dscan_72h_business_day"],
    MONITOR_COLUMNS["fake_sign_rate"],
    MONITOR_COLUMNS["complaint_rate"],
]

TOTAL_ROW_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
TOTAL_ROW_FONT = Font(name="\u5b8b\u4f53", size=10, bold=True, color="000000")
GROUP_FONT = Font(name="\u5b8b\u4f53", size=10, bold=True, color="000000")
GROUP_ALIGNMENT = Alignment(horizontal="center", vertical="center")
GREEN_PASS_FONT = Font(name="\u5b8b\u4f53", size=10, color="000000")
GREEN_FAIL_FONT = Font(name="\u5b8b\u4f53", size=10, color="FF0000")
GREEN_TOTAL_PASS_FONT = Font(name="\u5b8b\u4f53", size=10, bold=True, color="000000")
GREEN_TOTAL_FAIL_FONT = Font(name="\u5b8b\u4f53", size=10, bold=True, color="FF0000")
QUALITY_PASS_FONT = Font(name="\u5b8b\u4f53", size=10, color="000000")
QUALITY_FAIL_FONT = Font(name="\u5b8b\u4f53", size=10, color="FF0000")

@dataclass(frozen=True)
class Project:
    area: str
    dsp: str
    routes: list[str]
    start_date: str | None = None


def active_projects(projects: list[Project], target_date: str) -> list[Project]:
    target_dt = parse_business_date(target_date)
    active: list[Project] = []
    for project in projects:
        if project.start_date:
            if parse_business_date(project.start_date) > target_dt:
                continue
        active.append(project)
    return merge_projects(active)


def merge_projects(projects: list[Project]) -> list[Project]:
    merged: list[Project] = []
    project_by_key: dict[tuple[str, str], int] = {}
    for project in projects:
        key = (normalize(project.area), normalize(project.dsp))
        existing_index = project_by_key.get(key)
        if existing_index is None:
            project_by_key[key] = len(merged)
            merged.append(project)
            continue

        existing = merged[existing_index]
        routes = list(existing.routes)
        for route in project.routes:
            if route not in routes:
                routes.append(route)
        merged[existing_index] = Project(
            area=existing.area,
            dsp=existing.dsp,
            routes=routes,
            start_date=existing.start_date,
        )
    return merged


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update DSP monitor workbook.")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"), help="Path to JSON config.")
    parser.add_argument(
        "--date",
        help="Business date to insert, e.g. 2026/5/19. Defaults to yesterday.",
    )
    parser.add_argument(
        "--allow-sunday",
        action="store_true",
        help="Allow inserting a Sunday date. By default Sundays are skipped.",
    )
    return parser.parse_args()


def load_config(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def normalize(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{value.year}/{value.month}/{value.day}"
    text = str(value).strip()
    if not text:
        return ""
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(text, fmt)
            return f"{dt.year}/{dt.month}/{dt.day}"
        except ValueError:
            pass
    return text.replace(" ", "")


def same_date(left: Any, right: Any) -> bool:
    return normalize(left) == normalize(right)


def parse_business_date(value: str) -> datetime:
    normalized = normalize(value)
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            pass
    raise ValueError(f"Invalid date: {value}. Use YYYY/M/D, e.g. 2026/5/19.")


def default_business_date() -> str:
    dt = datetime.today() - timedelta(days=1)
    return f"{dt.year}/{dt.month}/{dt.day}"


def default_output_path(workbook_path: Path) -> Path:
    return next_output_path(DEFAULT_OUTPUT_DIR)


def next_output_path(folder: Path) -> Path:
    today = datetime.today().strftime("%Y%m%d")
    candidate = folder / f"Update{today}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = folder / f"Update{today}_{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def latest_source_directory(input_dir: Path) -> Path:
    if not input_dir.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_dir}")
    candidates = []
    for child in input_dir.iterdir():
        if not child.is_dir() or not re.fullmatch(r"[0-9]{8}", child.name):
            continue
        try:
            datetime.strptime(child.name, "%Y%m%d")
        except ValueError:
            continue
        candidates.append(child)
    if not candidates:
        raise FileNotFoundError(f"输入目录中没有有效的 YYYYMMDD 日期文件夹: {input_dir}")
    return max(candidates, key=lambda path: path.name)


def resolve_daily_paths(args: argparse.Namespace) -> tuple[Path, Path, Path, Path, Path]:
    source_dir = latest_source_directory(DEFAULT_SOURCE_DIR)
    sources = tuple(source_dir / name for name in (
        DEFAULT_QUALITY_WORKBOOK,
        DEFAULT_CENTER_TIMING_WORKBOOK,
        DEFAULT_SIGNIN_EXCLUDED_WORKBOOK,
        DEFAULT_SIGNIN_INCLUDED_WORKBOOK,
    ))
    missing = [path.name for path in sources if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            f"最新日期文件夹 {source_dir} 缺少必需数据源（不会改用旧日期）:\n"
            + "\n".join(f"- {name}" for name in missing)
        )
    print(f"Input: {source_dir}")
    return (*sources, next_output_path(DEFAULT_OUTPUT_DIR))


def save_output_workbook(wb: Workbook, output_path: Path) -> Path:
    # Exclusive creation also prevents overwrites when two runs finish together.
    output_dir = DEFAULT_OUTPUT_DIR.resolve()
    if output_dir == DEFAULT_SOURCE_DIR.resolve() or DEFAULT_SOURCE_DIR.resolve() in output_dir.parents:
        raise ValueError("output 不能指向 input 或其子目录。")
    if output_path.parent.resolve() != output_dir:
        raise ValueError("结果只能保存到项目根目录的 output 文件夹。")
    output_dir.mkdir(parents=True, exist_ok=True)
    while True:
        try:
            stream = output_path.open("xb")
        except FileExistsError:
            output_path = next_output_path(output_dir)
            continue
        with stream:
            wb.save(stream)
        return output_path


def previous_service_date(target_date: str, holidays: list[str]) -> str:
    holiday_set = {normalize(item) for item in holidays}
    dt = parse_business_date(target_date) - timedelta(days=1)
    while dt.weekday() == 6 or normalize(dt) in holiday_set:
        dt -= timedelta(days=1)
    return f"{dt.year}/{dt.month}/{dt.day}"


def service_days_before(target_date: str, days: int, holidays: list[str]) -> str:
    holiday_set = {normalize(item) for item in holidays}
    dt = parse_business_date(target_date)
    moved = 0
    while moved < days:
        dt -= timedelta(days=1)
        if dt.weekday() == 6 or normalize(dt) in holiday_set:
            continue
        moved += 1
    return f"{dt.year}/{dt.month}/{dt.day}"


def calendar_days_before(target_date: str, days: int) -> str:
    dt = parse_business_date(target_date) - timedelta(days=days)
    return f"{dt.year}/{dt.month}/{dt.day}"


def latest_source_date(ws: Worksheet, header_map: dict[str, int], source_header_row: int) -> str:
    if "date" not in header_map:
        raise ValueError("Source sheet missing required date header.")

    latest: datetime | None = None
    current_date: Any = None
    for row in range(source_header_row + 1, ws.max_row + 1):
        row_date = ws.cell(row, header_map["date"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if current_date in (None, ""):
            continue
        try:
            parsed = parse_business_date(normalize(current_date))
        except ValueError:
            continue
        if parsed.weekday() == 6:
            continue
        if latest is None or parsed > latest:
            latest = parsed

    if latest is None:
        raise ValueError("No usable source date found in DSP质量总览.")
    return f"{latest.year}/{latest.month}/{latest.day}"


def source_dates(ws: Worksheet, header_map: dict[str, int], source_header_row: int, allow_sunday: bool) -> list[str]:
    if "date" not in header_map:
        raise ValueError("Source sheet missing required date header.")

    dates: dict[datetime, str] = {}
    current_date: Any = None
    for row in range(source_header_row + 1, ws.max_row + 1):
        row_date = ws.cell(row, header_map["date"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if current_date in (None, ""):
            continue
        try:
            parsed = parse_business_date(normalize(current_date))
        except ValueError:
            continue
        if parsed.weekday() == 6 and not allow_sunday:
            continue
        dates[parsed] = f"{parsed.year}/{parsed.month}/{parsed.day}"
    return [dates[key] for key in sorted(dates)]


def monitor_dates(ws: Worksheet, data_start_row: int) -> set[str]:
    dates: set[str] = set()
    for row in range(data_start_row, ws.max_row + 1):
        value = ws.cell(row, MONITOR_COLUMNS["date"]).value
        if value in (None, ""):
            continue
        try:
            dates.add(normalize(value))
        except ValueError:
            continue
    return dates


def resolve_source_headers(
    ws: Worksheet,
    preferred_header_row: int,
    aliases: dict[str, list[str]],
) -> tuple[int, dict[str, int]]:
    rows_to_try = [preferred_header_row, 1, 2, 3, 4, 5]
    for row in dict.fromkeys(rows_to_try):
        if row < 1 or row > ws.max_row:
            continue
        headers = find_header_columns(ws, row, aliases)
        if all(field in headers for field in ["date", "area", "dsp"]):
            return row, headers
    return preferred_header_row, find_header_columns(ws, preferred_header_row, aliases)


def worksheet_or_first(wb: Any, configured_name: str | None) -> Worksheet:
    if configured_name and configured_name in wb.sheetnames:
        return wb[configured_name]
    return wb[wb.sheetnames[0]]


def clean_excel_text(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    text = ILLEGAL_CHARACTERS_RE.sub("", value)
    text = text.replace("\ufeff", "").replace("\x00", "")
    return text.strip()


def decoded_text_score(text: str) -> int:
    if not text:
        return -1_000_000
    control_count = sum(
        1 for char in text
        if ord(char) < 32 and char not in ("\t", "\n", "\r")
    )
    replacement_count = text.count("\ufffd")
    useful_count = len(re.findall(r"[\u4e00-\u9fffA-Za-z0-9,%/()._\-\s]", text))
    return useful_count - (control_count * 100) - (replacement_count * 200)


def decode_csv_text(path: Path) -> str:
    raw = path.read_bytes()
    encodings = [
        "utf-8-sig",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
        "gb18030",
        "cp936",
        "cp1252",
    ]
    candidates: list[tuple[int, str]] = []
    for encoding in encodings:
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        candidates.append((decoded_text_score(text), text))
    if candidates:
        return max(candidates, key=lambda item: item[0])[1]
    return raw.decode("utf-8", errors="replace")


def read_csv_rows(path: Path) -> list[list[str]]:
    text = decode_csv_text(path)
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [
        [clean_excel_text(cell) for cell in row]
        for row in csv.reader(io.StringIO(text), dialect)
    ]


def load_tabular_workbook(path: Path, data_only: bool = True) -> Workbook:
    if path.suffix.lower() != ".csv":
        return load_workbook(path, data_only=data_only)

    wb = Workbook()
    ws = wb.active
    ws.title = path.stem[:31] or "Sheet1"
    for row in read_csv_rows(path):
        ws.append([clean_excel_text(cell) for cell in row])
    return wb


def find_header_columns(ws: Worksheet, header_row: int, aliases: dict[str, list[str]]) -> dict[str, int]:
    header_by_name: dict[str, int] = {}
    for cell in ws[header_row]:
        label = normalize(cell.value)
        if label:
            header_by_name[label] = cell.column

    found: dict[str, int] = {}
    for field, names in aliases.items():
        for name in names:
            col = header_by_name.get(normalize(name))
            if col:
                found[field] = col
                break
    return found


def copy_cell_format(source: Cell, target: Cell) -> None:
    if source.has_style:
        target._style = copy.copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy.copy(source.alignment)
    if source.border:
        target.border = copy.copy(source.border)
    if source.fill:
        target.fill = copy.copy(source.fill)
    if source.font:
        target.font = copy.copy(source.font)
    target.protection = copy.copy(source.protection)


def copy_row_format(ws: Worksheet, source_row: int, target_row: int, max_col: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        copy_cell_format(ws.cell(source_row, col), ws.cell(target_row, col))


def clear_inserted_rows(ws: Worksheet, start_row: int, row_count: int, max_col: int) -> None:
    for row in range(start_row, start_row + row_count):
        for col in range(1, max_col + 1):
            ws.cell(row, col).value = None


def apply_number_formats(ws: Worksheet, start_row: int, row_count: int) -> None:
    for row in range(start_row, start_row + row_count):
        for col in PERCENT_COLUMNS:
            ws.cell(row, col).number_format = "0.00%"


def apply_total_row_style(ws: Worksheet, row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        cell = ws.cell(row, col)
        cell.fill = copy.copy(TOTAL_ROW_FILL)
        cell.font = copy.copy(TOTAL_ROW_FONT)


def apply_green_metric_font(ws: Worksheet, row: int) -> None:
    is_total_row = is_total_label(ws.cell(row, MONITOR_COLUMNS["route"]).value)
    for metric in GREEN_METRICS:
        cell = ws.cell(row, MONITOR_COLUMNS[metric])
        if is_total_row:
            cell.font = copy.copy(TOTAL_ROW_FONT)
            continue
        value = percent_or_none(cell.value)
        if value == "-":
            continue
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            continue
        cell.font = copy.copy(GREEN_PASS_FONT if numeric_value >= 0.95 else GREEN_FAIL_FONT)


def apply_quality_metric_font(ws: Worksheet, row: int) -> None:
    count_metrics = [
        "fake_sign_qty",
        "complaint_qty",
        "suspected_lost_qty",
        "tracking_break_qty",
    ]
    for metric in count_metrics:
        cell = ws.cell(row, MONITOR_COLUMNS[metric])
        try:
            value = float(cell.value)
        except (TypeError, ValueError):
            continue
        cell.font = copy.copy(QUALITY_FAIL_FONT if value > 5 else QUALITY_PASS_FONT)

    rate_metrics = ["fake_sign_rate", "complaint_rate"]
    for metric in rate_metrics:
        cell = ws.cell(row, MONITOR_COLUMNS[metric])
        try:
            value = float(cell.value)
        except (TypeError, ValueError):
            continue
        cell.font = copy.copy(QUALITY_FAIL_FONT if value >= 0.005 else QUALITY_PASS_FONT)


def merge_project_quality_cells(ws: Worksheet, first_project_row: int, total_row: int) -> None:
    if total_row <= first_project_row + 1:
        return

    last_detail_row = total_row - 1
    for metric in QUALITY_METRICS:
        col = MONITOR_COLUMNS[metric]
        value = ws.cell(first_project_row, col).value
        number_format = ws.cell(first_project_row, col).number_format
        for row in range(first_project_row, last_detail_row + 1):
            force_set_cell_value(ws, row, col, None)
        ws.merge_cells(start_row=first_project_row, start_column=col, end_row=last_detail_row, end_column=col)
        detail_cell = ws.cell(first_project_row, col)
        detail_cell.value = value
        detail_cell.number_format = number_format
        detail_cell.alignment = copy.copy(GROUP_ALIGNMENT)

        total_cell = ws.cell(total_row, col)
        total_cell.value = value
        total_cell.number_format = number_format
        total_cell.fill = copy.copy(TOTAL_ROW_FILL)
        total_cell.font = copy.copy(TOTAL_ROW_FONT)
    apply_quality_metric_font(ws, first_project_row)


def rebuild_quality_metric_merges_from_routes(ws: Worksheet, start_row: int, projects: list[Project]) -> None:
    route_map: dict[str, Project] = {}
    for project in merge_projects(projects):
        for route in project.routes:
            route_map[normalize(route)] = project

    quality_cols = [MONITOR_COLUMNS[metric] for metric in QUALITY_METRICS]
    for col in quality_cols:
        unmerge_intersecting_range(ws, start_row, ws.max_row, col, col)

    row = start_row
    while row <= ws.max_row:
        route = normalize(ws.cell(row, MONITOR_COLUMNS["route"]).value)
        if route not in route_map:
            row += 1
            continue

        first_project_row = row
        detail_rows = [row]
        row += 1
        while row <= ws.max_row:
            route = normalize(ws.cell(row, MONITOR_COLUMNS["route"]).value)
            if is_total_label(route):
                total_row = row
                if len(detail_rows) > 1:
                    merge_project_quality_cells(ws, first_project_row, total_row)
                else:
                    for metric in QUALITY_METRICS:
                        col = MONITOR_COLUMNS[metric]
                        total_cell = ws.cell(total_row, col)
                        total_cell.value = ws.cell(first_project_row, col).value
                        total_cell.number_format = ws.cell(first_project_row, col).number_format
                        total_cell.fill = copy.copy(TOTAL_ROW_FILL)
                        total_cell.font = copy.copy(TOTAL_ROW_FONT)
                    apply_quality_metric_font(ws, first_project_row)
                break
            if route in route_map:
                detail_rows.append(row)
                row += 1
                continue
            break
        row += 1


def unmerge_intersecting_range(ws: Worksheet, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
    ranges_to_unmerge = []
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.max_row >= start_row
            and merged_range.min_row <= end_row
            and merged_range.max_col >= start_col
            and merged_range.min_col <= end_col
        ):
            ranges_to_unmerge.append(str(merged_range))
    for merged_range in ranges_to_unmerge:
        ws.unmerge_cells(merged_range)


def force_set_cell_value(ws: Worksheet, row: int, col: int, value: Any) -> None:
    existing = ws._cells.get((row, col))
    if isinstance(existing, MergedCell):
        del ws._cells[(row, col)]
    ws.cell(row, col).value = value


def rebuild_date_column_merges(ws: Worksheet, start_row: int) -> None:
    date_col = MONITOR_COLUMNS["date"]
    date_ranges = []
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_col <= date_col <= merged_range.max_col
            and merged_range.max_row >= start_row
        ):
            date_ranges.append(
                (
                    str(merged_range),
                    ws.cell(merged_range.min_row, date_col).value,
                    max(merged_range.min_row, start_row),
                    merged_range.max_row,
                )
            )

    for range_name, _, _, _ in date_ranges:
        ws.unmerge_cells(range_name)

    for _, value, first_row, last_row in date_ranges:
        for row in range(first_row, last_row + 1):
            try:
                force_set_cell_value(ws, row, date_col, value)
            except AttributeError:
                blocking_ranges = [
                    str(merged_range) for merged_range in ws.merged_cells.ranges
                    if (
                        merged_range.min_row <= row <= merged_range.max_row
                        and merged_range.min_col <= date_col <= merged_range.max_col
                    )
                ]
                for merged_range in blocking_ranges:
                    ws.unmerge_cells(merged_range)
                force_set_cell_value(ws, row, date_col, value)

    current_value = None
    current_start = None
    for row in range(start_row, ws.max_row + 2):
        value = ws.cell(row, date_col).value if row <= ws.max_row else None
        if value not in (None, ""):
            value = normalize(value)

        if value != current_value:
            if current_value not in (None, "") and current_start is not None and row - current_start > 1:
                ws.merge_cells(start_row=current_start, start_column=date_col, end_row=row - 1, end_column=date_col)
                ws.cell(current_start, date_col).value = current_value
            current_value = value
            current_start = row if value not in (None, "") else None


def apply_date_column_blocks(ws: Worksheet, blocks: list[tuple[str, int, int]]) -> None:
    date_col = MONITOR_COLUMNS["date"]
    for target_date, first_row, last_row in blocks:
        unmerge_intersecting_range(ws, first_row, last_row, date_col, date_col)
        for row in range(first_row, last_row + 1):
            force_set_cell_value(ws, row, date_col, None)
        if last_row > first_row:
            ws.merge_cells(start_row=first_row, start_column=date_col, end_row=last_row, end_column=date_col)
        ws.cell(first_row, date_col).value = target_date


def apply_vertical_merge_blocks(ws: Worksheet, col: int, blocks: list[tuple[Any, int, int]]) -> None:
    for value, first_row, last_row in blocks:
        unmerge_intersecting_range(ws, first_row, last_row, col, col)
        for row in range(first_row, last_row + 1):
            force_set_cell_value(ws, row, col, None)
        if last_row > first_row:
            ws.merge_cells(start_row=first_row, start_column=col, end_row=last_row, end_column=col)
        ws.cell(first_row, col).value = value
        ws.cell(first_row, col).font = copy.copy(GROUP_FONT)
        ws.cell(first_row, col).alignment = copy.copy(GROUP_ALIGNMENT)


def rebuild_all_date_merges_from_markers(ws: Worksheet, start_row: int) -> None:
    date_col = MONITOR_COLUMNS["date"]
    markers: dict[int, str] = {}
    ranges_to_unmerge = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col <= date_col <= merged_range.max_col and merged_range.max_row >= start_row:
            value = ws.cell(merged_range.min_row, date_col).value
            if value not in (None, ""):
                markers[max(merged_range.min_row, start_row)] = normalize(value)
            ranges_to_unmerge.append(str(merged_range))

    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row, date_col).value
        if value not in (None, ""):
            markers[row] = normalize(value)

    for range_name in ranges_to_unmerge:
        ws.unmerge_cells(range_name)
    for row in range(start_row, ws.max_row + 1):
        force_set_cell_value(ws, row, date_col, None)

    sorted_markers = sorted(markers.items())
    for idx, (first_row, date_value) in enumerate(sorted_markers):
        last_row = sorted_markers[idx + 1][0] - 1 if idx + 1 < len(sorted_markers) else ws.max_row
        if last_row > first_row:
            ws.merge_cells(start_row=first_row, start_column=date_col, end_row=last_row, end_column=date_col)
        cell = ws.cell(first_row, date_col)
        cell.value = date_value
        cell.font = copy.copy(GROUP_FONT)
        cell.alignment = copy.copy(GROUP_ALIGNMENT)


def format_existing_date_cells(ws: Worksheet, start_row: int) -> None:
    date_col = MONITOR_COLUMNS["date"]
    seen_rows: set[int] = set()
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col <= date_col <= merged_range.max_col and merged_range.max_row >= start_row:
            cell = ws.cell(merged_range.min_row, date_col)
            if cell.value not in (None, ""):
                cell.font = copy.copy(GROUP_FONT)
                cell.alignment = copy.copy(GROUP_ALIGNMENT)
                seen_rows.add(merged_range.min_row)
    for row in range(start_row, ws.max_row + 1):
        if row in seen_rows:
            continue
        cell = ws.cell(row, date_col)
        if cell.value not in (None, ""):
            cell.font = copy.copy(GROUP_FONT)
            cell.alignment = copy.copy(GROUP_ALIGNMENT)


def rebuild_area_dsp_merges_from_routes(ws: Worksheet, start_row: int, projects: list[Project]) -> None:
    area_col = MONITOR_COLUMNS["area"]
    dsp_col = MONITOR_COLUMNS["dsp"]
    route_col = MONITOR_COLUMNS["route"]
    route_map: dict[str, Project] = {}
    for project in projects:
        for route in project.routes:
            route_map[normalize(route)] = project

    row_projects: dict[int, Project] = {}
    last_project: Project | None = None
    for row in range(start_row, ws.max_row + 1):
        route = normalize(ws.cell(row, route_col).value)
        if route in route_map:
            last_project = route_map[route]
            row_projects[row] = last_project
        elif is_total_label(route) and last_project is not None:
            row_projects[row] = last_project
        else:
            last_project = None

    unmerge_intersecting_range(ws, start_row, ws.max_row, area_col, dsp_col)
    for row in range(start_row, ws.max_row + 1):
        if row in row_projects:
            project = row_projects[row]
            force_set_cell_value(ws, row, area_col, project.area)
            force_set_cell_value(ws, row, dsp_col, project.dsp)

    area_blocks: list[tuple[str, int, int]] = []
    block_area = None
    block_start = None
    for row in range(start_row, ws.max_row + 2):
        project = row_projects.get(row)
        area = project.area if project else None
        if area != block_area:
            if block_area is not None and block_start is not None and row - block_start > 1:
                area_blocks.append((block_area, block_start, row - 1))
            block_area = area
            block_start = row if area is not None else None

    dsp_blocks: list[tuple[str, int, int]] = []
    block_dsp = None
    block_start = None
    for row in range(start_row, ws.max_row + 2):
        project = row_projects.get(row)
        dsp = project.dsp if project else None
        if dsp != block_dsp:
            if block_dsp is not None and block_start is not None and row - block_start > 1:
                dsp_blocks.append((block_dsp, block_start, row - 1))
            block_dsp = dsp
            block_start = row if dsp is not None else None

    apply_vertical_merge_blocks(ws, area_col, area_blocks)
    apply_vertical_merge_blocks(ws, dsp_col, dsp_blocks)


def expanded_date_by_row(ws: Worksheet, start_row: int) -> dict[int, str]:
    date_col = MONITOR_COLUMNS["date"]
    dates: dict[int, str] = {}
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col <= date_col <= merged_range.max_col and merged_range.max_row >= start_row:
            value = ws.cell(merged_range.min_row, date_col).value
            if value not in (None, ""):
                normalized = normalize(value)
                for row in range(max(start_row, merged_range.min_row), merged_range.max_row + 1):
                    dates[row] = normalized
    current_date = None
    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row, date_col).value
        if value not in (None, ""):
            current_date = normalize(value)
        if current_date:
            dates.setdefault(row, current_date)
    return dates


def is_blank_like(value: Any) -> bool:
    return value in (None, "") or (isinstance(value, str) and value.strip() == "-")


def is_total_label(value: Any) -> bool:
    normalized = normalize(value)
    return normalized in {normalize(TOTAL_LABEL), "북셕"}


def collect_existing_date_blocks(ws: Worksheet, start_row: int) -> list[tuple[str, int, int]]:
    date_col = MONITOR_COLUMNS["date"]
    markers: dict[int, str] = {}
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col <= date_col <= merged_range.max_col and merged_range.max_row >= start_row:
            value = ws.cell(merged_range.min_row, date_col).value
            if value not in (None, ""):
                markers[merged_range.min_row] = normalize(value)
    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row, date_col).value
        if value not in (None, ""):
            markers[row] = normalize(value)

    sorted_markers = sorted(markers.items())
    blocks: list[tuple[str, int, int]] = []
    for idx, (first_row, date_value) in enumerate(sorted_markers):
        last_row = sorted_markers[idx + 1][0] - 1 if idx + 1 < len(sorted_markers) else ws.max_row
        blocks.append((date_value, first_row, last_row))
    return blocks


def merged_anchor(ws: Worksheet, row: int, col: int) -> tuple[int, int]:
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
            return merged_range.min_row, merged_range.min_col
    return row, col


def merged_cell_value(ws: Worksheet, row: int, col: int) -> Any:
    anchor_row, anchor_col = merged_anchor(ws, row, col)
    return ws.cell(anchor_row, anchor_col).value


def set_station_metric_value(ws: Worksheet, row: int, metric: str, value: Any, total_row: bool) -> None:
    col = MONITOR_COLUMNS[metric]
    anchor_row, anchor_col = merged_anchor(ws, row, col)
    force_set_cell_value(ws, anchor_row, anchor_col, value)
    cell = ws.cell(anchor_row, anchor_col)
    if metric.endswith("_rate"):
        cell.number_format = "0.00%"
    if total_row:
        cell.font = copy.copy(TOTAL_ROW_FONT)
        cell.fill = copy.copy(TOTAL_ROW_FILL)
    else:
        apply_quality_metric_font(ws, anchor_row)


def station_level_updates(metrics: dict[str, Any]) -> dict[str, Any]:
    updates: dict[str, Any] = {}
    for metric in QUALITY_METRICS:
        raw_value = metrics.get(metric)
        if metric.endswith("_rate"):
            updates[metric] = percent_or_none(raw_value)
        else:
            updates[metric] = display_value(raw_value)
    return updates


def backfill_station_quality_metrics(
    ws: Worksheet,
    source_ws: Worksheet,
    source_headers: dict[str, int],
    source_header_row: int,
    projects: list[Project],
    start_row: int,
) -> None:
    project_keys = {
        (normalize(project.area), normalize(project.dsp))
        for project in merge_projects(projects)
    }
    row_dates = expanded_date_by_row(ws, start_row)
    source_cache: dict[str, dict[tuple[str, str], dict[str, Any]]] = {}

    for row in range(start_row, ws.max_row + 1):
        route = normalize(ws.cell(row, MONITOR_COLUMNS["route"]).value)
        if not route or (route not in {normalize(TOTAL_LABEL), "攵侅厱"} and route == ""):
            continue

        row_date = row_dates.get(row)
        if not row_date:
            continue

        area = normalize(merged_cell_value(ws, row, MONITOR_COLUMNS["area"]))
        dsp = normalize(merged_cell_value(ws, row, MONITOR_COLUMNS["dsp"]))
        if not area or not dsp or (area, dsp) not in project_keys:
            continue

        if row_date not in source_cache:
            source_cache[row_date] = build_source_index(source_ws, source_headers, row_date, source_header_row)
        metrics = source_cache[row_date].get((area, dsp))
        if metrics is None:
            continue

        is_total = is_total_label(route)
        for metric, new_value in station_level_updates(metrics).items():
            set_station_metric_value(ws, row, metric, new_value, is_total)


# Backward-compatible name used by main_v2.
def backfill_customer_metrics(
    ws: Worksheet,
    source_ws: Worksheet,
    source_headers: dict[str, int],
    source_header_row: int,
    projects: list[Project],
    start_row: int,
) -> None:
    backfill_station_quality_metrics(ws, source_ws, source_headers, source_header_row, projects, start_row)

def build_source_index(
    ws: Worksheet,
    header_map: dict[str, int],
    target_date: str,
    source_header_row: int,
) -> dict[tuple[str, str], dict[str, Any]]:
    required = ["date", "area", "dsp"]
    missing = [field for field in required if field not in header_map]
    if missing:
        raise ValueError(f"Source sheet missing required headers: {', '.join(missing)}")

    index: dict[tuple[str, str], dict[str, Any]] = {}
    current_date: Any = None
    current_area: Any = None
    for row in range(source_header_row + 1, ws.max_row + 1):
        row_date = ws.cell(row, header_map["date"]).value
        row_area = ws.cell(row, header_map["area"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if row_area not in (None, ""):
            current_area = row_area

        if not same_date(current_date, target_date):
            continue
        area = normalize(current_area)
        dsp = normalize(ws.cell(row, header_map["dsp"]).value)
        if not (area and dsp):
            continue

        metrics: dict[str, Any] = {}
        for metric in BLUE_METRICS:
            col = header_map.get(metric)
            metrics[metric] = ws.cell(row, col).value if col else None
        index[(area, dsp)] = metrics
    return index


def build_center_timing_index(ws: Worksheet, target_date: str) -> dict[tuple[str, str, str], dict[str, Any]]:
    headers = find_header_columns(
        ws,
        1,
        {
            "date": ["中心签出日期"],
            "area": ["转运中心 "],
            "dsp": ["站点 "],
            "route": ["路区 "],
            "pod_2400": ["2400签收率"],
        },
    )
    required = ["date", "area", "dsp", "route", "pod_2400"]
    missing = [field for field in required if field not in headers]
    if missing:
        raise ValueError(f"中心签出时效 missing required headers: {', '.join(missing)}")

    index: dict[tuple[str, str, str], dict[str, Any]] = {}
    current_date: Any = None
    current_area: Any = None
    current_dsp: Any = None
    for row in range(2, ws.max_row + 1):
        row_date = ws.cell(row, headers["date"]).value
        row_area = ws.cell(row, headers["area"]).value
        row_dsp = ws.cell(row, headers["dsp"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if row_area not in (None, ""):
            current_area = row_area
        if row_dsp not in (None, ""):
            current_dsp = row_dsp

        if not same_date(current_date, target_date):
            continue
        area = normalize(current_area)
        dsp = normalize(current_dsp)
        route = normalize(ws.cell(row, headers["route"]).value)
        if not (area and dsp and route):
            continue
        denominator = parse_number(ws.cell(row, 7).value)
        numerator = parse_number(ws.cell(row, 8).value)
        index[(area, dsp, route)] = {
            "rate": value_with_min_volume(ws.cell(row, headers["pod_2400"]).value, denominator),
            "numerator": numerator,
            "denominator": denominator,
        }
    return index


def build_center_dispatch_index(ws: Worksheet, target_date: str) -> dict[tuple[str, str, str], Any]:
    headers = find_header_columns(
        ws,
        1,
        {
            "date": ["中心签出日期"],
            "area": ["转运中心 "],
            "dsp": ["站点 "],
            "route": ["路区 "],
            "dispatch_qty": ["中心签出量"],
        },
    )
    required = ["date", "area", "dsp", "route", "dispatch_qty"]
    missing = [field for field in required if field not in headers]
    if missing:
        raise ValueError(f"中心签出时效 missing required headers: {', '.join(missing)}")

    index: dict[tuple[str, str, str], Any] = {}
    current_date: Any = None
    current_area: Any = None
    current_dsp: Any = None
    for row in range(2, ws.max_row + 1):
        row_date = ws.cell(row, headers["date"]).value
        row_area = ws.cell(row, headers["area"]).value
        row_dsp = ws.cell(row, headers["dsp"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if row_area not in (None, ""):
            current_area = row_area
        if row_dsp not in (None, ""):
            current_dsp = row_dsp

        if not same_date(current_date, target_date):
            continue
        area = normalize(current_area)
        dsp = normalize(current_dsp)
        route = normalize(ws.cell(row, headers["route"]).value)
        if not (area and dsp and route):
            continue
        volume = parse_number(ws.cell(row, headers["dispatch_qty"]).value)
        if volume is None:
            value = display_value(ws.cell(row, headers["dispatch_qty"]).value)
        elif volume.is_integer():
            value = int(volume)
        else:
            value = volume
        index[(area, dsp, route)] = value
    return index


def build_signin_timing_index(ws: Worksheet, target_date: str) -> dict[tuple[str, str, str], dict[str, Any]]:
    headers = find_header_columns(
        ws,
        1,
        {
            "date": ["中心签入日期"],
            "area": ["转运中心  "],
            "dsp": ["站点  "],
            "route": ["路区(快递员工作区域) "],
            "ascan_dscan_72h": ["72H妥投率"],
        },
    )
    required = ["date", "area", "dsp", "route", "ascan_dscan_72h"]
    missing = [field for field in required if field not in headers]
    if missing:
        raise ValueError(f"中心签入时效 missing required headers: {', '.join(missing)}")

    index: dict[tuple[str, str, str], dict[str, Any]] = {}
    current_date: Any = None
    current_area: Any = None
    current_dsp: Any = None
    for row in range(2, ws.max_row + 1):
        row_date = ws.cell(row, headers["date"]).value
        row_area = ws.cell(row, headers["area"]).value
        row_dsp = ws.cell(row, headers["dsp"]).value
        if row_date not in (None, ""):
            current_date = row_date
        if row_area not in (None, ""):
            current_area = row_area
        if row_dsp not in (None, ""):
            current_dsp = row_dsp

        if not same_date(current_date, target_date):
            continue
        area = normalize(current_area)
        dsp = normalize(current_dsp)
        route = normalize(ws.cell(row, headers["route"]).value)
        if not (area and dsp and route):
            continue
        denominator = parse_number(ws.cell(row, 7).value)
        numerator = parse_number(ws.cell(row, 12).value)
        index[(area, dsp, route)] = {
            "rate": value_with_min_volume(ws.cell(row, headers["ascan_dscan_72h"]).value, denominator),
            "numerator": numerator,
            "denominator": denominator,
        }
    return index


def percent_or_none(value: Any) -> Any:
    if value is None or value == "":
        return "-"
    if isinstance(value, str) and value.strip().endswith("%"):
        return float(value.strip().rstrip("%")) / 100
    return value


def display_value(value: Any) -> Any:
    if value is None or value == "":
        return "-"
    return value


def parse_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def value_with_min_volume(rate: Any, volume: Any, min_volume: int = 10) -> Any:
    parsed_volume = parse_number(volume)
    if parsed_volume is None:
        return "-"
    if parsed_volume <= min_volume:
        return "-"
    return percent_or_none(rate)


def weighted_rate_from_items(items: list[Any], min_denominator: int = 10) -> Any:
    numerator_total = 0.0
    denominator_total = 0.0
    found = False
    for item in items:
        if not isinstance(item, dict):
            continue
        numerator = parse_number(item.get("numerator"))
        denominator = parse_number(item.get("denominator"))
        if numerator is None or denominator is None:
            continue
        numerator_total += numerator
        denominator_total += denominator
        found = True
    if not found or denominator_total <= min_denominator:
        return "-"
    return numerator_total / denominator_total


def metric_item_rate(item: Any) -> Any:
    if isinstance(item, dict):
        return item.get("rate", "-")
    if item is None:
        return "-"
    return item


def combine_detail_values(values: list[Any]) -> Any:
    clean_values = [value for value in values if value not in (None, "")]
    if not clean_values:
        return "-"
    if len(clean_values) == 1:
        return clean_values[0]
    if any(isinstance(value, str) and value.strip() == "-" for value in clean_values):
        return "-"
    try:
        return sum(float(value) for value in clean_values)
    except (TypeError, ValueError):
        return clean_values[0]


def timing_route_key(area: str, dsp: str, route: str) -> str:
    return f"{normalize(area)}|{normalize(dsp)}|{normalize(route)}"


def resolve_timing_route(config: dict[str, Any], area: str, dsp: str, route: str) -> str:
    overrides = config.get("center_timing_route_overrides", {})
    return overrides.get(timing_route_key(area, dsp, route), route)


def write_monitor_block(
    ws: Worksheet,
    projects: list[Project],
    source_index: dict[tuple[str, str], dict[str, Any]],
    timing_index: dict[tuple[str, str, str], dict[str, Any]],
    center_dispatch_index: dict[tuple[str, str, str], Any],
    signin_excluded_index: dict[tuple[str, str, str], dict[str, Any]],
    signin_included_index: dict[tuple[str, str, str], dict[str, Any]],
    config: dict[str, Any],
    target_date: str,
    start_row: int,
    max_col: int,
) -> tuple[list[str], list[tuple[str, int, int]], list[tuple[str, int, int]]]:
    missing_rows: list[str] = []
    dsp_ranges: list[tuple[str, int, int]] = []
    current_row = start_row
    area_ranges: list[tuple[str, int, int]] = []
    current_area: str | None = None
    current_area_start = start_row

    for project in projects:
        if current_area is None:
            current_area = project.area
        elif project.area != current_area:
            area_ranges.append((current_area, current_area_start, current_row - 1))
            current_area = project.area
            current_area_start = current_row

        detail_rows: list[int] = []
        first_project_row = current_row

        for route in project.routes:
            key = (normalize(project.area), normalize(project.dsp))
            metrics = source_index.get(key)
            if metrics is None:
                missing_rows.append(f"{project.area} / {project.dsp}")
                metrics = {}

            ws.cell(current_row, MONITOR_COLUMNS["date"]).value = target_date
            ws.cell(current_row, MONITOR_COLUMNS["area"]).value = project.area
            ws.cell(current_row, MONITOR_COLUMNS["dsp"]).value = project.dsp
            ws.cell(current_row, MONITOR_COLUMNS["route"]).value = route

            for metric in BLUE_METRICS:
                value = metrics.get(metric)
                if metric == "dispatch_qty":
                    timing_route = resolve_timing_route(config, project.area, project.dsp, route)
                    dispatch_key = (normalize(project.area), normalize(project.dsp), normalize(timing_route))
                    value = center_dispatch_index.get(dispatch_key, value)
                if metric.endswith("_rate"):
                    value = percent_or_none(value)
                else:
                    value = display_value(value)
                ws.cell(current_row, MONITOR_COLUMNS[metric]).value = value

            for metric in GREEN_METRICS:
                ws.cell(current_row, MONITOR_COLUMNS[metric]).value = "-"
            timing_route = resolve_timing_route(config, project.area, project.dsp, route)
            pod_key = (normalize(project.area), normalize(project.dsp), normalize(timing_route))
            ws.cell(current_row, MONITOR_COLUMNS["pod_2400"]).value = metric_item_rate(timing_index.get(pod_key))
            signin_key = (normalize(project.area), normalize(project.dsp), normalize(route))
            ws.cell(current_row, MONITOR_COLUMNS["ascan_dscan_72h_weekend_holiday"]).value = (
                metric_item_rate(signin_excluded_index.get(signin_key))
            )
            ws.cell(current_row, MONITOR_COLUMNS["ascan_dscan_72h_business_day"]).value = (
                metric_item_rate(signin_included_index.get(signin_key))
            )
            apply_green_metric_font(ws, current_row)
            apply_quality_metric_font(ws, current_row)

            detail_rows.append(current_row)
            current_row += 1

        total_row = current_row
        ws.cell(total_row, MONITOR_COLUMNS["date"]).value = target_date
        ws.cell(total_row, MONITOR_COLUMNS["area"]).value = project.area
        ws.cell(total_row, MONITOR_COLUMNS["dsp"]).value = project.dsp
        ws.cell(total_row, MONITOR_COLUMNS["route"]).value = TOTAL_LABEL

        for metric in BLUE_METRICS:
            col = MONITOR_COLUMNS[metric]
            detail_values = [ws.cell(row, col).value for row in detail_rows]
            if metric in QUALITY_METRICS:
                ws.cell(total_row, col).value = detail_values[0] if detail_values else "-"
            else:
                ws.cell(total_row, col).value = combine_detail_values(detail_values)

        pod_items: list[Any] = []
        signin_excluded_items: list[Any] = []
        signin_included_items: list[Any] = []
        for route in project.routes:
            timing_route = resolve_timing_route(config, project.area, project.dsp, route)
            pod_key = (normalize(project.area), normalize(project.dsp), normalize(timing_route))
            signin_key = (normalize(project.area), normalize(project.dsp), normalize(route))
            pod_items.append(timing_index.get(pod_key))
            signin_excluded_items.append(signin_excluded_index.get(signin_key))
            signin_included_items.append(signin_included_index.get(signin_key))

        ws.cell(total_row, MONITOR_COLUMNS["pod_2400"]).value = weighted_rate_from_items(pod_items)
        ws.cell(total_row, MONITOR_COLUMNS["ascan_dscan_72h_weekend_holiday"]).value = weighted_rate_from_items(signin_excluded_items)
        ws.cell(total_row, MONITOR_COLUMNS["ascan_dscan_72h_business_day"]).value = weighted_rate_from_items(signin_included_items)

        ws.merge_cells(
            start_row=first_project_row,
            start_column=MONITOR_COLUMNS["dsp"],
            end_row=total_row,
            end_column=MONITOR_COLUMNS["dsp"],
        )
        dsp_ranges.append((project.dsp, first_project_row, total_row))
        apply_total_row_style(ws, total_row, max_col)
        merge_project_quality_cells(ws, first_project_row, total_row)

        current_row += 1

    if current_area is not None:
        area_ranges.append((current_area, current_area_start, current_row - 1))

    ws.merge_cells(
        start_row=start_row,
        start_column=MONITOR_COLUMNS["date"],
        end_row=current_row - 1,
        end_column=MONITOR_COLUMNS["date"],
    )
    ws.cell(start_row, MONITOR_COLUMNS["date"]).value = target_date

    for area, first_row, last_row in area_ranges:
        ws.merge_cells(
            start_row=first_row,
            start_column=MONITOR_COLUMNS["area"],
            end_row=last_row,
            end_column=MONITOR_COLUMNS["area"],
        )
        ws.cell(first_row, MONITOR_COLUMNS["area"]).value = area

    return missing_rows, area_ranges, dsp_ranges

def main() -> None:
    args = parse_args()
    config = load_config(Path(args.config))

    workbook_path, center_timing_path, signin_excluded_path, signin_included_path, output_path = resolve_daily_paths(args)
    wb = load_workbook(workbook_path)

    monitor_ws = wb[config["monitor_sheet"]]
    source_ws = wb[config["source_sheet"]]

    projects = [
        Project(
            area=item["area"],
            dsp=item["dsp"],
            routes=item["routes"],
            start_date=item.get("start_date"),
        )
        for item in config["projects"]
    ]
    start_row = int(config.get("data_start_row", 4))
    source_header_row = int(config.get("source_header_row", 1))

    source_header_row, source_headers = resolve_source_headers(
        source_ws,
        source_header_row,
        config["metric_aliases"],
    )
    target_date = normalize(args.date) if args.date else latest_source_date(source_ws, source_headers, source_header_row)
    target_dt = parse_business_date(target_date)
    if target_dt.weekday() == 6 and not args.allow_sunday:
        print(f"Skipped: {target_date} is Sunday, and DSP质量总览 normally has no Sunday data.")
        return
    print(f"Monitor date: {target_date}")
    source_index = build_source_index(source_ws, source_headers, target_date, source_header_row)

    max_col = max(MONITOR_COLUMNS.values())
    monitor_ws.insert_rows(start_row, rows_to_insert)
    template_row = start_row + rows_to_insert
    for row in range(start_row, start_row + rows_to_insert):
        copy_row_format(monitor_ws, template_row, row, max_col)
    clear_inserted_rows(monitor_ws, start_row, rows_to_insert, max_col)
    apply_number_formats(monitor_ws, start_row, rows_to_insert)

    holidays = config.get("france_holidays", [])
    pod_source_date = previous_service_date(target_date, holidays)
    timing_index: dict[tuple[str, str, str], dict[str, Any]] = {}
    center_dispatch_index: dict[tuple[str, str, str], Any] = {}
    timing_wb = load_tabular_workbook(center_timing_path, data_only=True)
    timing_sheet_name = config.get("center_timing_sheet") or timing_wb.sheetnames[0]
    timing_index = build_center_timing_index(timing_wb[timing_sheet_name], pod_source_date)
    center_dispatch_index = build_center_dispatch_index(timing_wb[timing_sheet_name], target_date)
    print(f"2400POD source date: {pod_source_date}")

    signin_excluded_index: dict[tuple[str, str, str], dict[str, Any]] = {}
    signin_included_index: dict[tuple[str, str, str], dict[str, Any]] = {}
    signin_excluded_date = service_days_before(target_date, 3, holidays)
    signin_included_date = calendar_days_before(target_date, 3)
    signin_excluded_wb = load_tabular_workbook(signin_excluded_path, data_only=True)
    signin_excluded_sheet = config.get("signin_excluded_sheet") or signin_excluded_wb.sheetnames[0]
    signin_excluded_index = build_signin_timing_index(
        signin_excluded_wb[signin_excluded_sheet],
        signin_excluded_date,
    )
    print(f"ASCAN-DSCAN 72H excluded source date: {signin_excluded_date}")

    signin_included_wb = load_tabular_workbook(signin_included_path, data_only=True)
    signin_included_sheet = config.get("signin_included_sheet") or signin_included_wb.sheetnames[0]
    signin_included_index = build_signin_timing_index(
        signin_included_wb[signin_included_sheet],
        signin_included_date,
    )
    print(f"ASCAN-DSCAN 72H included source date: {signin_included_date}")

    missing_rows = write_monitor_block(
        monitor_ws,
        projects,
        source_index,
        timing_index,
        center_dispatch_index,
        signin_excluded_index,
        signin_included_index,
        config,
        target_date,
        start_row,
        max_col,
    )

    output_path = save_output_workbook(wb, output_path)

    print(f"Saved: {output_path}")
    if missing_rows:
        print("Rows not found in source sheet:")
        for item in missing_rows:
            print(f"- {item}")


def main_v2() -> None:
    args = parse_args()
    config = load_config(Path(args.config))

    workbook_path, center_timing_path, signin_excluded_path, signin_included_path, output_path = resolve_daily_paths(args)
    wb = load_workbook(workbook_path)

    monitor_ws = wb[config["monitor_sheet"]]
    source_ws = wb[config["source_sheet"]]

    projects = [
        Project(area=item["area"], dsp=item["dsp"], routes=item["routes"])
        for item in config["projects"]
    ]
    rows_to_insert = sum(len(project.routes) + 1 for project in projects)
    start_row = int(config.get("data_start_row", 4))
    source_header_row = int(config.get("source_header_row", 1))
    source_header_row, source_headers = resolve_source_headers(
        source_ws,
        source_header_row,
        config["metric_aliases"],
    )

    if args.date:
        dates_to_update = [normalize(args.date)]
    else:
        existing_dates = monitor_dates(monitor_ws, start_row)
        dates_to_update = [
            date for date in source_dates(source_ws, source_headers, source_header_row, args.allow_sunday)
            if date not in existing_dates
        ]

    if not dates_to_update:
        print("No missing monitor dates to update.")

    holidays = config.get("france_holidays", [])
    max_col = max(MONITOR_COLUMNS.values())

    timing_wb = load_tabular_workbook(center_timing_path, data_only=True)
    timing_ws = worksheet_or_first(timing_wb, config.get("center_timing_sheet"))

    signin_excluded_wb = load_tabular_workbook(signin_excluded_path, data_only=True)
    signin_excluded_ws = worksheet_or_first(signin_excluded_wb, config.get("signin_excluded_sheet"))

    signin_included_wb = load_tabular_workbook(signin_included_path, data_only=True)
    signin_included_ws = worksheet_or_first(signin_included_wb, config.get("signin_included_sheet"))

    all_missing_rows: list[str] = []
    inserted_date_blocks: list[tuple[str, int, int]] = []
    existing_date_blocks = collect_existing_date_blocks(monitor_ws, start_row)
    inserted_area_blocks: list[tuple[str, int, int]] = []
    inserted_dsp_blocks: list[tuple[str, int, int]] = []
    for target_date in dates_to_update:
        target_dt = parse_business_date(target_date)
        if target_dt.weekday() == 6 and not args.allow_sunday:
            print(f"Skipped: {target_date} is Sunday, and DSP质量总览 normally has no Sunday data.")
            continue

        print(f"Monitor date: {target_date}")
        projects_for_date = active_projects(projects, target_date)
        rows_to_insert = sum(len(project.routes) + 1 for project in projects_for_date)
        inserted_date_blocks = [
            (date, first_row + rows_to_insert, last_row + rows_to_insert)
            for date, first_row, last_row in inserted_date_blocks
        ]
        existing_date_blocks = [
            (date, first_row + rows_to_insert, last_row + rows_to_insert)
            for date, first_row, last_row in existing_date_blocks
        ]
        inserted_area_blocks = [
            (area, first_row + rows_to_insert, last_row + rows_to_insert)
            for area, first_row, last_row in inserted_area_blocks
        ]
        inserted_dsp_blocks = [
            (dsp, first_row + rows_to_insert, last_row + rows_to_insert)
            for dsp, first_row, last_row in inserted_dsp_blocks
        ]
        inserted_date_blocks.append((target_date, start_row, start_row + rows_to_insert - 1))
        source_index = build_source_index(source_ws, source_headers, target_date, source_header_row)

        monitor_ws.insert_rows(start_row, rows_to_insert)
        template_row = start_row + rows_to_insert
        for row in range(start_row, start_row + rows_to_insert):
            copy_row_format(monitor_ws, template_row, row, max_col)
        clear_inserted_rows(monitor_ws, start_row, rows_to_insert, max_col)
        apply_number_formats(monitor_ws, start_row, rows_to_insert)
        unmerge_intersecting_range(
            monitor_ws,
            start_row,
            start_row + rows_to_insert - 1,
            MONITOR_COLUMNS["date"],
            MONITOR_COLUMNS["dsp"],
        )

        pod_source_date = previous_service_date(target_date, holidays)
        timing_index = build_center_timing_index(timing_ws, pod_source_date)
        center_dispatch_index = build_center_dispatch_index(timing_ws, target_date)
        print(f"2400POD source date: {pod_source_date}")

        signin_excluded_date = service_days_before(target_date, 3, holidays)
        signin_included_date = calendar_days_before(target_date, 3)
        signin_excluded_index = build_signin_timing_index(signin_excluded_ws, signin_excluded_date)
        print(f"ASCAN-DSCAN 72H excluded source date: {signin_excluded_date}")
        signin_included_index = build_signin_timing_index(signin_included_ws, signin_included_date)
        print(f"ASCAN-DSCAN 72H included source date: {signin_included_date}")

        missing_rows, area_blocks, dsp_blocks = write_monitor_block(
            monitor_ws,
            projects_for_date,
            source_index,
            timing_index,
            center_dispatch_index,
            signin_excluded_index,
            signin_included_index,
            config,
            target_date,
            start_row,
            max_col,
        )
        all_missing_rows.extend(f"{target_date}: {item}" for item in missing_rows)
        inserted_area_blocks.extend(area_blocks)
        inserted_dsp_blocks.extend(dsp_blocks)

    apply_date_column_blocks(monitor_ws, inserted_date_blocks + existing_date_blocks)
    format_existing_date_cells(monitor_ws, start_row)
    apply_number_formats(monitor_ws, start_row, monitor_ws.max_row - start_row + 1)
    apply_vertical_merge_blocks(monitor_ws, MONITOR_COLUMNS["area"], inserted_area_blocks)
    apply_vertical_merge_blocks(monitor_ws, MONITOR_COLUMNS["dsp"], inserted_dsp_blocks)
    rebuild_area_dsp_merges_from_routes(monitor_ws, start_row, merge_projects(projects))
    rebuild_quality_metric_merges_from_routes(monitor_ws, start_row, projects)
    backfill_customer_metrics(monitor_ws, source_ws, source_headers, source_header_row, projects, start_row)
    apply_number_formats(monitor_ws, start_row, monitor_ws.max_row - start_row + 1)

    output_path = save_output_workbook(wb, output_path)

    print(f"Saved: {output_path}")
    if all_missing_rows:
        print("Rows not found in source sheet:")
        for item in all_missing_rows:
            print(f"- {item}")


if __name__ == "__main__":
    try:
        main_v2()
    except (FileNotFoundError, ValueError) as exc:
        raise SystemExit(f"错误: {exc}") from None

