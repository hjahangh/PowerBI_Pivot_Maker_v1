#!/usr/bin/env python3
"""Create the requested three-level budget pivot table from a raw Excel file.

By default, the output hierarchy matches the supplied example:

    Dept Name -> Cost Group -> Cost Category -> Grand Total

The three source columns can be changed with --group-columns. Group values are
always discovered from the current data; no department, group, or category
names are hard-coded.

Only Python's standard library is required.
"""

from __future__ import annotations

import argparse
import csv
import io
import posixpath
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, TextIO, Tuple
from xml.etree import ElementTree
from xml.sax.saxutils import escape


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "data.xlsx"
DEFAULT_OUTPUT = SCRIPT_DIR / "PowerBI_Pivot.xlsx"
DEFAULT_GROUP_FIELDS: Tuple[str, str, str] = (
    "Dept Name",
    "Cost Group",
    "Cost Category",
)
VALUE_FIELDS: Tuple[str, ...] = (
    "Spent Budget",
    "Pending Spent",
    "Authorized Budget",
    "Remaing Budget",
)
OUTPUT_HEADER: Tuple[str, ...] = (
    "Row Labels",
    "Sum of Spent Budget",
    "Sum of Pending Spent",
    "Sum of Authorized Budget",
    "Sum of Remaing Budget",
)


@dataclass
class Totals:
    """Sums plus flags that distinguish an empty value from a real zero."""

    values: List[Decimal] = field(
        default_factory=lambda: [Decimal("0") for _ in VALUE_FIELDS]
    )
    has_value: List[bool] = field(
        default_factory=lambda: [False for _ in VALUE_FIELDS]
    )

    def add(self, amounts: Iterable[Optional[Decimal]]) -> None:
        for index, amount in enumerate(amounts):
            if amount is not None:
                self.values[index] += amount
                self.has_value[index] = True


@dataclass
class PivotData:
    """Aggregated values for every level of the pivot hierarchy."""

    departments: Dict[str, Totals] = field(default_factory=dict)
    groups: Dict[Tuple[str, str], Totals] = field(default_factory=dict)
    categories: Dict[Tuple[str, str, str], Totals] = field(default_factory=dict)
    grand_total: Totals = field(default_factory=Totals)
    detail_count: int = 0

    def add(
        self,
        labels: Tuple[str, str, str],
        amounts: Iterable[Optional[Decimal]],
    ) -> None:
        department, group, category = labels
        amount_list = list(amounts)
        self.departments.setdefault(department, Totals()).add(amount_list)
        self.groups.setdefault((department, group), Totals()).add(amount_list)
        self.categories.setdefault((department, group, category), Totals()).add(
            amount_list
        )
        self.grand_total.add(amount_list)
        self.detail_count += 1

    def rows(self) -> Iterator[Tuple[str, str, Totals]]:
        """Yield formatted hierarchy rows in deterministic display order."""

        for department in sorted(self.departments, key=str.casefold):
            yield "department", department, self.departments[department]

            groups = sorted(
                (group for dept, group in self.groups if dept == department),
                key=str.casefold,
            )
            for group in groups:
                yield "group", group, self.groups[(department, group)]

                categories = sorted(
                    (
                        category
                        for dept, grp, category in self.categories
                        if dept == department and grp == group
                    ),
                    key=str.casefold,
                )
                for category in categories:
                    yield (
                        "category",
                        category,
                        self.categories[(department, group, category)],
                    )

        yield "grand", "Grand Total", self.grand_total


def decode_csv(path: Path) -> str:
    """Decode Excel-style CSV files, including Windows-1252 exports."""

    data = path.read_bytes()
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252")


XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XLSX_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def excel_column_index(reference: str) -> int:
    """Convert an Excel cell reference such as AA12 to a zero-based column."""

    letters = "".join(character for character in reference if character.isalpha())
    if not letters:
        raise ValueError(f"Invalid Excel cell reference: {reference!r}")
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - ord("A") + 1
    return number - 1


def read_shared_strings(workbook: zipfile.ZipFile) -> List[str]:
    """Load the optional shared-string table from an Excel workbook."""

    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ElementTree.fromstring(workbook.read("xl/sharedStrings.xml"))
    text_tag = f"{{{XLSX_MAIN_NS}}}t"
    return ["".join(node.text or "" for node in item.iter(text_tag)) for item in root]


def excel_cell_value(cell: ElementTree.Element, shared_strings: List[str]) -> str:
    """Return the stored value of an Excel cell as text."""

    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        text_tag = f"{{{XLSX_MAIN_NS}}}t"
        return "".join(node.text or "" for node in cell.iter(text_tag))

    value_node = cell.find(f"{{{XLSX_MAIN_NS}}}v")
    if value_node is None or value_node.text is None:
        return ""
    value = value_node.text

    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError) as error:
            raise ValueError(f"Invalid shared-string index: {value!r}") from error
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value


def resolve_worksheet_path(
    workbook: zipfile.ZipFile, sheet_name: Optional[str]
) -> Tuple[str, str]:
    """Resolve a worksheet name to its XML part inside the workbook."""

    workbook_root = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
    sheet_tag = f"{{{XLSX_MAIN_NS}}}sheet"
    sheets = list(workbook_root.iter(sheet_tag))
    if not sheets:
        raise ValueError("The input workbook contains no worksheets")

    if sheet_name is None:
        sheet = next((item for item in sheets if item.get("state") != "hidden"), sheets[0])
    else:
        sheet = next((item for item in sheets if item.get("name") == sheet_name), None)
        if sheet is None:
            sheet = next(
                (item for item in sheets if (item.get("name") or "").casefold() == sheet_name.casefold()),
                None,
            )
        if sheet is None:
            available = ", ".join(repr(item.get("name")) for item in sheets)
            raise ValueError(f"Worksheet {sheet_name!r} not found; available: {available}")

    relationship_id = sheet.get(f"{{{XLSX_DOC_REL_NS}}}id")
    relationships_root = ElementTree.fromstring(
        workbook.read("xl/_rels/workbook.xml.rels")
    )
    relationship_tag = f"{{{XLSX_PACKAGE_REL_NS}}}Relationship"
    targets = {
        item.get("Id"): item.get("Target")
        for item in relationships_root.iter(relationship_tag)
    }
    target = targets.get(relationship_id)
    if not target:
        raise ValueError(f"Cannot locate worksheet XML for {sheet.get('name')!r}")

    if target.startswith("/"):
        worksheet_path = target.lstrip("/")
    elif target.startswith("xl/"):
        worksheet_path = target
    else:
        worksheet_path = posixpath.normpath(posixpath.join("xl", target))
    return worksheet_path, sheet.get("name") or ""


def read_xlsx_table(
    path: Path, sheet_name: Optional[str]
) -> Tuple[List[str], List[Tuple[int, Dict[str, str]]], str]:
    """Read the first table-like worksheet from an .xlsx file."""

    try:
        workbook = zipfile.ZipFile(path)
    except zipfile.BadZipFile as error:
        raise ValueError(f"Not a valid .xlsx workbook: {path}") from error

    with workbook:
        worksheet_path, selected_sheet = resolve_worksheet_path(workbook, sheet_name)
        shared_strings = read_shared_strings(workbook)
        worksheet = ElementTree.fromstring(workbook.read(worksheet_path))

    row_tag = f"{{{XLSX_MAIN_NS}}}row"
    cell_tag = f"{{{XLSX_MAIN_NS}}}c"
    worksheet_rows: List[Tuple[int, Dict[int, str]]] = []
    for fallback_number, row in enumerate(worksheet.iter(row_tag), start=1):
        row_number = int(row.get("r") or fallback_number)
        values: Dict[int, str] = {}
        for position, cell in enumerate(row.findall(cell_tag)):
            reference = cell.get("r")
            column = excel_column_index(reference) if reference else position
            values[column] = excel_cell_value(cell, shared_strings)
        if any(value.strip() for value in values.values()):
            worksheet_rows.append((row_number, values))

    if not worksheet_rows:
        raise ValueError(f"Worksheet {selected_sheet!r} is empty")

    _header_row_number, header_cells = worksheet_rows[0]
    last_header_column = max(header_cells)
    fieldnames = [header_cells.get(column, "").strip() for column in range(last_header_column + 1)]
    named_fields = [field for field in fieldnames if field]
    if len(named_fields) != len(set(named_fields)):
        raise ValueError(f"Worksheet {selected_sheet!r} contains duplicate column names")

    records: List[Tuple[int, Dict[str, str]]] = []
    for row_number, cells in worksheet_rows[1:]:
        record = {
            field: cells.get(column, "")
            for column, field in enumerate(fieldnames)
            if field
        }
        records.append((row_number, record))
    return fieldnames, records, selected_sheet


def read_csv_table(path: Path) -> Tuple[List[str], List[Tuple[int, Dict[str, str]]], str]:
    """Read a CSV input for backward compatibility."""

    reader = csv.DictReader(io.StringIO(decode_csv(path), newline=""))
    if reader.fieldnames is None:
        raise ValueError("The input CSV has no header row")
    fieldnames = [str(field) for field in reader.fieldnames]
    records: List[Tuple[int, Dict[str, str]]] = []
    for row_number, row in enumerate(reader, start=2):
        normalized_row = {
            str(key): "" if value is None else str(value)
            for key, value in row.items()
            if key is not None
        }
        records.append((row_number, normalized_row))

    return fieldnames, records, "CSV"


def read_input_table(
    path: Path, sheet_name: Optional[str]
) -> Tuple[List[str], List[Tuple[int, Dict[str, str]]], str]:
    """Read a supported spreadsheet input."""

    suffix = path.suffix.casefold()
    if suffix == ".xlsx":
        return read_xlsx_table(path, sheet_name)
    if suffix == ".csv":
        return read_csv_table(path)
    raise ValueError("Input must have an .xlsx or .csv extension")


def parse_amount(value: Optional[str], field_name: str, row_number: int) -> Optional[Decimal]:
    """Convert common spreadsheet currency forms to Decimal; blanks remain None."""

    if value is None or not value.strip():
        return None

    text = value.strip()
    if text in {"-", "$-"}:
        return Decimal("0")

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = text.replace("$", "").replace(",", "").strip()

    try:
        amount = Decimal(text)
    except InvalidOperation as error:
        raise ValueError(
            f"Row {row_number}: {field_name!r} has invalid amount {value!r}"
        ) from error
    return -amount if negative else amount


def aggregate_records(
    fieldnames: Iterable[str],
    records: Iterable[Tuple[int, Dict[str, str]]],
    group_fields: Tuple[str, str, str],
) -> PivotData:
    """Validate and aggregate normalized spreadsheet records."""

    required = set(group_fields + VALUE_FIELDS)
    missing = sorted(required.difference(fieldnames))
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    pivot = PivotData()
    for row_number, row in records:
        labels = (
            (row.get(group_fields[0]) or "").strip(),
            (row.get(group_fields[1]) or "").strip(),
            (row.get(group_fields[2]) or "").strip(),
        )

        # Ignore report totals, filter notes, and completely blank rows because
        # none of them contain the three hierarchy values.
        if not any(labels):
            continue
        if not all(labels):
            raise ValueError(
                f"Row {row_number}: hierarchy fields must all be populated; got {labels!r}"
            )

        amounts = [
            parse_amount(row.get(field), field, row_number) for field in VALUE_FIELDS
        ]
        pivot.add(labels, amounts)

    return pivot


def load_pivot(
    input_path: Path,
    group_fields: Tuple[str, str, str] = DEFAULT_GROUP_FIELDS,
    sheet_name: Optional[str] = None,
) -> PivotData:
    """Read a spreadsheet and return its aggregated pivot data."""

    fieldnames, records, _selected_sheet = read_input_table(input_path, sheet_name)
    return aggregate_records(fieldnames, records, group_fields)


def load_totals(
    input_path: Path,
    group_fields: Tuple[str, str, str] = DEFAULT_GROUP_FIELDS,
    sheet_name: Optional[str] = None,
) -> Tuple[
    Dict[str, Totals],
    Dict[Tuple[str, str], Totals],
    Dict[Tuple[str, str, str], Totals],
    Totals,
    int,
]:
    """Return the legacy aggregate tuple for callers that import this script."""

    pivot = load_pivot(input_path, group_fields, sheet_name)
    return (
        pivot.departments,
        pivot.groups,
        pivot.categories,
        pivot.grand_total,
        pivot.detail_count,
    )


def format_amount(amount: Decimal, style: str) -> str:
    rounded = amount.quantize(Decimal("0.01"))
    if style == "numeric":
        return f"{rounded:.2f}"
    if rounded == 0:
        return "$-"
    if rounded < 0:
        return f"-${abs(rounded):,.2f}"
    return f"${rounded:,.2f}"


def output_row(label: str, totals: Totals, style: str) -> List[str]:
    return [label] + [
        format_amount(amount, style) if present else ""
        for amount, present in zip(totals.values, totals.has_value)
    ]


def write_csv(pivot: PivotData, output_file: TextIO, style: str) -> int:
    """Write aggregated pivot data as CSV."""

    writer = csv.writer(output_file, lineterminator="\n")
    writer.writerow(OUTPUT_HEADER)
    for _tier, label, totals in pivot.rows():
        writer.writerow(output_row(label, totals, style))
    return pivot.detail_count


def column_name(number: int) -> str:
    """Return an Excel column name for a one-based column number."""

    name = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def string_cell(reference: str, value: str, style: int) -> str:
    return (
        f'<c r="{reference}" s="{style}" t="inlineStr">'
        f"<is><t>{escape(value)}</t></is></c>"
    )


def number_cell(reference: str, value: Optional[Decimal], style: int) -> str:
    if value is None:
        return f'<c r="{reference}" s="{style}"/>'
    return f'<c r="{reference}" s="{style}"><v>{value}</v></c>'


def make_sheet_xml(rows: List[Tuple[str, str, Totals]]) -> str:
    """Build the worksheet XML used inside an .xlsx workbook."""

    sheet_rows: List[str] = []
    header_cells = "".join(
        string_cell(f"{column_name(column)}1", heading, 1)
        for column, heading in enumerate(OUTPUT_HEADER, start=1)
    )
    sheet_rows.append(
        f'<row r="1" ht="30" customHeight="1">{header_cells}</row>'
    )

    styles = {
        "department": (2, 3, 0),
        "group": (4, 5, 1),
        "category": (6, 7, 2),
        "grand": (8, 9, 0),
    }
    for row_number, (tier, label, totals) in enumerate(rows, start=2):
        label_style, amount_style, outline_level = styles[tier]
        cells = [string_cell(f"A{row_number}", label, label_style)]
        for column, (amount, present) in enumerate(
            zip(totals.values, totals.has_value), start=2
        ):
            cells.append(
                number_cell(
                    f"{column_name(column)}{row_number}",
                    amount if present else None,
                    amount_style,
                )
            )
        height = 22 if tier in {"department", "grand"} else 19
        sheet_rows.append(
            f'<row r="{row_number}" ht="{height}" customHeight="1" '
            f'outlineLevel="{outline_level}">{"".join(cells)}</row>'
        )

    last_row = len(rows) + 1
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><outlinePr summaryBelow="0"/></sheetPr>
  <dimension ref="A1:E{last_row}"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/>
  <cols>
    <col min="1" max="1" width="43" customWidth="1"/>
    <col min="2" max="5" width="23" customWidth="1"/>
  </cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:E{last_row}"/>
</worksheet>'''


def write_xlsx(pivot: PivotData, output_path: Path) -> int:
    """Write aggregated pivot data as a styled Excel workbook."""

    rows = list(pivot.rows())

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Pivot Table" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00;[Red]-&quot;$&quot;#,##0.00;&quot;$-&quot;"/></numFmts>
  <fonts count="3">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAF3F8"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left/><right/><top/><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment indent="1" vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment indent="2" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as workbook_file:
        workbook_file.writestr("[Content_Types].xml", content_types)
        workbook_file.writestr("_rels/.rels", root_rels)
        workbook_file.writestr("xl/workbook.xml", workbook)
        workbook_file.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook_file.writestr("xl/styles.xml", styles)
        workbook_file.writestr("xl/worksheets/sheet1.xml", make_sheet_xml(rows))
    return pivot.detail_count


def timestamped_path(path: Path, now: Optional[datetime] = None) -> Path:
    """Insert a local timestamp before a file's extension."""

    timestamp = (now or datetime.now()).strftime("%Y-%m-%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}{path.suffix}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert raw budget data to the supplied pivot-table layout."
    )
    parser.add_argument(
        "input",
        type=Path,
        nargs="?",
        default=DEFAULT_INPUT,
        help="input workbook path (default: data.xlsx beside this script)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=(
            "output .xlsx or .csv path "
            "(default: PowerBI_Pivot.xlsx beside this script)"
        ),
    )
    parser.add_argument(
        "--format",
        choices=("accounting", "numeric"),
        default="accounting",
        help="amount format for CSV output (default: accounting)",
    )
    parser.add_argument(
        "--group-columns",
        nargs=3,
        default=DEFAULT_GROUP_FIELDS,
        metavar=("TIER_1", "TIER_2", "TIER_3"),
        help=(
            "three source columns that define the hierarchy "
            "(default: 'Dept Name' 'Cost Group' 'Cost Category')"
        ),
    )
    parser.add_argument(
        "--sheet",
        help="worksheet name to read (default: first visible worksheet)",
    )
    parser.add_argument(
        "--no-timestamp",
        action="store_true",
        help="use the exact output filename instead of appending a timestamp",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    group_fields = tuple(args.group_columns)
    input_path = args.input.expanduser()
    requested_output = args.output.expanduser()

    if not input_path.is_file():
        raise SystemExit(f"Input workbook not found: {input_path}")
    if input_path.resolve() == requested_output.resolve():
        raise SystemExit("Input and output paths must be different")

    output_path = (
        requested_output
        if args.no_timestamp
        else timestamped_path(requested_output)
    )
    suffix = output_path.suffix.casefold()
    if suffix not in {".xlsx", ".csv"}:
        raise SystemExit("Output must have an .xlsx or .csv extension")

    pivot = load_pivot(input_path, group_fields, args.sheet)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if suffix == ".xlsx":
        detail_count = write_xlsx(pivot, output_path)
    else:
        with output_path.open("w", encoding="utf-8-sig", newline="") as output_file:
            detail_count = write_csv(pivot, output_file, args.format)

    print(f"Wrote {output_path} from {detail_count} detail rows")


if __name__ == "__main__":
    main()
