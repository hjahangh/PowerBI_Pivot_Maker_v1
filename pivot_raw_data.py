#!/usr/bin/env python3
"""The three source columns can be changed with --group-columns. Group values are
always discovered from the current data; no department, group, or category
names are hard-coded.

Only Python's standard library is required.
"""

from __future__ import annotations

import argparse
import csv
import io
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, TextIO, Tuple
from xml.sax.saxutils import escape


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "data.csv"
DEFAULT_OUTPUT = SCRIPT_DIR / "pivoted_report.xlsx"
DEFAULT_GROUP_FIELDS: Tuple[str, ...] = (
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


def decode_csv(path: Path) -> str:
    """Decode Excel-style CSV files, including Windows-1252 exports."""

    data = path.read_bytes()
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252")


def parse_amount(value: Optional[str], field_name: str, row_number: int) -> Optional[Decimal]:
    """Convert common CSV currency forms to Decimal; blanks remain None."""

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


def load_totals(
    input_path: Path,
    group_fields: Tuple[str, str, str] = DEFAULT_GROUP_FIELDS,
) -> Tuple[
    Dict[str, Totals],
    Dict[Tuple[str, str], Totals],
    Dict[Tuple[str, str, str], Totals],
    Totals,
    int,
]:
    """Read detail rows and aggregate all three hierarchy levels."""

    reader = csv.DictReader(io.StringIO(decode_csv(input_path), newline=""))
    if reader.fieldnames is None:
        raise ValueError("The input CSV has no header row")

    required = set(group_fields + VALUE_FIELDS)
    missing = sorted(required.difference(reader.fieldnames))
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    department_totals: Dict[str, Totals] = {}
    group_totals: Dict[Tuple[str, str], Totals] = {}
    category_totals: Dict[Tuple[str, str, str], Totals] = {}
    grand_total = Totals()
    detail_count = 0

    for row_number, row in enumerate(reader, start=2):
        labels = tuple((row.get(field) or "").strip() for field in group_fields)

        # Ignore completely blank lines in otherwise valid CSV files.
        if not any(labels):
            continue
        if not all(labels):
            raise ValueError(
                f"Row {row_number}: hierarchy fields must all be populated; got {labels!r}"
            )

        amounts = [
            parse_amount(row.get(field), field, row_number) for field in VALUE_FIELDS
        ]
        department, group, category = labels

        department_totals.setdefault(department, Totals()).add(amounts)
        group_totals.setdefault((department, group), Totals()).add(amounts)
        category_totals.setdefault((department, group, category), Totals()).add(amounts)
        grand_total.add(amounts)
        detail_count += 1

    return department_totals, group_totals, category_totals, grand_total, detail_count


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


def pivot_rows(
    department_totals: Dict[str, Totals],
    group_totals: Dict[Tuple[str, str], Totals],
    category_totals: Dict[Tuple[str, str, str], Totals],
    grand_total: Totals,
) -> Iterator[Tuple[str, str, Totals]]:
    """Yield each pivot row as (tier, label, totals)."""

    for department in sorted(department_totals, key=str.casefold):
        yield "department", department, department_totals[department]

        groups = sorted(
            (group for dept, group in group_totals if dept == department),
            key=str.casefold,
        )
        for group in groups:
            yield "group", group, group_totals[(department, group)]

            categories = sorted(
                (
                    category
                    for dept, grp, category in category_totals
                    if dept == department and grp == group
                ),
                key=str.casefold,
            )
            for category in categories:
                yield (
                    "category",
                    category,
                    category_totals[(department, group, category)],
                )

    yield "grand", "Grand Total", grand_total


def write_csv(
    input_path: Path,
    output_file: TextIO,
    style: str,
    group_fields: Tuple[str, str, str] = DEFAULT_GROUP_FIELDS,
) -> int:
    """Write the pivot as a CSV compatible with the supplied example."""

    (
        department_totals,
        group_totals,
        category_totals,
        grand_total,
        detail_count,
    ) = load_totals(input_path, group_fields)

    writer = csv.writer(output_file, lineterminator="\n")
    writer.writerow(OUTPUT_HEADER)

    rows = pivot_rows(
        department_totals, group_totals, category_totals, grand_total
    )
    for _tier, label, totals in rows:
        writer.writerow(output_row(label, totals, style))
    return detail_count


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


def write_xlsx(
    input_path: Path,
    output_path: Path,
    group_fields: Tuple[str, str, str] = DEFAULT_GROUP_FIELDS,
) -> int:
    """Write a styled Excel workbook without requiring third-party packages."""

    (
        department_totals,
        group_totals,
        category_totals,
        grand_total,
        detail_count,
    ) = load_totals(input_path, group_fields)
    rows = list(
        pivot_rows(
            department_totals, group_totals, category_totals, grand_total
        )
    )

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
    return detail_count


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
        help="input CSV path (default: data.csv beside this script)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=(
            "output .xlsx or .csv path "
            "(default: pivoted_report.xlsx beside this script)"
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
        raise SystemExit(f"Input CSV not found: {input_path}")
    if input_path.resolve() == requested_output.resolve():
        raise SystemExit("Input and output paths must be different")

    output_path = (
        requested_output
        if args.no_timestamp
        else timestamped_path(requested_output)
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.casefold()
    if suffix == ".xlsx":
        detail_count = write_xlsx(input_path, output_path, group_fields)
    elif suffix == ".csv":
        with output_path.open("w", encoding="utf-8-sig", newline="") as output_file:
            detail_count = write_csv(
                input_path, output_file, args.format, group_fields
            )
    else:
        raise SystemExit("Output must have an .xlsx or .csv extension")
    print(f"Wrote {output_path} from {detail_count} detail rows")


if __name__ == "__main__":
    main()
    
