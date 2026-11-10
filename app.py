from __future__ import annotations

import io
import re
import zipfile
from bisect import bisect_left
from copy import copy
from pathlib import Path

import msoffcrypto
import streamlit as st
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries


DESTINATION_CODES = ("EZE", "GUA", "SAL", "LIM", "SJO", "VCP", "SCL", "POS", "SAP", "UIO", "KIN")
CLIENTS = ("TEMU", "JTExpress", "BDT")
OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
SENSITIVE_TERMS = (
    "ALCOHOL",
    "ACETONE",
    "ESSENTIAL OIL",
    "PERFUME",
    "PERFUMING",
    "NAIL ART",
    "NAIL POLISH",
)
TERM_PATTERNS = tuple(
    (
        term,
        re.compile(
            r"(?<![A-Z0-9])" + r"[\s\u00a0]+".join(map(re.escape, term.split())) + r"(?![A-Z0-9])",
            re.IGNORECASE,
        ),
    )
    for term in SENSITIVE_TERMS
)


class ProcessingError(Exception):
    """An expected, user-correctable workbook or input error."""


def password_mapping() -> dict[str, str]:
    """Read destination passwords from Streamlit Secrets, never from source code."""
    if "manifest_passwords" not in st.secrets:
        return {}
    try:
        section = st.secrets["manifest_passwords"]
        return {str(code).strip().upper(): str(password) for code, password in section.items()}
    except (KeyError, AttributeError, TypeError):
        return {}


def decrypt_if_needed(
    file_bytes: bytes,
    file_suffix: str,
    destination_code: str,
    passwords: dict[str, str],
) -> bytes:
    if file_bytes.startswith(OLE_SIGNATURE):
        try:
            encrypted = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
            is_encrypted = encrypted.is_encrypted()
        except Exception:
            if file_suffix == ".xls":
                return file_bytes
            raise ProcessingError("This file is not a supported encrypted Excel workbook.")

        if not is_encrypted:
            if file_suffix == ".xls":
                return file_bytes
            raise ProcessingError("This file is not a supported encrypted Excel workbook.")

        if destination_code not in DESTINATION_CODES:
            raise ProcessingError("This file requires an opening password. Enter the matching three-character destination code, such as SAL.")
        password = passwords.get(destination_code)
        if not password:
            raise ProcessingError(
                "No password is configured for this destination. Add its password to the [manifest_passwords] section in Streamlit Secrets and try again."
            )

        try:
            encrypted.load_key(password=password, verify_password=True)
            decrypted = io.BytesIO()
            encrypted.decrypt(decrypted)
            return decrypted.getvalue()
        except ProcessingError:
            raise
        except Exception as exc:
            raise ProcessingError("Unable to open the file. Check the three-character destination code and confirm its password is configured in Secrets.") from exc

    if not zipfile.is_zipfile(io.BytesIO(file_bytes)):
        raise ProcessingError("Only .xlsx, .xlsm, and .xls files are supported. Upload a valid Excel workbook.")
    return file_bytes


def normalized_header(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def find_client_headers(ws, client: str) -> tuple[int, int, int] | None:
    for row_number in range(1, min(ws.max_row, 100) + 1):
        headers = {
            normalized_header(ws.cell(row=row_number, column=column).value): column
            for column in range(1, ws.max_column + 1)
            if normalized_header(ws.cell(row=row_number, column=column).value)
        }

        if client == "TEMU":
            if (
                normalized_header(ws.cell(row=row_number, column=3).value) == "bg number"
                and normalized_header(ws.cell(row=row_number, column=15).value) == "product description"
            ):
                return row_number, 3, 15
            continue

        if client == "JTExpress":
            dedupe_header = "tracking no."
            description_headers = ("content description", "product description")
        elif client == "BDT":
            dedupe_header = "tracking number"
            description_headers = ("product description", "content description")
        else:
            raise ProcessingError(f"Unsupported client: {client}")

        dedupe_column = headers.get(dedupe_header)
        description_column = next((headers[name] for name in description_headers if name in headers), None)
        if dedupe_column is not None and description_column is not None:
            return row_number, dedupe_column, description_column
    return None


def remove_blank_rows(ws) -> int:
    populated_rows = sorted(
        {
            cell.row
            for cell in ws._cells.values()
            if not isinstance(cell, MergedCell)
            and cell.value is not None
            and (not isinstance(cell.value, str) or cell.value.strip())
        }
    )
    if not populated_rows:
        return 0

    blank_ranges: list[tuple[int, int]] = []
    if populated_rows[0] > 1:
        blank_ranges.append((1, populated_rows[0] - 1))
    for previous_row, next_row in zip(populated_rows, populated_rows[1:]):
        if next_row > previous_row + 1:
            blank_ranges.append((previous_row + 1, next_row - 1))
    if ws.max_row > populated_rows[-1]:
        blank_ranges.append((populated_rows[-1] + 1, ws.max_row))
    if not blank_ranges:
        return 0

    first_blank_row = blank_ranges[0][0]
    if any(cell.data_type == "f" for cell in ws._cells.values()):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains formulas. Convert them to values before removing blank rows so formula references remain correct."
        )
    if any(merged.max_row >= first_blank_row for merged in ws.merged_cells.ranges):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains merged cells at or below blank rows. Unmerge them before processing."
        )
    if ws.conditional_formatting or ws.data_validations.dataValidation:
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains conditional formatting or data validation. Remove these rules before removing blank rows."
        )

    range_starts = [start for start, _ in blank_ranges]
    range_ends = [end for _, end in blank_ranges]
    cumulative_removed: list[int] = []
    total_removed = 0
    for start, end in blank_ranges:
        total_removed += end - start + 1
        cumulative_removed.append(total_removed)

    def remap_row(row_number: int) -> int | None:
        range_index = bisect_left(range_starts, row_number)
        if range_index < len(blank_ranges) and range_starts[range_index] == row_number:
            return None
        previous_range = range_index - 1
        if previous_range >= 0 and row_number <= range_ends[previous_range]:
            return None
        removed_before = cumulative_removed[previous_range] if previous_range >= 0 else 0
        return row_number - removed_before

    compacted_cells = {}
    for (old_row, column), cell in ws._cells.items():
        new_row = remap_row(old_row)
        if new_row is None:
            continue
        if new_row != old_row:
            cell.row = new_row
        compacted_cells[(new_row, column)] = cell
    ws._cells = compacted_cells

    old_dimensions = list(ws.row_dimensions.items())
    ws.row_dimensions.clear()
    for old_row, dimension in old_dimensions:
        new_row = remap_row(old_row)
        if new_row is None:
            continue
        dimension.index = new_row
        ws.row_dimensions[new_row] = dimension

    if ws.auto_filter.ref:
        min_col, min_row, max_col, max_row = range_boundaries(ws.auto_filter.ref)
        new_min_row = remap_row(min_row)
        new_max_row = remap_row(max_row)
        if new_min_row is not None and new_max_row is not None:
            ws.auto_filter.ref = (
                f"{get_column_letter(min_col)}{new_min_row}:"
                f"{get_column_letter(max_col)}{new_max_row}"
            )
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        new_min_row = remap_row(min_row)
        new_max_row = remap_row(max_row)
        if new_min_row is not None and new_max_row is not None:
            table.ref = (
                f"{get_column_letter(min_col)}{new_min_row}:"
                f"{get_column_letter(max_col)}{new_max_row}"
            )
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
    return total_removed


def remove_duplicate_rows(ws, header_row: int, key_column: int) -> int:
    seen: set[str] = set()
    rows_to_remove: list[int] = []
    for row_number in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row=row_number, column=key_column).value
        if value is None or not str(value).strip():
            continue
        key = str(value).strip().casefold()
        if key in seen:
            rows_to_remove.append(row_number)
        else:
            seen.add(key)

    if not rows_to_remove:
        return 0

    if any(cell.data_type == "f" and cell.row > header_row for cell in ws._cells.values()):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains formulas in its data area. Convert them to values before processing to avoid shifting formula references."
        )
    if any(merged.max_row > header_row for merged in ws.merged_cells.ranges):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains merged cells in its data area, so duplicate rows cannot be removed safely."
        )
    if ws.conditional_formatting or ws.data_validations.dataValidation:
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains conditional formatting or data validation. Remove these rules before processing to prevent their references from shifting."
        )

    def remap_row(row_number: int) -> int | None:
        removed_before = bisect_left(rows_to_remove, row_number)
        if removed_before < len(rows_to_remove) and rows_to_remove[removed_before] == row_number:
            return None
        return row_number - removed_before if row_number > header_row else row_number

    # Compact the worksheet once instead of shifting every later row once per duplicate.
    compacted_cells = {}
    for (old_row, column), cell in ws._cells.items():
        new_row = remap_row(old_row)
        if new_row is None:
            continue
        if new_row != old_row:
            cell.row = new_row
        compacted_cells[(new_row, column)] = cell
    ws._cells = compacted_cells

    old_dimensions = list(ws.row_dimensions.items())
    ws.row_dimensions.clear()
    for old_row, dimension in old_dimensions:
        new_row = remap_row(old_row)
        if new_row is None:
            continue
        dimension.index = new_row
        ws.row_dimensions[new_row] = dimension

    if ws.auto_filter.ref:
        min_col, min_row, max_col, max_row = range_boundaries(ws.auto_filter.ref)
        new_min_row = remap_row(min_row)
        new_max_row = remap_row(max_row)
        if new_min_row is not None and new_max_row is not None:
            ws.auto_filter.ref = (
                f"{get_column_letter(min_col)}{new_min_row}:"
                f"{get_column_letter(max_col)}{new_max_row}"
            )
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        new_min_row = remap_row(min_row)
        new_max_row = remap_row(max_row)
        if new_min_row is not None and new_max_row is not None:
            table.ref = (
                f"{get_column_letter(min_col)}{new_min_row}:"
                f"{get_column_letter(max_col)}{new_max_row}"
            )
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
    return len(rows_to_remove)


def remove_header_only_columns(ws, header_row: int, protected_columns: set[int]) -> int:
    columns_to_remove = []
    for column in range(1, ws.max_column + 1):
        header_value = ws.cell(row=header_row, column=column).value
        if column in protected_columns or header_value is None or not str(header_value).strip():
            continue
        if all(
            ws.cell(row=row_number, column=column).value is None
            or not str(ws.cell(row=row_number, column=column).value).strip()
            for row_number in range(header_row + 1, ws.max_row + 1)
        ):
            columns_to_remove.append(column)

    if not columns_to_remove:
        return 0

    first_removed_column = columns_to_remove[0]
    if any(cell.data_type == "f" for cell in ws._cells.values()):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains formulas. Convert them to values before removing header-only columns."
        )
    if any(merged.max_col >= first_removed_column for merged in ws.merged_cells.ranges):
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains merged cells at or after columns being removed. Unmerge them before processing."
        )
    if ws.conditional_formatting or ws.data_validations.dataValidation:
        raise ProcessingError(
            f"Worksheet '{ws.title}' contains conditional formatting or data validation. Remove these rules before removing columns."
        )

    removed_set = set(columns_to_remove)

    def remap_column(column: int) -> int | None:
        removed_before = bisect_left(columns_to_remove, column)
        if removed_before < len(columns_to_remove) and columns_to_remove[removed_before] == column:
            return None
        return column - removed_before

    # Compact cells once, preserving their styles while removing every header-only column.
    compacted_cells = {}
    for (row_number, old_column), cell in ws._cells.items():
        new_column = remap_column(old_column)
        if new_column is None:
            continue
        if new_column != old_column:
            cell.column = new_column
        compacted_cells[(row_number, new_column)] = cell
    ws._cells = compacted_cells

    old_dimensions = list(ws.column_dimensions.items())
    ws.column_dimensions.clear()
    for key, dimension in old_dimensions:
        first_column = dimension.min or column_index_from_string(key)
        last_column = dimension.max or first_column
        for old_column in range(first_column, last_column + 1):
            new_column = remap_column(old_column)
            if new_column is None:
                continue
            new_dimension = copy(dimension)
            new_dimension.min = new_column
            new_dimension.max = new_column
            new_dimension.index = get_column_letter(new_column)
            ws.column_dimensions[new_dimension.index] = new_dimension

    def remap_reference(reference: str | None) -> str | None:
        if not reference:
            return reference
        min_column, min_row, max_column, max_row = range_boundaries(reference)
        retained_columns = [
            remap_column(column)
            for column in range(min_column, max_column + 1)
            if column not in removed_set
        ]
        retained_columns = [column for column in retained_columns if column is not None]
        if not retained_columns:
            return None
        return (
            f"{get_column_letter(min(retained_columns))}{min_row}:"
            f"{get_column_letter(max(retained_columns))}{max_row}"
        )

    if ws.auto_filter.ref:
        ws.auto_filter.ref = remap_reference(ws.auto_filter.ref)
    for table in ws.tables.values():
        table.ref = remap_reference(table.ref)
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref
    return len(columns_to_remove)


def remove_sensitive_terms(ws, header_row: int, description_column: int) -> dict[str, int]:
    counts = {term: 0 for term in SENSITIVE_TERMS}
    for row_number in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(row=row_number, column=description_column)
        if not isinstance(cell.value, str) or cell.data_type == "f":
            continue
        cleaned = cell.value
        for term, pattern in TERM_PATTERNS:
            matches = pattern.findall(cleaned)
            counts[term] += len(matches)
            cleaned = pattern.sub("", cleaned)
        cleaned = re.sub(r"[ \t]{2,}", " ", cleaned).strip()
        cell.value = cleaned
    return counts


def load_legacy_xls(file_bytes: bytes) -> Workbook:
    try:
        source_book = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
    except xlrd.XLRDError as exc:
        raise ProcessingError(
            "Unable to read this .xls workbook. If it is password-protected, open it in Excel and save an unprotected .xlsx copy first."
        ) from exc

    workbook = Workbook()
    workbook.remove(workbook.active)

    for source_sheet in source_book.sheets():
        sheet_title = source_sheet.name[:31] or "Sheet"
        ws = workbook.create_sheet(title=sheet_title)
        if getattr(source_sheet, "visibility", 0) == 1:
            ws.sheet_state = "hidden"

        for row_index in range(source_sheet.nrows):
            for column_index in range(source_sheet.ncols):
                source_cell = source_sheet.cell(row_index, column_index)
                if source_cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue

                if source_cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(source_cell.value, source_book.datemode)
                elif source_cell.ctype == xlrd.XL_CELL_ERROR:
                    value = str(source_cell.value)
                else:
                    value = source_cell.value

                target_cell = ws.cell(row=row_index + 1, column=column_index + 1, value=value)
                xf = source_book.xf_list[source_cell.xf_index]
                cell_format = source_book.format_map.get(xf.format_key)
                if cell_format is not None:
                    target_cell.number_format = cell_format.format_str

        for column_index, column_info in source_sheet.colinfo_map.items():
            dimension = ws.column_dimensions[get_column_letter(column_index + 1)]
            dimension.width = column_info.width / 256
            dimension.hidden = bool(column_info.hidden)

        for row_index, row_info in source_sheet.rowinfo_map.items():
            dimension = ws.row_dimensions[row_index + 1]
            if row_info.height != 0xFFFF:
                dimension.height = row_info.height / 20
            dimension.hidden = bool(row_info.hidden)

        for row_start, row_end, column_start, column_end in source_sheet.merged_cells:
            ws.merge_cells(
                start_row=row_start + 1,
                end_row=row_end,
                start_column=column_start + 1,
                end_column=column_end,
            )

    if not workbook.worksheets:
        workbook.create_sheet("Sheet1")
    return workbook


def apply_font_size_and_print_settings(workbook) -> None:
    font_ids = {}
    for old_font_id, font in enumerate(tuple(workbook._fonts)):
        updated_font = copy(font)
        updated_font.sz = 16
        font_ids[old_font_id] = workbook._fonts.add(updated_font)

    for ws in workbook.worksheets:
        updated_styles = {}
        default_style = workbook._cell_styles[0]
        for cell in ws._cells.values():
            if isinstance(cell, MergedCell):
                continue
            original_style = cell._style or default_style
            style_key = tuple(original_style)
            updated_style = updated_styles.get(style_key)
            if updated_style is None:
                updated_style = copy(original_style)
                updated_style.fontId = font_ids[original_style.fontId]
                updated_styles[style_key] = updated_style
            cell._style = updated_style

        setup = ws.sheet_properties.pageSetUpPr
        if setup is None:
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
        else:
            setup.fitToPage = True
            setup.autoPageBreaks = False

        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75
        ws.page_margins.header = 0.3
        ws.page_margins.footer = 0.3

        if ws.max_row and ws.max_column:
            ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"


def process_workbook(
    file_bytes: bytes,
    file_name: str,
    client: str,
    destination_code: str,
    passwords: dict[str, str],
) -> tuple[bytes, str, dict[str, dict[str, object]]]:
    suffix = Path(file_name).suffix.lower()
    if suffix not in (".xlsx", ".xlsm", ".xls"):
        raise ProcessingError("Only .xlsx, .xlsm, and .xls files are supported.")
    if client not in CLIENTS:
        raise ProcessingError(f"Unsupported client: {client}")

    workbook_bytes = decrypt_if_needed(file_bytes, suffix, destination_code, passwords)
    if suffix == ".xls":
        workbook = load_legacy_xls(workbook_bytes)
    else:
        try:
            workbook = load_workbook(
                io.BytesIO(workbook_bytes),
                data_only=False,
                keep_vba=suffix == ".xlsm",
            )
        except Exception as exc:
            raise ProcessingError("Unable to read the Excel workbook. Confirm that the file is not damaged and is in .xlsx or .xlsm format.") from exc

    summaries: dict[str, dict[str, object]] = {}
    matching_sheets = []
    for ws in workbook.worksheets:
        headers = find_client_headers(ws, client)
        if headers is None:
            continue
        matching_sheets.append((ws, *headers))

    if not matching_sheets:
        expected_headers = {
            "TEMU": "column C = BG Number and column O = PRODUCT DESCRIPTION",
            "JTExpress": "TRACKING NO. and CONTENT DESCRIPTION (or PRODUCT DESCRIPTION)",
            "BDT": "Tracking Number and PRODUCT DESCRIPTION",
        }[client]
        raise ProcessingError(
            f"Could not find the expected {client} headers. Confirm that one row contains {expected_headers}."
        )

    for ws, header_row, key_column, description_column in matching_sheets:
        blank_rows_removed = remove_blank_rows(ws)
        refreshed_headers = find_client_headers(ws, client)
        if refreshed_headers is None:
            raise ProcessingError(f"Could not find the expected headers after removing blank rows from worksheet '{ws.title}'.")
        header_row, key_column, description_column = refreshed_headers

        columns_removed = 0
        if client in ("JTExpress", "BDT"):
            columns_removed = remove_header_only_columns(
                ws,
                header_row,
                protected_columns={key_column, description_column},
            )
            refreshed_headers = find_client_headers(ws, client)
            if refreshed_headers is None:
                raise ProcessingError(f"Could not find the expected headers after removing empty columns from worksheet '{ws.title}'.")
            header_row, key_column, description_column = refreshed_headers

        duplicates_removed = remove_duplicate_rows(ws, header_row, key_column)
        terms_removed = remove_sensitive_terms(ws, header_row, description_column)
        summaries[ws.title] = {
            "Blank rows removed": blank_rows_removed,
            "Columns removed": columns_removed,
            "Duplicate rows removed": duplicates_removed,
            "Sensitive term removals": terms_removed,
        }

    apply_font_size_and_print_settings(workbook)
    try:
        if hasattr(workbook, "calculation"):
            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
            workbook.calculation.calcMode = "auto"
    except (AttributeError, TypeError):
        pass

    output = io.BytesIO()
    try:
        workbook.save(output)
    except Exception as exc:
        raise ProcessingError("Unable to save the processed Excel file. Confirm that the file is not damaged and try again.") from exc

    output_bytes = output.getvalue()
    extension = ".xlsm" if suffix == ".xlsm" else ".xlsx"
    return output_bytes, extension, summaries


def safe_output_name(awb: str, extension: str) -> str:
    safe_awb = re.sub(r"[^A-Za-z0-9._-]+", "", awb.strip()).strip("._-")
    if not safe_awb:
        raise ProcessingError("Enter a valid export AWB#.")
    return f"{safe_awb}manifest{extension}"


def main() -> None:
    st.set_page_config(page_title="Manifest Excel Cleaner", page_icon="📄", layout="centered")
    st.title("Manifest Excel Cleaner")
    st.write("Choose a client, upload its manifest, enter a destination code if the workbook is password-protected, and download the cleaned result.")

    client = st.selectbox("Client", options=CLIENTS)
    uploaded_file = st.file_uploader("Upload a manifest (.xlsx, .xlsm, or .xls)", type=["xlsx", "xlsm", "xls"])
    destination_code = st.text_input("Three-character destination code", max_chars=3, placeholder="e.g. SAL").strip().upper()
    awb = st.text_input("Export AWB#", placeholder="e.g. 810-51887706").strip()

    if st.button("Process and generate file", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.error("Upload a manifest file first.")
        elif not awb:
            st.error("Enter an export AWB#.")
        elif destination_code and len(destination_code) != 3:
            st.error("The destination code must contain exactly three characters.")
        else:
            try:
                output_bytes, extension, summaries = process_workbook(
                    uploaded_file.getvalue(),
                    uploaded_file.name,
                    client,
                    destination_code,
                    password_mapping(),
                )
                output_name = safe_output_name(awb, extension)
                st.success("Processing complete. The output workbook is not protected with an opening password.")
                for sheet_name, summary in summaries.items():
                    st.subheader(sheet_name)
                    st.write(f"Blank rows removed: {summary['Blank rows removed']}")
                    if summary["Columns removed"]:
                        st.write(f"Header-only columns removed: {summary['Columns removed']}")
                    st.write(f"Duplicate rows removed: {summary['Duplicate rows removed']}")
                    term_counts = summary["Sensitive term removals"]
                    st.write("Sensitive term removals: " + "; ".join(f"{term}: {count}" for term, count in term_counts.items()))
                st.download_button(
                    label="Download processed manifest",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                        if extension == ".xlsm"
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )
            except ProcessingError as exc:
                st.error(str(exc))
            except Exception:
                st.error("Processing failed. Check the file and your inputs, then try again.")

    st.caption(
        "Print settings: narrow margins, landscape orientation, all columns fit on one page, and 16 pt font. "
        "Duplex printing is controlled by the printer driver; select two-sided printing in the print dialog."
    )


if __name__ == "__main__":
    main()
