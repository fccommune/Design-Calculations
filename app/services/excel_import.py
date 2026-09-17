"""
Parses an uploaded valve BOM workbook into a list of row dicts, same shape
as the old WinForms app's LoadDataFromExcel (Form2.cs): first row = headers
(blank/duplicate headers get "ColumnN" / "_2", "_3" ... suffixes), every
row after that = one dict of {header: value}.
"""
import openpyxl


class ExcelImportError(Exception):
    pass


def _cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _header_str(value) -> str:
    """Unlike _cell_str, this does NOT strip whitespace: some of the
    original app's expected headers have a deliberate trailing space (e.g.
    "Valve_Size ") that the lookup code matches on exactly - see
    app/services/gad_globe_lookup.py. Only used to decide blankness."""
    return "" if value is None else str(value)


def _unique_headers(raw_headers):
    used = set()
    headers = []
    for i, raw in enumerate(raw_headers, start=1):
        raw_name = _header_str(raw)
        name = raw_name if raw_name.strip() else f"Column{i}"
        final = name
        suffix = 2
        while final.lower() in used:
            final = f"{name}_{suffix}"
            suffix += 1
        used.add(final.lower())
        headers.append(final)
    return headers


def parse_excel_rows(file_stream) -> tuple:
    """Returns (headers: list[str], rows: list[dict]).

    Raises ExcelImportError on anything that isn't a readable workbook.
    """
    try:
        workbook = openpyxl.load_workbook(file_stream, data_only=True, read_only=True)
    except Exception as exc:
        raise ExcelImportError(f"Failed to load Excel: {exc}") from exc

    worksheet = workbook.worksheets[0]
    rows_iter = worksheet.iter_rows(values_only=True)

    try:
        raw_header_row = next(rows_iter)
    except StopIteration:
        raise ExcelImportError("The workbook's first sheet is empty.")

    headers = _unique_headers(raw_header_row)

    rows = []
    for raw_row in rows_iter:
        if raw_row is None or all(v is None for v in raw_row):
            continue
        row = {}
        for i, header in enumerate(headers):
            value = raw_row[i] if i < len(raw_row) else None
            row[header] = "" if value is None else value
        rows.append(row)

    workbook.close()
    return headers, rows
