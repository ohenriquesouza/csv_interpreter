import csv
import os
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from settings import DEFAULT_THEME

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

NULL_FILL = PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid")
ALERT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
THIN = Side(style="thin")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_IMAGE_NAME_KEYWORDS = {"imagem", "foto", "image", "photo", "figura", "img", "picture", "thumb", "thumbnail"}
_IMAGE_EXT_RE = re.compile(r'\.(jpg|jpeg|png|gif|webp|bmp|tiff|svg)(\?.*)?$', re.IGNORECASE)


# --- Theme utilities ---

def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"{r:02X}{g:02X}{b:02X}"


def _relative_luminance(r: int, g: int, b: int) -> float:
    def _ch(x: int) -> float:
        v = x / 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * _ch(r) + 0.7152 * _ch(g) + 0.0722 * _ch(b)


def _lighten(r: int, g: int, b: int, amount: float) -> tuple[int, int, int]:
    return (
        min(255, int(r + (255 - r) * amount)),
        min(255, int(g + (255 - g) * amount)),
        min(255, int(b + (255 - b) * amount)),
    )


def _darken(r: int, g: int, b: int, amount: float) -> tuple[int, int, int]:
    return (
        max(0, int(r * (1 - amount))),
        max(0, int(g * (1 - amount))),
        max(0, int(b * (1 - amount))),
    )


def build_theme(hex_color: str) -> dict:
    r, g, b = _hex_to_rgb(hex_color)
    lum = _relative_luminance(r, g, b)

    header_hex = _rgb_to_hex(r, g, b)

    lr, lg, lb = _lighten(r, g, b, 0.72)
    footer_hex = _rgb_to_hex(lr, lg, lb)

    header_text = "FFFFFF" if lum < 0.35 else "1A1A1A"

    dr, dg, db = _darken(r, g, b, 0.45)
    footer_text = _rgb_to_hex(dr, dg, db)

    return {
        "header_fill": PatternFill(start_color=header_hex, end_color=header_hex, fill_type="solid"),
        "footer_fill": PatternFill(start_color=footer_hex, end_color=footer_hex, fill_type="solid"),
        "header_text": header_text,
        "footer_text": footer_text,
    }


# --- Data helpers ---

def _is_null(value: str) -> bool:
    return str(value).strip().lower() in ("", "nan", "none", "null", "$null$", "n/a", "na")


def _read_csv(csv_path: str) -> tuple[list[str], list[list[str]]]:
    # Match the previous reader's text handling, including UTF-8 BOMs, quoted
    # newlines, blank lines, duplicate headers and missing trailing fields.
    csv.field_size_limit(2**31 - 1)
    with open(csv_path, encoding="utf-8-sig", newline="") as source:
        record_lines = []

        def lines():
            for line in source:
                record_lines.append(line)
                yield line

        def records():
            reader = csv.reader(lines(), strict=True)
            for row in reader:
                blank = not "".join(record_lines).strip()
                record_lines.clear()
                if not blank:
                    yield row

        rows = records()
        raw_headers = next(rows, None)
        if raw_headers is None:
            raise ValueError("O arquivo CSV está vazio ou não contém colunas.")

        headers = [name or f"Unnamed: {i}" for i, name in enumerate(raw_headers)]
        reserved = set(headers)
        used = set()
        # Named columns take precedence over generated names for empty headers.
        for i in sorted(range(len(headers)), key=lambda i: not bool(raw_headers[i])):
            name = headers[i]
            if name in used:
                suffix = 1
                while f"{name}.{suffix}" in reserved or f"{name}.{suffix}" in used:
                    suffix += 1
                name = f"{name}.{suffix}"
            headers[i] = name
            used.add(name)

        data = []
        index_columns = None
        for row in rows:
            if index_columns is None:
                # The previous reader treated extra leading fields in the first
                # record as an implicit index, which is not exported to Excel.
                index_columns = max(0, len(row) - len(headers))
            expected = len(headers) + index_columns
            if len(row) > expected:
                raise ValueError(
                    f"Linha CSV com {len(row)} campos; eram esperados {expected}."
                )
            row += [""] * (expected - len(row))
            data.append(row[index_columns:])
    return headers, data


def _is_image_col(col_name: str, values) -> bool:
    if any(kw in col_name.lower() for kw in _IMAGE_NAME_KEYWORDS):
        return True
    non_null = [value for value in values if not _is_null(value)]
    return bool(non_null) and sum(
        bool(_IMAGE_EXT_RE.search(value)) for value in non_null
    ) > len(non_null) / 2


def _reorder_image_cols(
    headers: list[str], rows: list[list[str]],
) -> tuple[list[str], list[list[str]]]:
    image_cols = [
        i for i, name in enumerate(headers)
        if _is_image_col(name, (row[i] for row in rows))
    ]
    if not image_cols:
        return headers, rows
    image_set = set(image_cols)
    order = [i for i in range(len(headers)) if i not in image_set] + image_cols
    return [headers[i] for i in order], [[row[i] for i in order] for row in rows]


def _find_null_alerts(
    headers: list[str], rows: list[list[str]], threshold: int = 10,
) -> list[str]:
    id_index = next((i for i, name in enumerate(headers) if "id" in name.lower()), None)
    if id_index is None:
        return []

    id_col = headers[id_index]
    counts_by_id = {}
    for row in rows:
        counts = counts_by_id.setdefault(row[id_index], [0] * len(headers))
        for i, value in enumerate(row):
            if i != id_index and _is_null(value):
                counts[i] += 1
    alerts = []
    for unit_id, counts in counts_by_id.items():
        cols_with_issues = [
            f"{name} ({counts[i]})" for i, name in enumerate(headers)
            if i != id_index and counts[i] > threshold
        ]
        if cols_with_issues:
            alerts.append(
                f"⚠️ {id_col} {unit_id} apresentou registros vazios em: "
                f"{', '.join(cols_with_issues)}. Possível erro no recebimento dos dados."
            )
    return alerts


def process_csv(
    csv_path: str,
    output_dir: str = OUTPUT_DIR,
    protect: bool = False,
    theme_color: str = DEFAULT_THEME,
    null_treatment: str = "na",
) -> str:
    headers, rows = _read_csv(csv_path)
    headers, rows = _reorder_image_cols(headers, rows)
    theme = build_theme(theme_color)

    wb = Workbook()
    ws = wb.active
    ws.title = os.path.splitext(os.path.basename(csv_path))[0][:31]

    total_cols = len(headers)

    # --- Header row ---
    header_font = Font(bold=True, color=theme["header_text"], size=12)
    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = theme["header_fill"]
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    ws.row_dimensions[1].height = 30

    # --- Data rows ---
    null_font = Font(color="999999", italic=True)
    null_count = 0
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if _is_null(value):
                null_count += 1
                if null_treatment == "na":
                    cell.value = "N/A"
                    cell.fill = NULL_FILL
                    cell.font = null_font
                elif null_treatment == "zero":
                    cell.value = 0
                else:  # "vazio"
                    cell.value = ""
            else:
                cell.value = value
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = CELL_BORDER

    # --- Footer row ---
    footer_row = len(rows) + 2

    def _write_footer_row(row_idx: int, text: str, fill: PatternFill, font: Font) -> None:
        ws.merge_cells(
            start_row=row_idx, start_column=1,
            end_row=row_idx, end_column=total_cols,
        )
        cell = ws.cell(row=row_idx, column=1)
        cell.value = text
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = CELL_BORDER
        ws.row_dimensions[row_idx].height = 25
        for col_idx in range(2, total_cols + 1):
            ws.cell(row=row_idx, column=col_idx).border = CELL_BORDER

    _write_footer_row(
        footer_row,
        (
            f"Total rows: {len(rows)}  |  "
            f"Columns: {total_cols}  |  "
            f"Empty cells: {null_count}  |  "
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ),
        theme["footer_fill"],
        Font(bold=True, color=theme["footer_text"], size=11),
    )

    alerts = _find_null_alerts(headers, rows)
    for i, alert_msg in enumerate(alerts, start=1):
        _write_footer_row(
            footer_row + i,
            alert_msg,
            ALERT_FILL,
            Font(bold=True, color="7F6000", size=11),
        )

    # --- Column widths ---
    for col_idx, col_name in enumerate(headers, start=1):
        longest_value = max((len(row[col_idx - 1]) for row in rows), default=0)
        max_len = max(len(col_name), longest_value) + 4
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len, 50)

    ws.freeze_panes = "A2"

    if protect:
        ws.protection.sheet = True
        ws.protection.selectLockedCells = False

    os.makedirs(output_dir, exist_ok=True)
    filename = os.path.splitext(os.path.basename(csv_path))[0]
    out_path = os.path.join(output_dir, f"{filename}.xlsx")
    wb.save(out_path)
    return out_path


def main():
    if not os.path.isdir(DATA_DIR):
        print(f"[error] Data folder not found: {DATA_DIR}")
        return

    csv_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".csv")]
    if not csv_files:
        print("[info] No CSV files found in data/")
        return

    print(f"[info] Found {len(csv_files)} file(s)\n")
    for name in sorted(csv_files):
        path = os.path.join(DATA_DIR, name)
        print(f"  Processing: {name}")
        out = process_csv(path)
        print(f"  Saved:      {out}\n")

    print("[done]")


if __name__ == "__main__":
    main()
