import os
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

HEADER_FILL = PatternFill(start_color="00AD32", end_color="00AD32", fill_type="solid")
NULL_FILL = PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid")
FOOTER_FILL = PatternFill(start_color="C9EAB8", end_color="C9EAB8", fill_type="solid")
THIN = Side(style="thin")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


ALERT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def _is_null(value: str) -> bool:
    return str(value).strip().lower() in ("", "nan", "none", "null", "$null$", "n/a", "na")


def _find_id_col(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        if "unidade" in col.lower() and "produ" in col.lower():
            return col
    return None


def _find_null_alerts(df: pd.DataFrame, threshold: int = 10) -> list[str]:
    id_col = _find_id_col(df)
    if id_col is None:
        return []

    alerts = []

    for unit_id, group in df.groupby(id_col, sort=False):
        cols_with_issues = []
        for col in df.columns:
            if col == id_col:
                continue
            null_count = group[col].apply(_is_null).sum()
            if null_count > threshold:
                cols_with_issues.append(f"{col} ({null_count})")
        if cols_with_issues:
            alerts.append(
                f"⚠️ Unidade de Producao {unit_id} apresentou registros vazios em: "
                f"{', '.join(cols_with_issues)}. Possível erro no recebimento dos dados."
            )
    return alerts


def process_csv(csv_path: str, output_dir: str = OUTPUT_DIR, protect: bool = False) -> str:
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)

    wb = Workbook()
    ws = wb.active
    ws.title = os.path.splitext(os.path.basename(csv_path))[0][:31]

    total_cols = len(df.columns)

    # --- Header row ---
    header_font = Font(bold=True, color="FFFFFF", size=12)
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    ws.row_dimensions[1].height = 30

    # --- Data rows ---
    null_font = Font(color="999999", italic=True)
    null_count = 0
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if _is_null(value):
                cell.value = "N/A"
                cell.fill = NULL_FILL
                cell.font = null_font
                null_count += 1
            else:
                cell.value = value
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = CELL_BORDER

    # --- Footer row ---
    footer_row = len(df) + 2

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
            f"Total rows: {len(df)}  |  "
            f"Columns: {total_cols}  |  "
            f"Empty cells: {null_count}  |  "
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ),
        FOOTER_FILL,
        Font(bold=True, color="005C1A", size=11),
    )

    alerts = _find_null_alerts(df)
    for i, alert_msg in enumerate(alerts, start=1):
        _write_footer_row(
            footer_row + i,
            alert_msg,
            ALERT_FILL,
            Font(bold=True, color="7F6000", size=11),
        )

    # --- Column widths ---
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_values = df.iloc[:, col_idx - 1].astype(str)
        max_len = max(len(str(col_name)), col_values.str.len().max()) + 4
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len, 50)

    # Freeze header row
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
