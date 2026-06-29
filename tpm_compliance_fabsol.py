"""
TPM × FW Rate Card Compliance Validator (FabSol) — v2.1
Date: 2026-06-29

v2.1 changes:
✓ NEW: Filter CD_Category == "FAB SOL" (replaces Sales_Org=='7001' filter)
✓ NEW: Filter Quarter_Year == <sheet's Q_Y> (hard filter, like H&H)
✓ Justification "Quarter mismatch" no longer needed (rows pre-filtered)

v2.0 features preserved:
✓ TPM sheet picker
✓ Rich progress bars + emoji status lines (H&H-style UX)
✓ Rate Card SET indexing from {AJ,L,V,X}/{J,N}/{H,AA}
✓ WPRM → RSP Promo auto-rename
✓ Single-source version constant in banner / title / Metadata sheet
"""
from __future__ import annotations

import os
import re
import sys
import warnings
import logging
import contextlib
import io as _io
import traceback
from datetime import date
from pathlib import Path

import polars as pl
import pandas as pd
import questionary
from questionary import Style as QStyle
from openpyxl import load_workbook
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import (
    Progress, SpinnerColumn, TextColumn, BarColumn,
    TimeElapsedColumn, MofNCompleteColumn,
)
from rich import box

# ============================================================
# VERSION
# ============================================================
__version__       = "2.1"
__version_date__  = "2026-06-29"
__version_notes__ = "Filter CD_Category=='FAB SOL' + Quarter_Year==<sheet Q_Y>"

console = Console()

if os.name == "nt":
    try:
        os.system(f"title TPM FabSol Compliance v{__version__}")
    except Exception:
        pass

warnings.filterwarnings("ignore")
logging.getLogger("polars").setLevel(logging.ERROR)

# ============================================================
# PATH
# ============================================================
if getattr(sys, "frozen", False):
    WORK_DIR = Path(sys.executable).parent
else:
    WORK_DIR = Path(__file__).parent

# ============================================================
# CONFIG
# ============================================================
CD_CATEGORY_KEEP = "FAB SOL"
TOLERANCE = 0.5

DROP_COLUMNS = ["RSP CCBT", "GAP", "promo mechanic"]

COLUMN_RENAME_MAP = {"WPRM": "RSP Promo"}

REQUIRED_COLUMNS = [
    "PromoGroupProductDesc",
    "Instore_Start",
    "Instore_End",
    "TPM_InvestmentDescription",
    "RSP Promo",
]

# Excel column SET mapping for Rate_Number (v1.6 spec)
def col_letter_to_idx(letter: str) -> int:
    letter = letter.upper().strip()
    n = 0
    for c in letter:
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n - 1

PERIOD_COL_MAP = {
    "Weekly":     [col_letter_to_idx(c) for c in ("AJ", "L", "V", "X")],
    "Bi-weekly":  [col_letter_to_idx(c) for c in ("J", "N")],
    "Tri-weekly": [col_letter_to_idx(c) for c in ("J", "N")],
    "Monthly":    [col_letter_to_idx(c) for c in ("H", "AA")],
}

PERIOD_BUCKETS = {"Weekly": 7, "Bi-weekly": 14, "Tri-weekly": 21, "Monthly": 30}

COLOR_MAP_HEX = {
    "Comply":                 "#C6EFCE",
    "Not Comply":             "#FFC7CE",
    "NO rate card available": "#FFEB9C",
    "Missing Final_Rsp":      "#D9D9D9",
}

QSTYLE = QStyle([
    ("qmark",       "fg:#00aaff bold"),
    ("question",    "bold"),
    ("answer",      "fg:#00aaff bold"),
    ("pointer",     "fg:#ff8800 bold"),
    ("highlighted", "fg:#00aaff bold noreverse"),
    ("selected",    "noinherit"),
    ("instruction", "fg:#888888 italic"),
])

# ============================================================
# UI HELPERS
# ============================================================
def banner():
    console.print()
    console.print(Panel.fit(
        f"[bold cyan]TPM × FW Rate Card Compliance Validator (FabSol)[/]\n"
        f"[bold]Version {__version__}[/]  [dim]({__version_date__})[/]\n"
        f"[dim italic]{__version_notes__}[/]\n"
        f"[dim]Working dir: {WORK_DIR}[/]",
        border_style="cyan",
        box=box.DOUBLE,
        title=f"[bold magenta]v{__version__}[/]",
        subtitle=f"[dim]Python {sys.version_info.major}.{sys.version_info.minor}[/]",
    ))
    console.print()


def list_xlsx_files(folder: Path) -> list[Path]:
    return sorted([
        p for p in folder.iterdir()
        if p.suffix.lower() in (".xlsx", ".xlsb", ".xls")
        and not p.name.startswith("~$")
        and not p.name.startswith("FabSol_Compliance_")
        and not p.name.startswith("TPM_FabSol_Compliance_")
        and not p.name.startswith("TPM_Rate_Card_Compliance_Output")
        and p.is_file()
    ], key=lambda p: p.stat().st_mtime, reverse=True)


def pick_file(files: list[Path], prompt: str) -> Path:
    choices = [questionary.Choice(title=f.name, value=f) for f in files]
    answer = questionary.select(
        prompt, choices=choices, style=QSTYLE,
        instruction="(↑/↓ to move, Enter to select, Esc to cancel)",
    ).ask()
    if answer is None:
        console.print("[yellow]Cancelled by user[/]")
        sys.exit(0)
    return answer


def pick_sheet(prompt: str, sheets: list[str]) -> str:
    answer = questionary.select(
        prompt, choices=sheets, style=QSTYLE,
        instruction="(↑/↓ to move, Enter to select, Esc to cancel)",
    ).ask()
    if answer is None:
        console.print("[yellow]Cancelled by user[/]")
        sys.exit(0)
    return answer


def get_sheet_names(path: Path) -> list[str]:
    if path.suffix.lower() == ".xlsb":
        from pyxlsb import open_workbook
        with open_workbook(path) as wb:
            return wb.sheets
    wb = load_workbook(path, read_only=True, data_only=False)
    sheets = wb.sheetnames
    wb.close()
    return sheets


def derive_target_quarter(sheet_name: str) -> str | None:
    """Extract 'Q# YYYY' from a sheet name like 'Q3 2026' or 'Rate Card Q3-2026'."""
    m = re.search(r"Q\s*(\d).*?(\d{4})", sheet_name)
    return f"Q{m.group(1)} {m.group(2)}" if m else None


def make_progress() -> Progress:
    return Progress(
        SpinnerColumn(style="cyan"),
        TextColumn("[bold blue]{task.description}"),
        BarColumn(bar_width=40),
        MofNCompleteColumn(),
        TextColumn("[dim]•[/]"),
        TimeElapsedColumn(),
        console=console,
        transient=False,
    )

# ============================================================
# POLARS PROCESSING
# ============================================================
def _columns_to_drop(df_columns: list) -> list:
    drop_lower = {c.lower() for c in DROP_COLUMNS}
    result = []
    for col in df_columns:
        c = str(col).strip()
        if c.lower() in drop_lower:
            result.append(col)
        elif c.lower().startswith("unnamed") or c.startswith("__UNNAMED__"):
            result.append(col)
    return result


def _apply_column_renames(df: pl.DataFrame) -> pl.DataFrame:
    cols = set(df.columns)
    for original, standard in COLUMN_RENAME_MAP.items():
        if original in cols and standard not in cols:
            df = df.rename({original: standard})
            console.print(f"   🔁 Renamed [yellow]{original}[/] → [cyan]{standard}[/]")
    return df


def load_tpm(tpm_path: Path, sheet_name: str, progress: Progress) -> pl.DataFrame:
    task = progress.add_task("Loading TPM file (calamine)", total=1)
    if tpm_path.suffix.lower() == ".xlsb":
        df_pd = pd.read_excel(tpm_path, sheet_name=sheet_name, engine="pyxlsb")
        df = pl.from_pandas(df_pd)
    else:
        with contextlib.redirect_stderr(_io.StringIO()):
            df = pl.read_excel(tpm_path, sheet_name=sheet_name, engine="calamine")
    progress.update(task, advance=1)
    return df


def validate_and_filter(df: pl.DataFrame, progress: Progress) -> pl.DataFrame:
    """
    Step 1: rename WPRM → RSP Promo + check required cols.
    Step 2: filter CD_Category == 'FAB SOL'.
    (Quarter_Year filter is applied AFTER helper columns are built.)
    """
    task = progress.add_task("Validating & filtering", total=2)

    # 1) Rename + required columns check
    df = _apply_column_renames(df)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        hint = ""
        if "RSP Promo" in missing and "WPRM" in df.columns:
            hint = "\n[Hint] 'WPRM' present but not renamed."
        console.print(f"[red]❌ Missing required columns:[/] {missing}{hint}")
        sys.exit(1)
    progress.update(task, advance=1)

    # 2) CD_Category filter
    if "CD_Category" not in df.columns:
        console.print("[red]❌ Column 'CD_Category' not found in TPM[/]")
        sys.exit(1)
    before = df.height
    df = df.filter(
        pl.col("CD_Category").cast(pl.Utf8).str.strip_chars() == CD_CATEGORY_KEEP
    )
    console.print(f"   🎯 CD_Category=='{CD_CATEGORY_KEEP}': kept {df.height:,} / {before:,}")
    progress.update(task, advance=1)
    return df


def _extract_promo_type(text):
    if not isinstance(text, str): return None
    t = text.upper()
    if "BOGO"   in t: return "BOGO"
    if "2F1"    in t: return "2F1"
    if "LAKSUE" in t: return "LAKSUE"
    if "_PO"    in t: return "PO"
    return None


def _quarter_year(dt):
    if dt is None: return None
    try:
        return f"Q{((dt.month - 1) // 3) + 1} {dt.year}"
    except Exception:
        return None


def _closest_period(days):
    if days is None: return None
    try: d = int(days)
    except (TypeError, ValueError): return None
    return min(PERIOD_BUCKETS, key=lambda k: abs(PERIOD_BUCKETS[k] - d))


def add_helper_columns(df: pl.DataFrame, target_quarter: str | None,
                       progress: Progress) -> pl.DataFrame:
    """
    Adds Promo_Type, Final_Rsp, Duration_Days, Period, Quarter_Year,
    then HARD-FILTERS to Quarter_Year == target_quarter, then prints counts.
    """
    task = progress.add_task("Building helper columns", total=8)

    # 1. Promo_Type
    df = df.with_columns(
        pl.col("TPM_InvestmentDescription")
          .map_elements(_extract_promo_type, return_dtype=pl.Utf8)
          .alias("Promo_Type")
    ); progress.update(task, advance=1)

    # 2. Final_Rsp
    df = df.with_columns(
        pl.col("RSP Promo").cast(pl.Float64, strict=False).alias("Final_Rsp")
    ); progress.update(task, advance=1)

    # 3. Duration_Days
    df = df.with_columns(
        (pl.col("Instore_End").cast(pl.Date, strict=False)
         - pl.col("Instore_Start").cast(pl.Date, strict=False))
        .dt.total_days().alias("Duration_Days")
    ); progress.update(task, advance=1)

    # 4. Period bucket
    df = df.with_columns(
        pl.col("Duration_Days")
          .map_elements(_closest_period, return_dtype=pl.Utf8)
          .alias("Period")
    ); progress.update(task, advance=1)

    # 5. Quarter_Year
    df = df.with_columns(
        pl.col("Instore_Start").cast(pl.Date, strict=False)
          .map_elements(_quarter_year, return_dtype=pl.Utf8)
          .alias("Quarter_Year")
    ); progress.update(task, advance=1)

    # 6. HARD FILTER on Quarter_Year
    if target_quarter:
        before = df.height
        df = df.filter(pl.col("Quarter_Year") == target_quarter)
        console.print(f"   📅 Quarter_Year=='{target_quarter}': kept {df.height:,} / {before:,}")
    else:
        console.print("[yellow]⚠ Could not derive Q_Y from sheet name — skipping Quarter filter[/]")
    progress.update(task, advance=1)

    # 7. Print Promo_Type counts
    pt_counts = {row["Promo_Type"]: row["len"]
                 for row in df.group_by("Promo_Type").len().iter_rows(named=True)}
    console.print(f"   Promo_Type counts: {pt_counts}")
    progress.update(task, advance=1)

    # 8. Print Period counts
    pd_counts = {row["Period"]: row["len"]
                 for row in df.group_by("Period").len().iter_rows(named=True)}
    console.print(f"   Period counts:     {pd_counts}")
    progress.update(task, advance=1)
    return df

# ============================================================
# RATE CARD INDEXING (SET parsing)
# ============================================================
def parse_set_cell(v) -> list:
    if v is None or (isinstance(v, float) and pd.isna(v)): return []
    if isinstance(v, (int, float)):
        return [round(float(v), 1)]
    s = str(v).strip()
    if not s or s.lower() == "nan": return []
    parts = [p.strip() for p in re.split(r"[/,]", s) if p.strip()]
    out = []
    for p in parts:
        try:
            out.append(round(float(p), 1))
        except ValueError:
            try:
                out.append(round(float(p.rstrip("%").strip()), 1))
            except ValueError:
                out.append(p.upper())
    return out


def format_set(values: list) -> str | None:
    if not values: return None
    parts = []
    for v in values:
        parts.append(v if isinstance(v, str) else f"{v:g}")
    return "/".join(parts)


def find_header_row(ws, keyword="Promo Group"):
    for r in range(1, 30):
        for c in range(1, 30):
            v = ws.cell(r, c).value
            if isinstance(v, str) and keyword.lower() in v.lower():
                return r
    raise RuntimeError(f"Header row containing '{keyword}' not found")


def build_rate_card_index(path: Path, sheet: str, progress: Progress):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    header_row = find_header_row(ws, "Promo Group")
    console.print(f"   Header row at: {header_row}")

    ppg_col_idx = None
    period_col_idx = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str):
            vl = v.strip().lower()
            if "promo group" in vl or vl == "ppg":
                ppg_col_idx = c
            elif vl == "period":
                period_col_idx = c
    if ppg_col_idx is None:
        raise RuntimeError("PPG column not found in Rate Card header")

    index: dict[str, dict[str, list]] = {}
    first_data_row = header_row + 1
    task = progress.add_task("Indexing Rate Card", total=ws.max_row - header_row)
    row_count = 0
    for r in range(first_data_row, ws.max_row + 1):
        ppg = ws.cell(r, ppg_col_idx).value
        if ppg is None or str(ppg).strip() == "":
            progress.update(task, advance=1); continue
        key = str(ppg).strip()
        period_val = ws.cell(r, period_col_idx).value if period_col_idx else None
        period_key = str(period_val).strip() if period_val else None

        per_period: dict[str, list] = index.setdefault(key, {})
        for period_name, col_idx_list in PERIOD_COL_MAP.items():
            if period_key and period_key != period_name:
                continue
            collected: list = []
            for ci in col_idx_list:
                cell_val = ws.cell(r, ci + 1).value  # openpyxl is 1-indexed
                collected.extend(parse_set_cell(cell_val))
            if collected:
                per_period.setdefault(period_name, []).extend(collected)
        row_count += 1
        progress.update(task, advance=1)

    console.print(f"   Indexed [cyan]{row_count}[/] PPG rows")
    wb.close()
    return index


def lookup_rate_number(index, ppg, period):
    if ppg is None or period is None: return None, []
    key = str(ppg).strip()
    bucket = index.get(key)
    if not bucket: return None, []
    vals = bucket.get(period, [])
    return (format_set(vals) if vals else None), vals

# ============================================================
# COMPLIANCE EVALUATION
# (Quarter mismatch impossible — already filtered out)
# ============================================================
def justify(rate_str, final_rsp, vals):
    if not rate_str:
        return "NO rate card available"
    if final_rsp is None:
        return "Missing Final_Rsp"
    for v in vals:
        try:
            if abs(float(v) - float(final_rsp)) <= TOLERANCE:
                return "Comply"
        except (TypeError, ValueError):
            continue
    return "Not Comply"


def evaluate(df: pl.DataFrame, rc_index, progress: Progress) -> pl.DataFrame:
    task = progress.add_task("Evaluating compliance", total=df.height)
    rate_nums: list[str | None] = []
    rate_lists: list[list] = []
    justifications: list[str] = []

    for row in df.iter_rows(named=True):
        rs, vals = lookup_rate_number(
            rc_index, row.get("PromoGroupProductDesc"), row.get("Period")
        )
        j = justify(rs, row.get("Final_Rsp"), vals)
        rate_nums.append(rs)
        rate_lists.append(vals)
        justifications.append(j)
        progress.update(task, advance=1)

    df = df.with_columns([
        pl.Series("Rate_Number", rate_nums, dtype=pl.Utf8),
        pl.Series("Justification", justifications, dtype=pl.Utf8),
    ])
    return df

# ============================================================
# OUTPUT — xlsxwriter; colors ONLY Justification cell
# ============================================================
def write_output(df: pl.DataFrame, out_path: Path, progress: Progress):
    import xlsxwriter
    final_drop = _columns_to_drop(df.columns)
    if final_drop:
        df = df.drop(final_drop)

    task = progress.add_task("Writing Excel (xlsxwriter)", total=df.height)
    wb = xlsxwriter.Workbook(out_path, {"constant_memory": True})
    ws = wb.add_worksheet("Compliance")

    header_fmt = wb.add_format({"bold": True, "bg_color": "#305496",
                                "font_color": "white", "border": 1})
    fmt_map = {k: wb.add_format({"bg_color": v, "border": 1})
               for k, v in COLOR_MAP_HEX.items()}

    cols = df.columns
    for ci, col in enumerate(cols):
        ws.write(0, ci, col, header_fmt)
    try:
        just_idx = cols.index("Justification")
    except ValueError:
        just_idx = -1

    for ri, row in enumerate(df.iter_rows(), start=1):
        for ci, val in enumerate(row):
            if ci == just_idx and val in fmt_map:
                ws.write(ri, ci, val, fmt_map[val])
            else:
                if val is None:
                    ws.write_blank(ri, ci, None)
                else:
                    ws.write(ri, ci, val)
        progress.update(task, advance=1)

    # Metadata sheet
    meta = wb.add_worksheet("Metadata")
    meta.write(0, 0, "Key", header_fmt); meta.write(0, 1, "Value", header_fmt)
    meta_rows = [
        ("Version", __version__),
        ("Version Date", __version_date__),
        ("Notes", __version_notes__),
        ("Generated", date.today().isoformat()),
        ("Working Dir", str(WORK_DIR)),
        ("Rows", df.height),
        ("CD_Category Filter", CD_CATEGORY_KEEP),
    ]
    for i, (k, v) in enumerate(meta_rows, start=1):
        meta.write(i, 0, k); meta.write(i, 1, str(v))

    wb.close()

# ============================================================
# RICH SUMMARY DISPLAY
# ============================================================
def display_summary(df: pl.DataFrame, out_path: Path):
    counts_df = df.group_by("Justification").len().sort("len", descending=True)
    counts = {row["Justification"]: row["len"] for row in counts_df.iter_rows(named=True)}

    table = Table(title=f"📊 FabSol Compliance Summary  (v{__version__})",
                  box=box.ROUNDED, show_header=True, header_style="bold cyan")
    table.add_column("Justification", style="bold")
    table.add_column("Count", justify="right", style="cyan")
    total = df.height or 1
    for k, v in counts.items():
        table.add_row(k, f"{v:,}  ({v/total:.1%})")
    table.add_row("[bold]TOTAL[/]", f"[bold]{df.height:,}[/]")
    console.print(table)
    console.print(f"\n[green bold]✅ Saved:[/] {out_path}")

# ============================================================
# MAIN
# ============================================================
def main():
    banner()
    console.print(f"[dim]Startup — FabSol Validator v{__version__}[/]")

    files = list_xlsx_files(WORK_DIR)
    if not files:
        console.print(f"[red]❌ No Excel files in {WORK_DIR}[/]")
        sys.exit(1)

    # --- Pick TPM file + sheet
    tpm_path = pick_file(files, "Select the TPM file:")
    tpm_sheets = get_sheet_names(tpm_path)
    tpm_sheet = pick_sheet(f"Select sheet in TPM ({tpm_path.name}):", tpm_sheets)

    # --- Pick Rate Card file + sheet
    rate_path = pick_file(files, "Select the FW Rate Card file:")
    rate_sheets = get_sheet_names(rate_path)
    rate_sheet = pick_sheet(f"Select sheet in Rate Card ({rate_path.name}):", rate_sheets)

    # Quarter derived from RATE CARD sheet name (this drives the Quarter_Year filter)
    target_quarter = derive_target_quarter(rate_sheet)
    if target_quarter:
        console.print(f"[dim]Rate Card target quarter: [cyan]{target_quarter}[/][/]")
    else:
        console.print("[yellow]⚠ Could not parse quarter from sheet name "
                      f"'{rate_sheet}' — Quarter_Year filter will be skipped[/]")

    # Output filename includes Q_Y + date (mirrors H&H naming scheme)
    q_for_name = (target_quarter or "AllQ").replace(" ", "_")
    out_name = f"FabSol_Compliance_{q_for_name}_{date.today():%Y-%m-%d}.xlsx"
    out_path = WORK_DIR / out_name

    with make_progress() as progress:
        df = load_tpm(tpm_path, tpm_sheet, progress)
        df = validate_and_filter(df, progress)
        df = add_helper_columns(df, target_quarter, progress)
        rc_index = build_rate_card_index(rate_path, rate_sheet, progress)
        df = evaluate(df, rc_index, progress)
        write_output(df, out_path, progress)

    display_summary(df, out_path)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print(f"\n[yellow]Cancelled by user (v{__version__})[/]")
        sys.exit(0)
    except Exception as e:
        console.print(f"\n[red bold]❌ Error (v{__version__}): {e}[/]")
        console.print(Panel(traceback.format_exc(), title=f"Traceback — v{__version__}",
                            border_style="red", box=box.ROUNDED))
        sys.exit(1)