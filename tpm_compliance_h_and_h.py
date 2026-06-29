"""
Home & Hygiene (H&H) × FW Rate Card Compliance Validator — v2.3
Date: 2026-06-23

v2.3 changes:
  ✓ CD_Category filter = "H&H" (was "FAB SOL")
  ✓ Color ONLY the Justification cell (not the entire row)
  ✓ Output filename includes category + quarter

v2.2 — Calamine reader + xlsxwriter (fast)
v2.1 — Banner + styled picker + progress bar
v2.0 — Migrated pandas → polars
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
from datetime import date, datetime
from pathlib import Path

import polars as pl
import questionary
from questionary import Style as QStyle
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import (
    Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn, MofNCompleteColumn
)
from rich import box

# ============================================================
# VERSION
# ============================================================
__version__       = "2.3"
__version_date__  = "2026-06-23"
__version_notes__ = "Filter H&H category + Justification-cell-only coloring"

console = Console()

if os.name == "nt":
    try:
        os.system(f"title H&H Compliance v{__version__}")
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
CD_CATEGORY_KEEP = "H&H"           # ⬅ FIXED: was "FAB SOL"
TOLERANCE = 0.5

REQUIRED_COLUMNS = [
    "PromoGroupProductDesc",
    "Instore_Start",
    "Instore_End",
    "TPM_InvestmentDescription",
    "RSP Promo",
]

CUSTOMER_MAP = {
    "Cpaxtra": ["Modern Trade-Tesco", "Makro", "Modern Trade-7-ELEVEN"],
    "SYL":     ["Modern Trade-Casino"],
    "CJ":      ["Modern Trade-CJ EXPRESS"],
    "Tops":    ["Modern Trade-TOPS"],
    "Makro":   ["MAKRO-SHV Makro"],
}
TPM_CUSTOMER_TO_RC = {tpm_cust: rc_tag
                      for rc_tag, lst in CUSTOMER_MAP.items()
                      for tpm_cust in lst}

RC_COLUMNS = {
    "PO": {
        "Weekly":     ["K", "Y", "S", "AA"],
        "Bi-weekly":  ["I", "S"],
        "Tri-weekly": ["I", "S"],
        "Monthly":    ["G", "S"],
    },
    "LAKSUE": {
        "Weekly":     ["Q"], "Bi-weekly": ["Q"],
        "Tri-weekly": ["Q"], "Monthly":   ["Q"],
    },
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
    ("answer",      "fg:#00ff88 bold"),
    ("pointer",     "fg:#ff8800 bold"),
    ("highlighted", "fg:#00aaff bold"),
    ("selected",    "fg:#00ff88"),
    ("instruction", "fg:#888888 italic"),
])

# ============================================================
# UI HELPERS
# ============================================================
def banner():
    console.print()
    console.print(Panel.fit(
        f"[bold cyan]Home & Hygiene × Rate Card Compliance Validator[/]\n"
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
        and not p.name.startswith("HandH_Compliance_")
        and not p.name.startswith("HH_Compliance_")
        and p.is_file()
    ], key=lambda p: p.stat().st_mtime, reverse=True)

def pick_file(files: list[Path], prompt: str, suggest_keyword: str = "") -> Path:
    if suggest_keyword:
        kw = suggest_keyword.lower()
        files = [f for f in files if kw in f.name.lower()] + \
                [f for f in files if kw not in f.name.lower()]
    choices = [questionary.Choice(title=f.name, value=f) for f in files]
    answer = questionary.select(prompt, choices=choices, style=QSTYLE,
                                instruction="(↑/↓ to move, Enter to select)").ask()
    if answer is None:
        console.print("[yellow]Cancelled by user[/]"); sys.exit(0)
    return answer

def pick_sheet(prompt: str, sheets: list[str], *default_keywords) -> str:
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    default = next((s for s in sheets if all(norm(k) in norm(s) for k in default_keywords)), None)
    answer = questionary.select(prompt, choices=sheets, default=default, style=QSTYLE,
                                instruction="(↑/↓ to move, Enter to select)").ask()
    if answer is None:
        console.print("[yellow]Cancelled by user[/]"); sys.exit(0)
    return answer

def derive_target_quarter(sheet_name: str) -> str | None:
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
# STEP 1 — VALIDATION & FEATURE ENGINEERING
# ============================================================
def validate_columns(df: pl.DataFrame) -> pl.DataFrame:
    if "RSP Promo" not in df.columns and "WPRM" in df.columns:
        df = df.rename({"WPRM": "RSP Promo"})
        console.print("   🔁 Renamed [yellow]WPRM[/] → [green]RSP Promo[/]")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        console.print(f"[red]❌ Missing required columns:[/] {missing}")
        sys.exit(1)
    return df

def filter_cd_category(df: pl.DataFrame) -> pl.DataFrame:
    if "CD_Category" not in df.columns:
        console.print("[red]❌ Column 'CD_Category' not found in TPM[/]")
        sys.exit(1)
    before = df.height
    df = df.filter(
        pl.col("CD_Category").cast(pl.Utf8).str.strip_chars() == CD_CATEGORY_KEEP
    )
    console.print(f"   🎯 CD_Category=='{CD_CATEGORY_KEEP}': kept {df.height:,} / {before:,}")
    return df

def _extract_promo_type(text):
    if not isinstance(text, str): return None
    t = text.upper()
    if "BOGO"   in t: return "BOGO"
    if "2F1"    in t: return "2F1"
    if "LAKSUE" in t: return "LAKSUE"
    if "_PO"    in t: return "PO"
    return None

def _extract_promo_price(text, promo_type):
    if not isinstance(text, str) or promo_type is None: return None
    t = text.upper()
    if promo_type == "LAKSUE":
        m = re.search(r"LAKSUE\s*_?\s*(\d+(?:\.\d+)?)", t)
        return float(m.group(1)) if m else None
    if promo_type in ("PO", "2F1"):
        for pat in (r"2F1\s*_?\s*PO\s*(\d+(?:\.\d+)?)",
                    r"_PO\s*(\d+(?:\.\d+)?)",
                    r"\bPO\s*(\d+(?:\.\d+)?)"):
            m = re.search(pat, t)
            if m: return float(m.group(1))
    return None

def _closest_period(days):
    if days is None: return None
    try: d = int(days)
    except (TypeError, ValueError): return None
    return min(PERIOD_BUCKETS, key=lambda k: abs(PERIOD_BUCKETS[k] - d))

def _quarter_year(dt):
    if dt is None: return None
    try: return f"Q{((dt.month - 1) // 3) + 1} {dt.year}"
    except Exception: return None

def engineer_features(df: pl.DataFrame, target_quarter: str | None, progress: Progress) -> pl.DataFrame:
    task = progress.add_task("Building helper columns", total=8)

    df = df.with_columns(
        pl.col("TPM_InvestmentDescription")
          .map_elements(_extract_promo_type, return_dtype=pl.Utf8)
          .alias("Promo_Type")
    ); progress.update(task, advance=1)

    df = df.with_columns(
        pl.struct(["TPM_InvestmentDescription", "Promo_Type"])
          .map_elements(
              lambda s: _extract_promo_price(s["TPM_InvestmentDescription"], s["Promo_Type"]),
              return_dtype=pl.Float64)
          .alias("Promotion_Price")
    ); progress.update(task, advance=1)

    df = df.with_columns([
        pl.col("Instore_Start").cast(pl.Datetime, strict=False),
        pl.col("Instore_End").cast(pl.Datetime, strict=False),
    ]); progress.update(task, advance=1)

    df = df.with_columns(
        (pl.col("Instore_End") - pl.col("Instore_Start"))
          .dt.total_days().alias("DateRange_Days")
    ); progress.update(task, advance=1)

    df = df.with_columns(
        pl.col("DateRange_Days")
          .map_elements(_closest_period, return_dtype=pl.Utf8)
          .alias("Period")
    ); progress.update(task, advance=1)

    df = df.with_columns(
        pl.col("RSP Promo").cast(pl.Float64, strict=False)
    ).with_columns(
        pl.when(pl.col("RSP Promo").is_not_null() & (pl.col("RSP Promo") > 0))
          .then(pl.col("RSP Promo"))
          .otherwise(pl.col("Promotion_Price"))
          .alias("Final_Rsp")
    ); progress.update(task, advance=1)

    df = df.with_columns(
        pl.col("Instore_Start")
          .map_elements(_quarter_year, return_dtype=pl.Utf8)
          .alias("Quarter_Year")
    ); progress.update(task, advance=1)

    if target_quarter:
        before = df.height
        df = df.filter(pl.col("Quarter_Year") == target_quarter)
        console.print(f"   📅 Quarter_Year=='{target_quarter}': kept {df.height:,} / {before:,}")
    progress.update(task, advance=1)

    pt_counts = df.group_by("Promo_Type").len().to_dicts()
    pd_counts = df.group_by("Period").len().to_dicts()
    pt_str = ", ".join(f"{r['Promo_Type']}: {r['len']}" for r in pt_counts)
    pd_str = ", ".join(f"{r['Period']}: {r['len']}" for r in pd_counts)
    console.print(f"   [dim]Promo_Type counts:[/] {{{pt_str}}}")
    console.print(f"   [dim]Period counts:    [/] {{{pd_str}}}")
    return df

# ============================================================
# STEP 2 — RATE CARD LOOKUP
# ============================================================
def parse_cell_numbers(cell_value):
    if cell_value is None: return []
    nums = []
    for token in re.split(r"[\/,]", str(cell_value)):
        m = re.search(r"\d+(?:\.\d+)?", token.strip())
        if m: nums.append(float(m.group()))
    return nums

def parse_cell_raw(cell_value):
    if cell_value is None: return []
    s = str(cell_value).strip()
    if not s: return []
    return [seg.strip() for seg in re.split(r"[\/,]", s) if seg.strip()]

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

    aa_idx, ppg_idx = column_index_from_string("AA"), column_index_from_string("B")
    total_rows = ws.max_row - header_row
    task = progress.add_task("Indexing Rate Card", total=total_rows)

    index = {}
    for r in range(header_row + 1, ws.max_row + 1):
        ppg_val = ws.cell(r, ppg_idx).value
        if ppg_val and str(ppg_val).strip():
            key = str(ppg_val).strip()
            index.setdefault(key, {"row": r, "aa_value": ws.cell(r, aa_idx).value})
        progress.update(task, advance=1)
    console.print(f"   Indexed [bold]{len(index)}[/] PPG rows")
    return ws, index

def lookup_rate_number(ws, index, ppg, promo_type, period, tpm_customer):
    if ppg is None or promo_type is None or period is None:
        return "NOT FOUND", []
    if promo_type not in RC_COLUMNS:
        return "NOT FOUND", []
    key = str(ppg).strip()
    if key not in index:
        return "NOT FOUND", []

    row = index[key]["row"]
    cols = RC_COLUMNS[promo_type].get(period, [])
    pieces, all_nums = [], []
    for col_letter in cols:
        cell_val = ws.cell(row, column_index_from_string(col_letter)).value
        if col_letter == "AA":
            rc_tag = TPM_CUSTOMER_TO_RC.get(str(tpm_customer).strip() if tpm_customer else "")
            if rc_tag is None: continue
            aa_val = index[key]["aa_value"]
            if aa_val is None or str(aa_val).strip().lower() != rc_tag.lower():
                continue
        pieces.extend(parse_cell_raw(cell_val))
        all_nums.extend(parse_cell_numbers(cell_val))

    if not pieces:
        return "FOUND", []
    return " / ".join(pieces), all_nums

def justify(rate_number, final_rsp, rate_numbers_list):
    if rate_number == "NOT FOUND":
        return "NO rate card available"
    if final_rsp is None:
        return "Missing Final_Rsp"
    if rate_number == "FOUND" and not rate_numbers_list:
        return "Not Comply"
    if any(abs(final_rsp - r) <= TOLERANCE for r in rate_numbers_list):
        return "Comply"
    return "Not Comply"

def evaluate(df: pl.DataFrame, ws, rc_index, progress: Progress) -> pl.DataFrame:
    has_customer = "Customer" in df.columns
    task = progress.add_task("Evaluating compliance", total=df.height)

    rate_strs, justifications = [], []
    for row in df.iter_rows(named=True):
        rate_str, nums = lookup_rate_number(
            ws, rc_index,
            ppg=row.get("PromoGroupProductDesc"),
            promo_type=row.get("Promo_Type"),
            period=row.get("Period"),
            tpm_customer=row.get("Customer") if has_customer else None,
        )
        rate_strs.append(rate_str)
        justifications.append(justify(rate_str, row.get("Final_Rsp"), nums))
        progress.update(task, advance=1)

    return df.with_columns([
        pl.Series("Rate_Number", rate_strs, dtype=pl.Utf8),
        pl.Series("Justification", justifications, dtype=pl.Utf8),
    ])

# ============================================================
# OUTPUT — xlsxwriter (FAST) — colors ONLY Justification cell
# ============================================================
def write_output(df: pl.DataFrame, out_path: Path, progress: Progress):
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(out_path), {"constant_memory": True})
    ws = wb.add_worksheet("HH_Compliance")

    header_fmt = wb.add_format({
        "bold": True, "font_color": "white",
        "bg_color": "#305496", "align": "center", "valign": "vcenter",
    })

    # ⬇ Only the Justification cell gets a color fill
    just_fmts = {
        label: wb.add_format({"bg_color": hex_, "bold": True, "align": "center"})
        for label, hex_ in COLOR_MAP_HEX.items()
    }
    default_just_fmt = wb.add_format({"align": "center"})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})

    cols = df.columns
    jcol_idx = cols.index("Justification")
    date_col_idx = {i for i, name in enumerate(cols)
                    if df.schema[name] in (pl.Datetime, pl.Date)}

    # Header
    for c, col_name in enumerate(cols):
        ws.write(0, c, col_name, header_fmt)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, df.height, len(cols) - 1)

    # Body — only Justification cell gets a color; everything else default
    task = progress.add_task("Writing Excel (xlsxwriter)", total=df.height)
    for r, row in enumerate(df.iter_rows(), start=1):
        for c, v in enumerate(row):
            if c == jcol_idx:
                cell_fmt = just_fmts.get(v, default_just_fmt)
            elif c in date_col_idx:
                cell_fmt = date_fmt
            else:
                cell_fmt = None

            if v is None:
                if cell_fmt:
                    ws.write_blank(r, c, None, cell_fmt)
            elif c in date_col_idx and isinstance(v, datetime):
                ws.write_datetime(r, c, v.replace(tzinfo=None) if v.tzinfo else v, cell_fmt)
            elif isinstance(v, bool):
                if cell_fmt: ws.write_boolean(r, c, v, cell_fmt)
                else:        ws.write_boolean(r, c, v)
            elif isinstance(v, (int, float)):
                if cell_fmt: ws.write_number(r, c, v, cell_fmt)
                else:        ws.write_number(r, c, v)
            else:
                if cell_fmt: ws.write_string(r, c, str(v), cell_fmt)
                else:        ws.write_string(r, c, str(v))
        progress.update(task, advance=1)

    ws.set_column(jcol_idx, jcol_idx, 28)

    # Metadata sheet
    meta = wb.add_worksheet("Metadata")
    meta.set_column(0, 0, 20)
    meta.set_column(1, 1, 60)
    meta_data = [
        ("Version",   __version__),
        ("Date",      __version_date__),
        ("Notes",     __version_notes__),
        ("Category",  CD_CATEGORY_KEEP),
        ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Rows",      df.height),
        ("Columns",   len(cols)),
    ]
    for r, (k, v) in enumerate(meta_data):
        meta.write(r, 0, k, header_fmt)
        meta.write(r, 1, str(v))

    wb.close()

def display_summary(df: pl.DataFrame, out_path: Path):
    counts_df = df.group_by("Justification").len().sort("len", descending=True)
    counts = {row["Justification"]: row["len"] for row in counts_df.iter_rows(named=True)}

    table = Table(title=f"📊 H&H Compliance Summary  (v{__version__})", box=box.ROUNDED)
    table.add_column("Status", style="bold")
    table.add_column("Count", justify="right")
    table.add_column("%", justify="right")
    total = df.height
    for cat in ["Comply", "Not Comply", "NO rate card available", "Missing Final_Rsp"]:
        n = counts.get(cat, 0)
        pct = f"{(n/total*100):.1f}%" if total else "0%"
        style = {"Comply": "green", "Not Comply": "red",
                 "NO rate card available": "yellow",
                 "Missing Final_Rsp": "dim"}.get(cat, "")
        table.add_row(f"[{style}]{cat}[/]" if style else cat, str(n), pct)
    table.add_row("[bold]TOTAL[/]", str(total), "100%")
    console.print(table)
    console.print(Panel.fit(
        f"[green]✅ Output saved:[/] {out_path.name}\n[dim]{out_path}[/]",
        border_style="green", box=box.ROUNDED,
    ))

# ============================================================
# MAIN
# ============================================================
def main():
    banner()

    files = list_xlsx_files(WORK_DIR)
    if not files:
        console.print(f"[red]❌ No Excel files found in:[/] {WORK_DIR}")
        sys.exit(1)

    tpm_file = pick_file(files, "📄 Select TPM Data file:",     suggest_keyword="tpm")
    rc_file  = pick_file(files, "📄 Select FW Rate Card file:", suggest_keyword="rate")

    tpm_sheets = load_workbook(tpm_file, read_only=True).sheetnames
    rc_sheets  = load_workbook(rc_file,  read_only=True).sheetnames

    tpm_sheet = pick_sheet("📑 Select TPM data sheet:",            tpm_sheets, "raw")
    rc_sheet  = pick_sheet("📑 Select Rate Card sheet (Quarter):", rc_sheets,  "q3", "2026")

    target_quarter = derive_target_quarter(rc_sheet)
    console.print(f"\n   🎯 Target quarter: [bold]{target_quarter}[/]\n")

    with make_progress() as progress:
        load_task = progress.add_task("Loading TPM file (calamine)", total=1)
        with contextlib.redirect_stderr(_io.StringIO()):
            tpm = pl.read_excel(tpm_file, sheet_name=tpm_sheet, engine="calamine")
        progress.update(load_task, advance=1)
        console.print(f"   Loaded [bold]{tpm.height:,}[/] rows × {tpm.width} cols")

        v_task = progress.add_task("Validating & filtering", total=2)
        tpm = validate_columns(tpm);   progress.update(v_task, advance=1)
        tpm = filter_cd_category(tpm); progress.update(v_task, advance=1)

        tpm = engineer_features(tpm, target_quarter, progress)
        ws, rc_index = build_rate_card_index(rc_file, rc_sheet, progress)
        tpm = evaluate(tpm, ws, rc_index, progress)

        q_safe   = (target_quarter or "AllQuarters").replace(" ", "_")
        cat_safe = CD_CATEGORY_KEEP.replace("&", "_and_").replace(" ", "")
        out_name = f"{cat_safe}_Compliance_{q_safe}_{date.today():%Y-%m-%d}.xlsx"
        out_path = WORK_DIR / out_name

        write_output(tpm, out_path, progress)

    display_summary(tpm, out_path)

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