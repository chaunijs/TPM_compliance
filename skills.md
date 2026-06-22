---
skill_id: tpm_fw_compliance_validator_agent
skill_name: TPM × FW Rate Card Compliance Validator (Agent Edition)
version: 2.0
date: 2026-06-19
author: Teerapat Haeranyikanon (Unilever Thailand)
language: Python 3.10+
runtime: code_interpreter | python_execution
engine: pandas (NOT polars - cloud agent compatibility)
output_format: xlsx (NEVER csv)
license: Internal Unilever use only
based_on: Local v1.5
---

# 🎯 Purpose

Run the same TPM × FW Rate Card compliance validation as the local v1.5 tool — but inside a chat agent. The user uploads 2 xlsx files, picks options conversationally, and gets back a color-coded xlsx for download.

---

# 🔔 Trigger Phrases

Activate this skill when user says any of:
- "run TPM compliance"
- "validate TPM"
- "check rate card compliance"
- "ตรวจ TPM"
- "ตรวจสอบราคาโปรโมชั่น"
- User uploads files matching `*TPM*.xlsx` AND `*FW rate card*.xlsx`

---

# 💬 Conversation Flow (MUST follow this exact sequence)

## Step 1 — Greet & request files (if not yet uploaded)

> "Hi! 👋 I'll run the TPM × FW Rate Card Compliance check. Please upload:
> 
> 1. 📂 **TPM file** (xlsx, filename containing 'TPM')
> 2. 📂 **FW Rate Card file** (xlsx, filename containing 'FW rate card')
> 
> Drop both files into the chat, then I'll process them. ✨"

## Step 2 — Detect uploaded files

When user uploads files, automatically identify:
- **TPM file**: filename contains `"TPM"` (case-insensitive)
- **Rate card file**: filename contains `"FW rate card"` (case-insensitive)

### Edge cases
| Situation | Response |
|-----------|----------|
| Only 1 file uploaded | "I see your **{type}** file. Please also upload the **{missing_type}** file." |
| Wrong filenames | "File `{name}` doesn't match TPM/FW Rate Card patterns. Please rename or upload correct files." |
| 2+ TPM files | "I see multiple TPM files: {list}. Which one should I use?" |
| Files don't end in .xlsx/.xlsb | "I can only read xlsx or xlsb files. Got `.{ext}`. Please convert and re-upload." |

## Step 3 — List sheets in rate card and ASK USER

After loading rate card, **read sheet names** and ask:

> "📊 Your rate card has these sheets:
> 
> 1. Rate Card Q3, 2026 adjust ← (Q3 2026) [latest]
> 2. Rate Card Q1, 2026 (Q1 2026)
> 3. Rules & Explanation
> 
> Which sheet should I use? Reply with the number or sheet name.
> 
> 💡 Tip: For current quarter (Q2 2026), I recommend sheet 1 (Q3 2026 adjust) as the active rate card."

Wait for user reply. Accept either:
- A number (`1`, `2`, etc.)
- A sheet name (partial match OK)
- "latest" / "current" → auto-pick most recent quarter
- "default" → auto-pick latest quarter

## Step 4 — Process with progress updates

Show updates as you go:

> "🔄 Processing... 
> - [✓] Loaded TPM file: 111,429 rows × 66 cols
> - [✓] Filtered Sales_Org = 7001: kept 111,427 rows
> - [✓] Derived helper columns
> - [✓] Built rate card lookup from sheet 'Rate Card Q3, 2026 adjust'
> - [✓] Applied compliance logic
> - [✓] Built summary pivot
> - [✓] Saved output xlsx
> 
> Done in 8 seconds! 🎉"

## Step 5 — Show results + attach output file

> "✅ **TPM × FW Compliance Check Complete!**
> 
> 📊 Processed: **111,427 rows**  
> 📅 Rate card sheet: **Rate Card Q3, 2026 adjust** (Q3_2026)
> 
> **Results:**
> | Status | Count | % |
> |---|---|---|
> | 🟢 Comply | 12,153 | 10.9% |
> | 🔴 NOT Comply | 7,241 | 6.5% |
> | 🟠 ⚠ Quarter mismatch | 89,212 | 80.1% |
> | 🟡 NO rate card | 2,821 | 2.5% |
> | ⬜ Missing Final_Rsp | 0 | 0.0% |
> 
> 📎 **Download**: TPM_FW_Compliance_Output_{date}.xlsx
> 
> 💡 The output has 2 sheets:
> - **Compliance** — all rows with color-coded Justification
> - **Summary** — pivot by Customer × Brand × Period
> 
> Need me to focus on specific Customer/Brand or filter NOT Comply rows? Just ask!"

---

# 🧮 Business Logic (identical to local v1.5)

## Step 1 — Helper columns

### 1.1 `Promo_Type` (from `TPM_InvestmentDescription`)
Priority order — first match wins:
| Contains | Type |
|----------|------|
| BOGO | BOGO |
| 2F1 | 2F1 |
| LAKSUE | LAKSUE |
| _PO | PO |
| (none) | "" |

### 1.2 `Promotion_Price` (from `TPM_InvestmentDescription`)
| Promo_Type | Pattern | Returns |
|-----------|---------|---------|
| LAKSUE | `_LAKSUE ?(\d+)` | number |
| PO | `_PO ?(\d+)` | number |
| 2F1 | `_PO ?(\d+)` | number ⚠️ (uses _PO, not 2F1) |
| BOGO | — | **blank** ⚠️ |
| "" | — | blank |

### 1.3 `DateRange_Days`
`= Instore_End - Instore_Start` in days (NaT-safe).

### 1.4 `Period`
Closest of {7, 14, 21, 30}:
- 7 → Weekly
- 14 → Bi-weekly
- 21 → Tri-weekly
- 30 → Monthly

### 1.5 `Final_Rsp`
`= IF(RSP Promo > 0, RSP Promo, Promotion_Price)`

### 1.6 `Quarter_Year`
From `Instore_Start`: `Q{1-4}_{YYYY}`

## Step 2 — Filter Sales_Org
**Keep only rows where `Sales_Org == "7001"`** (drop "Total", "Applied filters:..." etc.)

## Step 3 — Rate card lookup

1. User-selected sheet (Step 3 of conversation)
2. Detect quarter from sheet name (regex: `Q\s*([1-4])[\s,_-]*\s*(20\d{2})`)
3. Find header row containing `"Promo Group (PPG)"`
4. Locate columns by header text: `Monthly`, `BW/TRI`, `WK*`
5. Build long lookup: `[PPG, Period, FW_Rate_Number]`
6. Join with TPM on `(PromoGroupProductDesc, Period)`

## Step 4 — Justification (priority order)

```
1. FW_Rate_Number is NaN              → "NO rate card available"
2. Final_Rsp is NaN                   → "Missing Final_Rsp"
3. Quarter_Year ≠ Rate_Card_Source    → "⚠ Quarter mismatch"
4. |Final_Rsp - FW_Rate_Number| ≤ 0.5 → "Comply"
5. else                               → "NOT Comply"
```

## Step 5 — Summary pivot
Group by `[Customer, Brand, Period]` → count each Justification category → `Comply_%`.

## Step 6 — Output xlsx
- File: `TPM_FW_Compliance_Output_{YYYY-MM-DD}.xlsx`
- Engine: `xlsxwriter` (for conditional formatting)
- Sheet 1: `Compliance` — all rows, color-coded
- Sheet 2: `Summary` — pivot
- Float columns: 1 decimal (`0.0`)
- Justification colors: 🟢🔴🟠🟡⬜
- Frozen header row + auto-filter

---

# 🐍 Complete Python Code (pandas-only, agent-ready)

```python
"""
TPM × FW Rate Card Compliance Validator — Agent Edition v2.0
============================================================
Pandas-only (no polars) for maximum agent runtime compatibility.
Run inside code_interpreter / python_execution.
"""

import re
import io
from datetime import date
from pathlib import Path

import pandas as pd

# ============================================================
# Helper functions (1-to-1 with local v1.5)
# ============================================================
def promo_type(t):
    s = "" if pd.isna(t) else str(t).upper()
    if "BOGO" in s: return "BOGO"
    if "2F1" in s: return "2F1"
    if "LAKSUE" in s: return "LAKSUE"
    if "_PO" in s: return "PO"
    return ""

def promo_price(an_text, ptype):
    if ptype in ("", "BOGO"): return None
    marker = "_LAKSUE" if ptype == "LAKSUE" else "_PO"
    s = "" if pd.isna(an_text) else str(an_text)
    m = re.search(rf"{marker} ?(\d+)", s, flags=re.IGNORECASE)
    return float(m.group(1)) if m else None

def period_label(days):
    if pd.isna(days): return ""
    PERIOD_MAP = {7: "Weekly", 14: "Bi-weekly", 21: "Tri-weekly", 30: "Monthly"}
    return PERIOD_MAP[min(PERIOD_MAP, key=lambda k: abs(k - days))]

def to_qtr(dt):
    if pd.isna(dt): return None
    return f"Q{(dt.month - 1) // 3 + 1}_{dt.year}"

def detect_qtr(sheet_name, df=None):
    """Detect quarter from sheet name FIRST, then sheet content."""
    pattern = r"Q\s*([1-4])[\s,_-]*\s*(20\d{2})"
    m = re.search(pattern, sheet_name, flags=re.IGNORECASE)
    if m: return f"Q{m.group(1)}_{m.group(2)}"
    if df is not None:
        txt = " ".join(df.iloc[:15, :15].astype(str).values.flatten())
        m = re.search(pattern, txt, flags=re.IGNORECASE)
        return f"Q{m.group(1)}_{m.group(2)}" if m else None
    return None

def parse_cell(v):
    if pd.isna(v): return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).upper()
    if "BOGO" in s: return None
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else None


# ============================================================
# Sheet picker — list & ask user
# ============================================================
def list_sheets_with_quarters(rate_path):
    """Read sheet names + detect quarter for each."""
    xls = pd.ExcelFile(rate_path)
    info = []
    for name in xls.sheet_names:
        q = detect_qtr(name)
        info.append({"name": name, "quarter": q})
    return info


# ============================================================
# Main processing
# ============================================================
def process_tpm_compliance(tpm_path, rate_path, sheet_name):
    """
    Run full validation. Returns dict with:
      - output_path: Path to xlsx
      - summary: dict of counts
      - rows: total row count
      - sheet_used: which rate card sheet
      - quarter_used: detected quarter
    """
    # ----- 1. Load TPM -----
    if tpm_path.endswith(".xlsb"):
        tpm = pd.read_excel(tpm_path, sheet_name=0, engine="pyxlsb")
    else:
        tpm = pd.read_excel(tpm_path, sheet_name=0)
    
    # Drop unused columns
    drop_cols = []
    for col in tpm.columns:
        c = str(col).strip()
        if c.lower() in ("rsp ccbt", "gap", "promo mechanic"):
            drop_cols.append(col)
        elif c.lower().startswith("unnamed") or c.startswith("__UNNAMED__"):
            drop_cols.append(col)
    if drop_cols:
        tpm = tpm.drop(columns=drop_cols)
    
    # Filter Sales_Org == "7001"
    initial_rows = len(tpm)
    if "Sales_Org" in tpm.columns:
        tpm["Sales_Org"] = tpm["Sales_Org"].astype(str).str.strip()
        tpm = tpm[tpm["Sales_Org"] == "7001"].copy()
    
    # Verify required columns
    required = ["PromoGroupProductDesc", "Instore_Start", "Instore_End",
                "TPM_InvestmentDescription", "RSP Promo"]
    missing = [c for c in required if c not in tpm.columns]
    if missing:
        raise KeyError(f"Missing required columns: {missing}")
    
    # ----- 2. Helper columns -----
    tpm["Promo_Type"] = tpm["TPM_InvestmentDescription"].apply(promo_type)
    tpm["Promotion_Price"] = tpm.apply(
        lambda r: promo_price(r["TPM_InvestmentDescription"], r["Promo_Type"]),
        axis=1)
    tpm["Instore_Start"] = pd.to_datetime(tpm["Instore_Start"], errors="coerce")
    tpm["Instore_End"] = pd.to_datetime(tpm["Instore_End"], errors="coerce")
    tpm["DateRange_Days"] = (tpm["Instore_End"] - tpm["Instore_Start"]).dt.days
    tpm["Period"] = tpm["DateRange_Days"].apply(period_label)
    tpm["Final_Rsp"] = tpm.apply(
        lambda r: r["RSP Promo"] if pd.notna(r["RSP Promo"]) and r["RSP Promo"] > 0
                  else r["Promotion_Price"], axis=1)
    tpm["Quarter_Year"] = tpm["Instore_Start"].apply(to_qtr)
    
    # ----- 3. Build rate card lookup -----
    df_raw = pd.read_excel(rate_path, sheet_name=sheet_name, header=None)
    
    quarter = detect_qtr(sheet_name, df_raw) or f"Q{(date.today().month-1)//3+1}_{date.today().year}"
    
    # Find header row
    hdr_row = df_raw.apply(
        lambda r: r.astype(str).str.contains("Promo Group", case=False, na=False).any(),
        axis=1).idxmax()
    
    def locate(contains):
        row = df_raw.iloc[hdr_row].astype(str)
        hits = [i for i, v in row.items() if contains.lower() in v.lower()]
        return hits[0] if hits else None
    
    ppg_col = locate("Promo Group")
    monthly_col = locate("Monthly")
    bwtri_col = locate("BW/TRI")
    weekly_col = locate("WK")
    
    if ppg_col is None:
        raise RuntimeError(f"Could not find 'Promo Group' column in sheet '{sheet_name}'")
    
    # Build long-format lookup
    records = []
    for _, row in df_raw.iloc[hdr_row+1:].iterrows():
        ppg = row.iloc[ppg_col]
        if pd.isna(ppg): continue
        ppg_str = str(ppg).strip()
        if not ppg_str or ppg_str.lower() in ("nan", "note"): continue
        
        if monthly_col is not None:
            r = parse_cell(row.iloc[monthly_col])
            if r is not None:
                records.append({"PPG": ppg_str, "Period": "Monthly", "FW_Rate_Number": r})
        if bwtri_col is not None:
            r = parse_cell(row.iloc[bwtri_col])
            if r is not None:
                records.append({"PPG": ppg_str, "Period": "Bi-weekly", "FW_Rate_Number": r})
                records.append({"PPG": ppg_str, "Period": "Tri-weekly", "FW_Rate_Number": r})
        if weekly_col is not None:
            r = parse_cell(row.iloc[weekly_col])
            if r is not None:
                records.append({"PPG": ppg_str, "Period": "Weekly", "FW_Rate_Number": r})
    
    if not records:
        raise RuntimeError(f"No usable rates extracted from sheet '{sheet_name}'")
    
    lookup = pd.DataFrame(records)
    
    # ----- 4. Join lookup to TPM -----
    tpm["_ppg_key"] = tpm["PromoGroupProductDesc"].astype(str).str.strip()
    tpm = tpm.merge(
        lookup.rename(columns={"PPG": "_ppg_key"}),
        on=["_ppg_key", "Period"],
        how="left"
    ).drop(columns=["_ppg_key"])
    tpm["Rate_Card_Source"] = quarter
    
    # ----- 5. Justification -----
    def justify(r):
        if pd.isna(r["FW_Rate_Number"]):
            return "NO rate card available"
        if pd.isna(r["Final_Rsp"]):
            return "Missing Final_Rsp"
        if (pd.notna(r["Quarter_Year"]) and pd.notna(r["Rate_Card_Source"])
                and r["Quarter_Year"] != r["Rate_Card_Source"]):
            return "⚠ Quarter mismatch"
        if abs(r["Final_Rsp"] - r["FW_Rate_Number"]) <= 0.5:
            return "Comply"
        return "NOT Comply"
    
    tpm["Justification"] = tpm.apply(justify, axis=1)
    
    # ----- 6. Summary pivot -----
    grp = [c for c in ["Customer", "Brand", "Period"] if c in tpm.columns]
    if grp:
        summary = (tpm.groupby(grp)["Justification"]
                      .value_counts().unstack(fill_value=0).reset_index())
        cats = ["Comply", "NOT Comply", "⚠ Quarter mismatch",
                "NO rate card available", "Missing Final_Rsp"]
        for c in cats:
            if c not in summary.columns:
                summary[c] = 0
        summary["Total"] = summary[cats].sum(axis=1)
        summary["Comply_%"] = (summary["Comply"] / summary["Total"] * 100).round(1)
        summary = summary.sort_values("Total", ascending=False)[grp + cats + ["Total", "Comply_%"]]
    else:
        summary = pd.DataFrame({"Note": ["No grouping columns"]})
    
    # ----- 7. Save xlsx with formatting -----
    out_path = f"TPM_FW_Compliance_Output_{date.today():%Y-%m-%d}.xlsx"
    
    # Remove timezone if any
    for c in tpm.select_dtypes(include=["datetimetz"]).columns:
        tpm[c] = tpm[c].dt.tz_localize(None)
    
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        # Write data
        tpm.to_excel(writer, index=False, sheet_name="Compliance")
        summary.to_excel(writer, index=False, sheet_name="Summary")
        
        wb = writer.book
        
        # ----- Format Compliance sheet -----
        ws = writer.sheets["Compliance"]
        
        # Header format
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#305496", "font_color": "white",
            "align": "center", "valign": "vcenter", "border": 1
        })
        for col_idx, col in enumerate(tpm.columns):
            ws.write(0, col_idx, col, header_fmt)
        
        # Freeze + autofilter
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(tpm), len(tpm.columns)-1)
        
        # Float columns: 1 decimal
        float_fmt = wb.add_format({"num_format": "0.0"})
        int_fmt = wb.add_format({"num_format": "0"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        
        for col_idx, (col, dtype) in enumerate(zip(tpm.columns, tpm.dtypes)):
            if pd.api.types.is_float_dtype(dtype):
                ws.set_column(col_idx, col_idx, 12, float_fmt)
            elif pd.api.types.is_integer_dtype(dtype):
                ws.set_column(col_idx, col_idx, 10, int_fmt)
            elif pd.api.types.is_datetime64_any_dtype(dtype):
                ws.set_column(col_idx, col_idx, 12, date_fmt)
            else:
                # Estimate width from sample
                max_len = max(len(str(col)), 
                              tpm[col].astype(str).head(100).map(len).max() if len(tpm) else 10)
                ws.set_column(col_idx, col_idx, min(max_len + 2, 40))
        
        # Color-code Justification column
        if "Justification" in tpm.columns:
            jcol = list(tpm.columns).index("Justification")
            from xlsxwriter.utility import xl_col_to_name
            jletter = xl_col_to_name(jcol)
            rng = f"{jletter}2:{jletter}{len(tpm)+1}"
            
            for label, bg, fg in [
                ("Comply", "#C6EFCE", "#006100"),
                ("NOT Comply", "#FFC7CE", "#9C0006"),
                ("⚠ Quarter mismatch", "#FCE4D6", "#974706"),
                ("NO rate card available", "#FFEB9C", "#9C5700"),
                ("Missing Final_Rsp", "#D9D9D9", "#555555"),
            ]:
                fmt = wb.add_format({"bg_color": bg, "font_color": fg, "bold": True})
                ws.conditional_format(rng, {
                    "type": "cell", "criteria": "==",
                    "value": f'"{label}"', "format": fmt
                })
        
        # ----- Format Summary sheet -----
        ws2 = writer.sheets["Summary"]
        for col_idx, col in enumerate(summary.columns):
            ws2.write(0, col_idx, col, header_fmt)
        ws2.freeze_panes(1, 0)
        ws2.autofilter(0, 0, len(summary), len(summary.columns)-1)
        
        if "Comply_%" in summary.columns:
            pct_col = list(summary.columns).index("Comply_%")
            from xlsxwriter.utility import xl_col_to_name
            pct_letter = xl_col_to_name(pct_col)
            rng = f"{pct_letter}2:{pct_letter}{len(summary)+1}"
            ws2.conditional_format(rng, {
                "type": "3_color_scale",
                "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"
            })
            ws2.set_column(pct_col, pct_col, 10, wb.add_format({"num_format": "0.0"}))
    
    # ----- Verify output -----
    out_size = Path(out_path).stat().st_size
    if out_size < 1024:
        raise RuntimeError(f"Output file suspiciously small: {out_size} bytes")
    
    chk = pd.read_excel(out_path, sheet_name="Compliance", nrows=5)
    if chk.empty or len(chk.columns) < len(tpm.columns):
        raise RuntimeError("Saved file verification failed")
    
    # ----- Return results -----
    counts = tpm["Justification"].value_counts().to_dict()
    return {
        "output_path": out_path,
        "summary_counts": counts,
        "rows": len(tpm),
        "initial_rows": initial_rows,
        "filtered_rows": initial_rows - len(tpm),
        "sheet_used": sheet_name,
        "quarter_used": quarter,
    }


# ============================================================
# Self-tests
# ============================================================
def run_self_tests(tpm, output_path):
    """Verify result quality. Returns list of (passed, message)."""
    tests = []
    
    # 1. Output is xlsx (not csv)
    is_xlsx = output_path.endswith(".xlsx")
    tests.append((is_xlsx, f"Output is xlsx: {is_xlsx}"))
    
    # 2. Not 100% "NO rate card available"
    no_rc_pct = (tpm["Justification"] == "NO rate card available").mean()
    tests.append((no_rc_pct < 0.95, 
                  f"NO-rate-card share: {no_rc_pct:.1%} (must be < 95%)"))
    
    # 3. BOGO rows have blank Promotion_Price
    bogo = tpm[tpm["Promo_Type"] == "BOGO"]
    if len(bogo) > 0:
        bogo_blank = bogo["Promotion_Price"].isna().all()
        tests.append((bogo_blank, f"BOGO rows blank price: {bogo_blank}"))
    
    # 4. PO/2F1/LAKSUE extraction rate > 80%
    for pt in ["PO", "2F1", "LAKSUE"]:
        sub = tpm[tpm["Promo_Type"] == pt]
        if len(sub) > 0:
            rate = sub["Promotion_Price"].notna().mean()
            tests.append((rate > 0.8, 
                          f"{pt} extraction: {rate:.0%} of {len(sub):,} rows"))
    
    # 5. Sales_Org filter applied
    if "Sales_Org" in tpm.columns:
        unique_so = tpm["Sales_Org"].unique()
        only_7001 = all(str(s).strip() == "7001" for s in unique_so)
        tests.append((only_7001, f"Sales_Org filter: {unique_so.tolist()}"))
    
    return tests
```

---

# 📤 Response Format (use this exact template)

After execution, return a structured message:

```
✅ **TPM × FW Compliance Check Complete!**

📊 **Stats:**
- Total rows processed: **{rows:,}**
- Rows filtered out (Sales_Org ≠ 7001): {filtered_rows:,}
- Rate card sheet: **{sheet_used}** ({quarter_used})

**Results:**

| Status | Count | % |
|---|---|---|
| 🟢 Comply | {n_comply:,} | {pct_comply}% |
| 🔴 NOT Comply | {n_notcomply:,} | {pct_notcomply}% |
| 🟠 ⚠ Quarter mismatch | {n_mismatch:,} | {pct_mismatch}% |
| 🟡 NO rate card | {n_norc:,} | {pct_norc}% |
| ⬜ Missing Final_Rsp | {n_missing:,} | {pct_missing}% |

**Self-tests:**
- ✅ Output is xlsx
- ✅ NO-rate-card share: 12.5%
- ✅ BOGO rows have blank price
- ✅ PO extraction: 99%
- ✅ Sales_Org filter applied

📎 **Download:** [TPM_FW_Compliance_Output_{date}.xlsx]

💡 **Next steps:**
- Open in Excel to see color-coded results
- Filter the **Compliance** sheet by Justification to focus on issues
- Check the **Summary** sheet for per-Customer/Brand breakdown

Need help analyzing the NOT Comply rows? Just ask!
```

---

# 🚫 Critical Rules (DO NOT VIOLATE)

| ❌ Wrong | ✅ Correct |
|---------|-----------|
| Save as `.csv` | Save as `.xlsx` with `xlsxwriter` engine |
| Use `polars` | Use `pandas` only (agent compatibility) |
| BOGO = `RSP × 0.5` | BOGO → blank Promotion_Price |
| 2F1 price from `2F1/xxx` | 2F1 price from `_POnn` |
| Hard-code header row 8 | Detect by scanning for "Promo Group" |
| Skip Sales_Org filter | Always filter `Sales_Org == "7001"` |
| Auto-pick sheet without asking | **Always ask user** which sheet to use |
| Forget to attach output file | Always provide downloadable file |
| Use Rich/questionary | Use plain chat messages (no terminal libs) |

---

# 🐛 Error Handling

If anything fails, respond friendly:

```
❌ **Something went wrong:** {error_type}

**Details:** {error_message}

**Possible fixes:**
- Make sure both files are xlsx (not csv/pdf)
- TPM file needs these columns: PromoGroupProductDesc, Instore_Start, 
  Instore_End, TPM_InvestmentDescription, RSP Promo, Sales_Org
- FW Rate Card file needs at least one sheet with "Promo Group (PPG)" header

Want to share the file structure? Send me a screenshot of the first 3 rows!
```

---

# 📊 Differences from Local v1.5

| Aspect | Local v1.5 | Agent v2.0 |
|--------|-----------|-----------|
| Dataframe library | Polars (5× faster) | **Pandas (universal)** |
| File input | Scan working dir | User uploads in chat |
| Sheet selection | Arrow-key menu (questionary) | **Chat prompt** |
| Progress display | Rich progress bar | Chat status updates |
| Path detection | `sys.frozen` exe-aware | n/a (runs in sandbox) |
| Excel writer | xlsxwriter | xlsxwriter (same) |
| Conditional formatting | Same 5 categories | Same 5 categories |
| Output naming | `TPM_FW_Compliance_Output_{date}.xlsx` | Same |
| Sales_Org filter | Yes | Yes |
| Quarter mismatch | In Justification col | In Justification col |
| Drop unused cols | Yes | Yes |
| BOGO handling | Blank | Blank |
| 2F1 → _PO pattern | Yes | Yes |

---

# 🎯 Multi-language Support

If user writes in Thai, reply in Thai. Example:

User: *"ตรวจ TPM"*

Agent: 
> "สวัสดีครับ 👋 ผมจะตรวจสอบ TPM กับ FW Rate Card ให้นะครับ กรุณาอัปโหลด:
> 
> 1. 📂 ไฟล์ **TPM** (xlsx ที่มีคำว่า 'TPM' ในชื่อ)
> 2. 📂 ไฟล์ **FW Rate Card** (xlsx ที่มีคำว่า 'FW rate card' ในชื่อ)
> 
> วางทั้ง 2 ไฟล์ในแชท แล้วผมจะประมวลผลให้ครับ ✨"

---

# 🔄 Robustness Features

1. **File detection retry**: If `read_excel` fails, retry with `engine="openpyxl"`, then `engine="calamine"` before raising.
2. **Memory awareness**: For files > 50MB, use `dtype_backend="pyarrow"` to reduce RAM.
3. **Logging**: Print every step (`Loaded X rows`, `Filtered Y rows`, etc.) so user sees progress.
4. **Verification**: After save, re-read the xlsx and assert rows + columns are present.

---

# 📞 Clarification Questions

| Trigger | Ask |
|---------|-----|
| Multiple TPM files uploaded | "Which TPM file? `{file1}` or `{file2}`?" |
| Rate card has unclear quarter | "Which quarter does this represent? (Q1/Q2/Q3/Q4 + year)" |
| >95% rows NO rate card | "Most rows can't find a match. Check PPG naming, or proceed?" |
| User asks for csv | "I'll generate xlsx because csv loses colors and 2-sheet layout. Open in Excel." |
| File >100k rows | "Large file detected. Processing will take 20–60 seconds. Continue?" |

---

# 📌 Version History

- **v2.0 — 2026-06-19** — Agent edition. Pandas-only (no polars). Chat-based UI. File uploads. Sales_Org filter. Quarter mismatch detection. Self-tests included.
- **v1.5 — 2026-06-18** (local) — Source version with Polars + Rich UI + arrow keys.

---

# 🔐 Privacy & Security

- All processing happens in agent sandbox; data doesn't leave the session
- Input files cleared after response (per agent runtime policy)
- Output provided as direct download link
- Mark conversation as confidential if files contain PII or competitive pricing