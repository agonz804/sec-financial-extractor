import streamlit as st
import pandas as pd
import requests
import time
import re
import io
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SEC Financial Data Extractor", layout="wide")

HEADERS = {"User-Agent": "FinancialDataExtractor contact@example.com"}

# ── EDGAR helpers ──────────────────────────────────────────────────────────────

def get_cik(ticker: str) -> str | None:
    r = requests.get("https://www.sec.gov/files/company_tickers.json", headers=HEADERS, timeout=15)
    for item in r.json().values():
        if item["ticker"].upper() == ticker.upper():
            return str(item["cik_str"]).zfill(10)
    return None

def get_company_name(cik: str) -> str:
    r = requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=HEADERS, timeout=15)
    return r.json().get("name", "Unknown")

def get_xbrl_facts(cik: str) -> dict:
    r = requests.get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json", headers=HEADERS, timeout=30)
    return r.json().get("facts", {})

def get_filings_index(cik: str, form_type: str) -> list[dict]:
    r = requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=HEADERS, timeout=15)
    data = r.json()
    filings = data.get("filings", {}).get("recent", {})
    results = []
    forms   = filings.get("form", [])
    dates   = filings.get("filingDate", [])
    accs    = filings.get("accessionNumber", [])
    docs    = filings.get("primaryDocument", [])
    for i, form in enumerate(forms):
        if form == form_type:
            results.append({"date": dates[i], "accession": accs[i].replace("-", ""), "primary_doc": docs[i]})
    return results

# ── XBRL: pull every concept the company actually filed ────────────────────────

# These concept names are so generic / dimensional they add noise rather than value
SKIP_CONCEPTS = {
    "EntityCommonStockSharesOutstanding", "EntityPublicFloat", "EntityNumberOfEmployees",
    "DocumentFiscalYearFocus", "DocumentFiscalPeriodFocus", "DocumentPeriodEndDate",
    "EntityRegistrantName", "EntityCentralIndexKey", "TradingSymbol",
    "CommonStockSharesAuthorized", "CommonStockParOrStatedValuePerShare",
    "CommonStockSharesIssued", "PreferredStockSharesAuthorized",
    "PreferredStockSharesIssued", "PreferredStockSharesOutstanding",
}

# Concepts that represent shares (not dollars) — keep as actual count / MM shares
SHARE_CONCEPTS = {
    "WeightedAverageNumberOfSharesOutstandingBasic",
    "WeightedAverageNumberOfDilutedSharesOutstanding",
    "CommonStockSharesOutstanding",
    "CommonStockSharesIssued",
}

# Concepts reported per-share in USD/shares
PER_SHARE_CONCEPTS = {
    "EarningsPerShareBasic",
    "EarningsPerShareDiluted",
    "BookValuePerShareBasic",
    "DividendsCommonStockCash",
}


def human_label(concept_name: str) -> str:
    """Convert CamelCase XBRL concept name to readable label, preserving acronyms."""
    # Split on capital letters, but keep sequences like 'PP&E', 'R&D', 'SGA' together
    s = re.sub(r"([A-Z][a-z]+)", r" \1", re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", concept_name))
    return s.strip()


def extract_all_concepts(facts: dict, is_annual: bool, years: int, cutoff_date: str) -> dict[str, dict[str, float | None]]:
    """
    Returns {concept_label: {period: value}} for every USD/shares concept
    the company filed in 10-K or 10-Q forms, converted to $MM where applicable.
    """
    valid_forms = ("10-K", "20-F") if is_annual else ("10-Q", "6-K")
    result: dict[str, dict] = {}

    for ns in ["us-gaap", "dei", "invest"]:
        ns_facts = facts.get(ns, {})
        for concept, concept_data in ns_facts.items():
            if concept in SKIP_CONCEPTS:
                continue

            units = concept_data.get("units", {})

            # Determine which unit bucket to use
            if "USD" in units:
                raw_entries = units["USD"]
                unit_type = "USD"
            elif "USD/shares" in units:
                raw_entries = units["USD/shares"]
                unit_type = "USD/shares"
            elif "shares" in units:
                raw_entries = units["shares"]
                unit_type = "shares"
            else:
                continue  # skip non-financial concepts (pure text, dates, etc.)

            # Filter to right form type and date range
            filtered = [
                e for e in raw_entries
                if e.get("form") in valid_forms
                and e.get("end", "") >= cutoff_date
                and e.get("val") is not None
            ]
            if not filtered:
                continue

            # Build period → value map, keeping latest-filed value per period
            period_map: dict[str, dict] = {}
            for e in filtered:
                end  = e["end"]
                filed = e.get("filed", "")
                val   = e["val"]

                if is_annual:
                    # For annual: only keep entries spanning ~12 months
                    start = e.get("start", "")
                    if start:
                        try:
                            span_days = (pd.Timestamp(end) - pd.Timestamp(start)).days
                            if not (300 <= span_days <= 400):
                                continue
                        except Exception:
                            pass
                    period_key = end[:4]  # fiscal year
                else:
                    # For quarterly: only keep entries spanning ~3 months
                    start = e.get("start", "")
                    if start:
                        try:
                            span_days = (pd.Timestamp(end) - pd.Timestamp(start)).days
                            if not (60 <= span_days <= 110):
                                continue
                        except Exception:
                            pass
                    period_key = end  # YYYY-MM-DD

                if period_key not in period_map or filed > period_map[period_key]["filed"]:
                    period_map[period_key] = {"val": val, "filed": filed}

            if not period_map:
                continue

            # Convert to $MM where applicable
            converted: dict[str, float | None] = {}
            for period_key, entry in period_map.items():
                val = entry["val"]
                if unit_type == "USD" and concept not in PER_SHARE_CONCEPTS:
                    converted[period_key] = round(val / 1_000_000, 3)
                elif unit_type == "shares" or concept in SHARE_CONCEPTS:
                    converted[period_key] = round(val / 1_000_000, 3)  # MM shares
                else:
                    # USD/shares (EPS) — keep as-is
                    converted[period_key] = round(val, 4)

            label = human_label(concept)
            # If same label from different namespaces, prefer us-gaap
            if label not in result or ns == "us-gaap":
                result[label] = converted

    return result


def get_sorted_periods(data: dict[str, dict], is_annual: bool) -> list[str]:
    all_periods: set[str] = set()
    for period_data in data.values():
        all_periods.update(period_data.keys())
    return sorted(all_periods, reverse=True)


# ── Statement classifier: sort concepts into IS / BS / CF buckets ─────────────

# These keyword fragments help assign concepts to the right statement tab
IS_KEYWORDS = [
    "revenue", "revenues", "sales", "royalt", "income", "loss", "expense", "cost",
    "profit", "gross", "operating", "ebitda", "interest", "tax", "earning", "margin",
    "amortization", "depreciation", "impairment", "restructur", "dividend",
]
BS_KEYWORDS = [
    "asset", "liabilit", "equity", "cash", "receivable", "inventory", "payable",
    "debt", "borrowing", "goodwill", "intangible", "investment", "deferred",
    "stockholder", "retained", "treasury", "accumulated", "capital", "prepaid",
    "property", "plant", "equipment", "lease", "right.of.use",
]
CF_KEYWORDS = [
    "cash provided", "cash used", "operating activit", "investing activit",
    "financing activit", "proceeds from", "payment", "repayment", "purchase of",
    "acquisition", "repurchase", "issuance", "capital expenditure", "capex",
    "free cash",
]

def classify_concept(label: str) -> str:
    """Return 'IS', 'BS', 'CF', or 'OTHER'."""
    lower = label.lower()
    cf_score = sum(1 for kw in CF_KEYWORDS if kw in lower)
    bs_score = sum(1 for kw in BS_KEYWORDS if kw in lower)
    is_score = sum(1 for kw in IS_KEYWORDS if kw in lower)
    scores = {"CF": cf_score, "BS": bs_score, "IS": is_score}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "OTHER"


# ── Excel builder ──────────────────────────────────────────────────────────────

HDR_FILL    = PatternFill("solid", start_color="1F3864")
SEC_FILL    = PatternFill("solid", start_color="2E75B6")
ALT_FILL    = PatternFill("solid", start_color="EBF3FB")
WHITE_FILL  = PatternFill("solid", start_color="FFFFFF")
HDR_FONT    = Font(name="Arial", bold=True, color="FFFFFF", size=10)
SEC_FONT    = Font(name="Arial", bold=True, color="FFFFFF", size=9)
LABEL_FONT  = Font(name="Arial", size=9)
BOLD_FONT   = Font(name="Arial", bold=True, size=9)
DATA_FONT   = Font(name="Arial", size=9, color="00008B")


def write_data_sheet(wb: Workbook, sheet_name: str, data: dict[str, dict],
                     periods: list[str], company_name: str, period_label: str):
    ws = wb.create_sheet(sheet_name[:31])
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B3"

    ncols = len(periods)

    # Row 1 — title
    ws.row_dimensions[1].height = 20
    ws.cell(row=1, column=1).value = f"{company_name}  |  {sheet_name}"
    ws.cell(row=1, column=1).font  = Font(name="Arial", bold=True, size=11, color="1F3864")
    note_col = ncols + 2
    ws.cell(row=1, column=note_col).value = "Dollar values in $MM  |  EPS in $/share  |  Share counts in MM shares  |  Source: SEC EDGAR XBRL (as reported)"
    ws.cell(row=1, column=note_col).font  = Font(name="Arial", size=8, color="808080", italic=True)

    # Row 2 — period headers
    ws.row_dimensions[2].height = 16
    ws.cell(row=2, column=1).value = period_label
    ws.cell(row=2, column=1).font  = HDR_FONT
    ws.cell(row=2, column=1).fill  = HDR_FILL
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center")
    for i, p in enumerate(periods):
        c = i + 2
        ws.cell(row=2, column=c).value     = p
        ws.cell(row=2, column=c).font      = HDR_FONT
        ws.cell(row=2, column=c).fill      = HDR_FILL
        ws.cell(row=2, column=c).alignment = Alignment(horizontal="center", vertical="center")

    # Data rows
    current_row = 3
    for i, (label, period_data) in enumerate(sorted(data.items())):
        ws.row_dimensions[current_row].height = 14
        fill = ALT_FILL if i % 2 == 0 else WHITE_FILL

        ws.cell(row=current_row, column=1).value     = label
        ws.cell(row=current_row, column=1).font      = LABEL_FONT
        ws.cell(row=current_row, column=1).fill      = fill
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="left", indent=1)

        for j, p in enumerate(periods):
            c   = j + 2
            val = period_data.get(p)
            cell = ws.cell(row=current_row, column=c)
            cell.fill      = fill
            cell.alignment = Alignment(horizontal="right")
            if val is not None:
                cell.value = val
                cell.font  = DATA_FONT
                # Format: per-share gets 2 decimals, everything else 1 decimal
                if abs(val) < 100 and val != int(val):
                    cell.number_format = '#,##0.00;(#,##0.00);"-"'
                else:
                    cell.number_format = '#,##0.0;(#,##0.0);"-"'
            else:
                cell.value = "—"
                cell.font  = Font(name="Arial", size=9, color="BBBBBB")

        current_row += 1

    ws.column_dimensions["A"].width = 52
    for i in range(len(periods)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 13


def write_raw_table_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, title: str):
    ws = wb.create_sheet(sheet_name[:31])
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1).value = title
    ws.cell(row=1, column=1).font  = Font(name="Arial", bold=True, size=10, color="1F3864")
    for c_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(row=2, column=c_idx)
        cell.value     = str(col)
        cell.font      = HDR_FONT
        cell.fill      = HDR_FILL
        cell.alignment = Alignment(horizontal="center")
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx + 3, column=c_idx)
            cell.value = val
            cell.font  = Font(name="Arial", size=9)
            cell.fill  = ALT_FILL if r_idx % 2 == 0 else WHITE_FILL
    for c_idx in range(1, len(df.columns) + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 24


# ── Segment / KPI HTML extraction ─────────────────────────────────────────────

JUNK_PATTERN = re.compile(
    r"table of contents|exhibit index|incorporated herein by reference|"
    r"certification of chief|pursuant to rule 13a|pursuant to section 906|"
    r"instance document|taxonomy extension|inline xbrl|"
    r"ernst.*young|deloitte|kpmg|pricewaterhousecoopers|/s/ |"
    r"trading arrangement|rule 10b5|shares to be sold|expiration date|"
    r"accounting standard|fasb issued|asc 842|adoption method|"
    r"bylaws|certificate of incorporation|indenture.*trustee",
    re.IGNORECASE
)

def is_useful_table(df: pd.DataFrame) -> bool:
    if df.shape[0] < 3 or df.shape[1] < 2:
        return False
    sample = " ".join(str(v) for v in df.iloc[:4].values.flatten() if v)
    if JUNK_PATTERN.search(sample):
        return False
    all_text = " ".join(str(v) for v in df.values.flatten() if v)
    if not re.search(r"\d{2,}", all_text):
        return False
    non_empty = sum(1 for v in df.values.flatten() if str(v).strip() not in ("", "nan", "None"))
    return non_empty / max(df.size, 1) >= 0.15


def fetch_filing_html(cik: str, accession: str, primary_doc: str) -> str | None:
    url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession}/{primary_doc}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            return r.text
    except Exception:
        pass
    return None


def extract_tables_from_html(html: str, keywords: list[str]) -> list[pd.DataFrame]:
    soup = BeautifulSoup(html, "html.parser")
    results = []
    seen = set()
    for kw in keywords:
        pattern = re.compile(r"\b" + re.escape(kw) + r"\b", re.IGNORECASE)
        for match in soup.find_all(string=pattern)[:5]:
            parent = match.parent
            for _ in range(8):
                if parent is None:
                    break
                table = parent.find_next("table")
                if table:
                    try:
                        dfs = pd.read_html(str(table))
                        if dfs:
                            df = dfs[0]
                            fp = str(df.values.tolist())[:300]
                            if fp not in seen and is_useful_table(df):
                                seen.add(fp)
                                results.append(df)
                    except Exception:
                        pass
                    break
                parent = parent.parent
    return results


def fetch_segment_data(cik: str, filings: list[dict], max_filings: int = 8) -> list[tuple[str, pd.DataFrame]]:
    keywords = [
        "segment revenue", "segment information", "revenue by segment",
        "geographic", "revenue by region", "revenue by geography",
        "disaggregated revenue", "revenue disaggregation",
        "customer concentration", "significant customer", "major customer",
        "royalt", "product sales", "collaborative",
        "subscribers", "active users", "monthly active", "annual recurring",
        "units sold", "same-store", "comparable store",
        "backlog", "bookings", "net revenue retention",
        "key performance", "operating metric",
    ]
    results = []
    seen_tables: set[str] = set()
    for filing in filings[:max_filings]:
        html = fetch_filing_html(cik, filing["accession"], filing["primary_doc"])
        if not html:
            continue
        for df in extract_tables_from_html(html, keywords):
            key = str(df.values.tolist())[:300]
            if key not in seen_tables:
                seen_tables.add(key)
                results.append((filing["date"], df.dropna(how="all").fillna("")))
        time.sleep(0.3)
    return results


# ── Streamlit UI ───────────────────────────────────────────────────────────────

def main():
    st.title("📊 SEC Financial Data Extractor")
    st.markdown("Pulls **every line item exactly as reported** from SEC EDGAR XBRL filings.")

    with st.sidebar:
        st.header("Settings")
        ticker  = st.text_input("Stock Ticker", placeholder="e.g. AAPL, MSFT, HALO").strip().upper()
        years   = st.slider("Years of History", 5, 20, 15)
        include_segments = st.checkbox("Extract Segment / KPI Tables (slower)", value=True)
        max_seg_filings  = st.slider("# Filings to Scan for Segments", 4, 20, 8)
        run_btn = st.button("🚀 Extract Data", type="primary", use_container_width=True)

    if not run_btn or not ticker:
        st.info("Enter a ticker in the sidebar and click **Extract Data** to begin.")
        st.markdown("""
**What this tool extracts (as reported — no mapping, no aggregation):**
- ✅ Every XBRL line item from Income Statement, Balance Sheet, Cash Flow
- ✅ Company's own labels, exactly as filed with the SEC
- ✅ Quarterly and Annual tabs for each statement
- ✅ Segment / geographic / KPI tables (best-effort from filing HTML)
- ✅ All dollar values in $MM, share counts in MM shares, EPS in $/share
        """)
        return

    progress = st.progress(0)
    status   = st.empty()

    try:
        status.text("🔍 Looking up CIK...")
        cik = get_cik(ticker)
        if not cik:
            st.error(f"Could not find CIK for '{ticker}'. Check the ticker and try again.")
            return
        progress.progress(5)

        company_name = get_company_name(cik)
        st.success(f"Found: **{company_name}** (CIK: {int(cik)})")

        status.text("📥 Downloading XBRL company facts...")
        facts = get_xbrl_facts(cik)
        progress.progress(20)

        cutoff_annual  = str(pd.Timestamp.now().year - years)
        cutoff_quarter = str((pd.Timestamp.now() - pd.DateOffset(years=years)).date())

        status.text("🔢 Extracting annual data (as reported)...")
        annual_data = extract_all_concepts(facts, is_annual=True,  years=years, cutoff_date=cutoff_annual + "-01-01")
        annual_periods = get_sorted_periods(annual_data, is_annual=True)
        progress.progress(40)

        status.text("🔢 Extracting quarterly data (as reported)...")
        qtr_data    = extract_all_concepts(facts, is_annual=False, years=years, cutoff_date=cutoff_quarter)
        qtr_periods = get_sorted_periods(qtr_data, is_annual=False)
        progress.progress(60)

        # Split into IS / BS / CF / OTHER buckets for each frequency
        def split_buckets(data: dict) -> dict[str, dict]:
            buckets: dict[str, dict] = {"IS": {}, "BS": {}, "CF": {}, "OTHER": {}}
            for label, period_data in data.items():
                buckets[classify_concept(label)][label] = period_data
            return buckets

        annual_buckets = split_buckets(annual_data)
        qtr_buckets    = split_buckets(qtr_data)

        segment_tables = []
        if include_segments:
            status.text("🔍 Fetching filings for segment/KPI data...")
            all_filings = sorted(
                get_filings_index(cik, "10-K") + get_filings_index(cik, "10-Q"),
                key=lambda x: x["date"], reverse=True
            )
            segment_tables = fetch_segment_data(cik, all_filings, max_seg_filings)
        progress.progress(80)

        status.text("📝 Building Excel workbook...")
        wb = Workbook()
        wb.remove(wb.active)

        sheet_configs = [
            ("Annual — Income Stmt",    annual_buckets["IS"],    list(reversed(annual_periods)),  "Fiscal Year",    True),
            ("Annual — Balance Sheet",  annual_buckets["BS"],    list(reversed(annual_periods)),  "Fiscal Year",    True),
            ("Annual — Cash Flow",      annual_buckets["CF"],    list(reversed(annual_periods)),  "Fiscal Year",    True),
            ("Annual — Other",          annual_buckets["OTHER"], list(reversed(annual_periods)),  "Fiscal Year",    True),
            ("Quarterly — Income Stmt", qtr_buckets["IS"],       qtr_periods[:60],               "Quarter Ended",  False),
            ("Quarterly — Balance Sht", qtr_buckets["BS"],       qtr_periods[:60],               "Quarter Ended",  False),
            ("Quarterly — Cash Flow",   qtr_buckets["CF"],       qtr_periods[:60],               "Quarter Ended",  False),
            ("Quarterly — Other",       qtr_buckets["OTHER"],    qtr_periods[:60],               "Quarter Ended",  False),
        ]

        for sheet_name, data, periods, period_label, is_annual in sheet_configs:
            if data and periods:
                write_data_sheet(wb, sheet_name, data, periods, company_name, period_label)

        for idx, (filing_date, df) in enumerate(segment_tables[:25]):
            write_raw_table_sheet(wb, f"Seg-KPI {filing_date} ({idx+1})", df,
                                  f"Extracted Table — Filing: {filing_date}")

        progress.progress(95)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        progress.progress(100)
        status.text("✅ Done!")

        fname = f"{ticker}_financials_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️ Download Excel File",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # Preview
        st.subheader("Preview — Annual Income Statement concepts")
        if annual_buckets["IS"]:
            preview = []
            show_periods = list(reversed(annual_periods))[:10]
            for label, pd_data in sorted(annual_buckets["IS"].items()):
                row = {"Line Item": label}
                for p in show_periods:
                    row[p] = pd_data.get(p, "—")
                preview.append(row)
            st.dataframe(pd.DataFrame(preview).set_index("Line Item"), use_container_width=True)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Annual IS concepts",   len(annual_buckets["IS"]))
        col2.metric("Annual BS concepts",   len(annual_buckets["BS"]))
        col3.metric("Annual CF concepts",   len(annual_buckets["CF"]))
        col4.metric("Seg/KPI tables found", len(segment_tables))

        if segment_tables:
            with st.expander(f"Preview first Seg/KPI table ({segment_tables[0][0]})"):
                st.dataframe(segment_tables[0][1], use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)


if __name__ == "__main__":
    main()
