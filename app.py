import streamlit as st
import pandas as pd
import requests
import time
import re
import io
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SEC Financial Data Extractor", layout="wide")

HEADERS = {"User-Agent": "FinancialDataExtractor contact@example.com"}

# ── EDGAR helpers ─────────────────────────────────────────────────────────────

def get_cik(ticker: str) -> str | None:
    url = "https://www.sec.gov/files/company_tickers.json"
    r = requests.get(url, headers=HEADERS, timeout=15)
    data = r.json()
    ticker_upper = ticker.upper()
    for item in data.values():
        if item["ticker"].upper() == ticker_upper:
            return str(item["cik_str"]).zfill(10)
    return None


def get_company_name(cik: str) -> str:
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS, timeout=15)
    return r.json().get("name", "Unknown")


def get_xbrl_facts(cik: str) -> dict:
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS, timeout=30)
    return r.json().get("facts", {})


def get_filings_index(cik: str, form_type: str) -> list[dict]:
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS, timeout=15)
    data = r.json()
    filings = data.get("filings", {}).get("recent", {})
    results = []
    forms = filings.get("form", [])
    dates = filings.get("filingDate", [])
    accessions = filings.get("accessionNumber", [])
    primary_docs = filings.get("primaryDocument", [])
    for i, form in enumerate(forms):
        if form == form_type:
            results.append({
                "date": dates[i],
                "accession": accessions[i].replace("-", ""),
                "primary_doc": primary_docs[i],
            })
    # Also check older filings pages
    older_url = f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type={form_type}&dateb=&owner=include&count=100&search_text="
    return results


# ── XBRL concept extraction ────────────────────────────────────────────────────

INCOME_STATEMENT_CONCEPTS = {
    # Revenue — try all common variants; many companies use one but not others
    "Revenue": [
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueNet",
        "SalesRevenueGoodsNet",
        "SalesRevenueServicesNet",
        "RevenueNet",
        "TotalRevenues",
    ],
    "Cost of Revenue": [
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfGoodsSold",
        "CostOfServices",
        "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
    ],
    "Gross Profit": ["GrossProfit"],  # calculated as fallback if blank
    "Amortization of Intangibles": [
        "AmortizationOfIntangibleAssets",
        "AmortizationOfAcquiredIntangibleAssets",
    ],
    "R&D Expense": [
        "ResearchAndDevelopmentExpense",
        "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost",
    ],
    "SG&A Expense": [
        "SellingGeneralAndAdministrativeExpense",
        "GeneralAndAdministrativeExpense",
        "SellingAndMarketingExpense",
    ],
    "Other Operating Expense": [
        "OtherOperatingIncomeExpenseNet",
        "OtherCostAndExpenseOperating",
        "RestructuringCharges",
        "ImpairmentOfIntangibleAssetsExcludingGoodwill",
        "GoodwillImpairmentLoss",
    ],
    "Operating Income": ["OperatingIncomeLoss"],
    "Interest Expense": [
        "InterestExpense",
        "InterestAndDebtExpense",
        "InterestExpenseDebt",
    ],
    "Interest & Investment Income": [
        "InvestmentIncomeNonoperating",
        "InvestmentIncomeInterest",
        "InterestIncomeOperating",
        "NonoperatingIncomeExpense",
        "OtherNonoperatingIncomeExpense",
        "OtherNonoperatingIncome",
    ],
    "Pretax Income": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    ],
    "Income Tax": ["IncomeTaxExpenseBenefit"],
    "Net Income": ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
    "EPS Basic": ["EarningsPerShareBasic"],
    "EPS Diluted": ["EarningsPerShareDiluted"],
    "Shares Basic": ["WeightedAverageNumberOfSharesOutstandingBasic", "CommonStockSharesOutstanding"],
    "Shares Diluted": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
    "Depreciation & Amortization": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
        "Depreciation",
    ],
    "EBITDA (calc)": [],
}

BALANCE_SHEET_CONCEPTS = {
    "Cash & Equivalents": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "Cash",
    ],
    "Short-term Investments": [
        "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
        "ShortTermInvestments",
        "MarketableSecuritiesCurrent",
        "AvailableForSaleSecuritiesCurrent",
    ],
    "Accounts Receivable": [
        "AccountsReceivableNetCurrent",
        "ReceivablesNetCurrent",
    ],
    "Inventory": ["InventoryNet", "InventoryGross"],
    "Other Current Assets": [
        "OtherAssetsCurrent",
        "PrepaidExpenseAndOtherAssetsCurrent",
        "PrepaidExpenseCurrent",
    ],
    "Total Current Assets": ["AssetsCurrent"],
    "PP&E Net": ["PropertyPlantAndEquipmentNet"],
    "Goodwill": ["Goodwill"],
    "Intangible Assets": [
        "IntangibleAssetsNetExcludingGoodwill",
        "FiniteLivedIntangibleAssetsNet",
        "IndefiniteLivedIntangibleAssetsExcludingGoodwill",
    ],
    "Long-term Investments": [
        "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent",
        "LongTermInvestments",
        "MarketableSecuritiesNoncurrent",
    ],
    "Other Long-term Assets": [
        "OtherAssetsNoncurrent",
        "DeferredIncomeTaxAssetsNet",
    ],
    "Total Assets": ["Assets"],
    "Accounts Payable": ["AccountsPayableCurrent"],
    "Accrued Liabilities": [
        "AccruedLiabilitiesCurrent",
        "EmployeeRelatedLiabilitiesCurrent",
        "AccruedEmployeeBenefitsCurrent",
    ],
    "Short-term Debt": [
        "ShortTermBorrowings",
        "LongTermDebtCurrent",
        "DebtCurrent",
        "ConvertibleNotesPayableCurrent",
        "NotesPayableCurrent",
    ],
    "Deferred Revenue Current": [
        "DeferredRevenueCurrent",
        "ContractWithCustomerLiabilityCurrent",
    ],
    "Other Current Liabilities": ["OtherLiabilitiesCurrent"],
    "Total Current Liabilities": ["LiabilitiesCurrent"],
    "Long-term Debt": [
        "LongTermDebtNoncurrent",
        "LongTermDebt",
        "ConvertibleLongTermNotesPayable",
        "SeniorLongTermNotes",
        "LongTermNotesPayable",
    ],
    "Deferred Revenue LT": [
        "DeferredRevenueNoncurrent",
        "ContractWithCustomerLiabilityNoncurrent",
    ],
    "Other Long-term Liabilities": [
        "OtherLiabilitiesNoncurrent",
        "DeferredIncomeTaxLiabilitiesNet",
    ],
    "Total Liabilities": ["Liabilities"],
    "Common Stock & APIC": [
        "AdditionalPaidInCapital",
        "AdditionalPaidInCapitalCommonStock",
    ],
    "Retained Earnings": ["RetainedEarningsAccumulatedDeficit"],
    "Treasury Stock": [
        "TreasuryStockValue",
        "TreasuryStockCommonValue",
    ],
    "Accumulated OCI": ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"],
    "Total Stockholders Equity": [
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ],
    "Total Liabilities & Equity": ["LiabilitiesAndStockholdersEquity"],
}

CASH_FLOW_CONCEPTS = {
    "Net Income": ["NetIncomeLoss", "ProfitLoss"],
    "D&A (CF)": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
        "Depreciation",
    ],
    "Amortization of Intangibles (CF)": [
        "AmortizationOfIntangibleAssets",
        "AmortizationOfAcquiredIntangibleAssets",
    ],
    "Stock-Based Compensation": [
        "ShareBasedCompensation",
        "AllocatedShareBasedCompensationExpense",
        "ShareBasedCompensationExpense",
    ],
    "Deferred Income Taxes": [
        "DeferredIncomeTaxExpenseBenefit",
        "DeferredIncomeTaxesAndTaxCredits",
    ],
    "Changes in Working Capital": [
        "IncreaseDecreaseInOperatingCapital",
        "IncreaseDecreaseInOperatingLiabilities",
    ],
    "Other Operating Activities": [
        "OtherNoncashIncomeExpense",
        "OtherOperatingActivitiesCashFlowStatement",
    ],
    "Cash from Operations": ["NetCashProvidedByUsedInOperatingActivities"],
    "Capex": ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "Acquisitions": [
        "PaymentsToAcquireBusinessesNetOfCashAcquired",
        "PaymentsToAcquireBusinessesGross",
    ],
    "Purchases of Investments": [
        "PaymentsToAcquireAvailableForSaleSecurities",
        "PaymentsToAcquireAvailableForSaleSecuritiesDebt",
        "PaymentsToAcquireInvestments",
        "PaymentsToAcquireMarketableSecurities",
    ],
    "Sales/Maturities of Investments": [
        "ProceedsFromSaleAndMaturityOfAvailableForSaleSecurities",
        "ProceedsFromSaleOfAvailableForSaleSecurities",
        "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
        "ProceedsFromSaleAndMaturityOfMarketableSecurities",
        "ProceedsFromSaleMaturityAndCollectionOfInvestments",
    ],
    "Cash from Investing": ["NetCashProvidedByUsedInInvestingActivities"],
    "Debt Issuance": [
        "ProceedsFromIssuanceOfLongTermDebt",
        "ProceedsFromConvertibleDebt",
        "ProceedsFromIssuanceOfDebt",
        "ProceedsFromNotesPayable",
        "ProceedsFromDebtMaturingInMoreThanThreeMonths",
        "ProceedsFromIssuanceOfSeniorLongTermDebt",
    ],
    "Debt Repayment": [
        "RepaymentsOfLongTermDebt",
        "RepaymentsOfConvertibleDebt",
        "RepaymentsOfDebt",
        "RepaymentsOfNotesPayable",
        "RepaymentsOfDebtMaturingInMoreThanThreeMonths",
    ],
    "Share Repurchases": [
        "PaymentsForRepurchaseOfCommonStock",
        "PaymentsRelatedToTaxWithholdingForShareBasedCompensation",
    ],
    "Dividends Paid": [
        "PaymentsOfDividends",
        "PaymentsOfDividendsCommonStock",
        "PaymentsOfOrdinaryDividends",
    ],
    "Stock Issuance": [
        "ProceedsFromIssuanceOfCommonStock",
        "ProceedsFromStockOptionsExercised",
        "ProceedsFromIssuanceOfSharesUnderIncentiveAndShareBasedCompensationPlansIncludingStockOptions",
    ],
    "Cash from Financing": ["NetCashProvidedByUsedInFinancingActivities"],
    "Net Change in Cash": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "CashAndCashEquivalentsPeriodIncreaseDecrease",
        "NetCashProvidedByUsedInContinuingOperations",
    ],
    "Free Cash Flow (calc)": [],
}


def extract_concept(facts: dict, concepts: list[str], unit_filter: str = "USD") -> list[dict]:
    """Return all data points for the first matching concept."""
    for ns in ["us-gaap", "ifrs-full"]:
        ns_facts = facts.get(ns, {})
        for concept in concepts:
            if concept in ns_facts:
                units = ns_facts[concept].get("units", {})
                data = units.get(unit_filter) or units.get("shares") or []
                return [d for d in data if d.get("form") in ("10-K", "10-Q", "20-F", "6-K")]
    return []


def to_millions(value, unit_filter="USD"):
    if value is None:
        return None
    if unit_filter in ("USD", "shares"):
        return round(value / 1_000_000, 3)
    # USD/shares (EPS) — keep as dollars per share, no conversion
    return round(value, 4)


def build_period_map(raw_data: list[dict], is_annual: bool, unit_filter: str = "USD") -> dict:
    """Convert raw EDGAR data points into {period_label: value} dict."""
    period_map = {}
    for d in raw_data:
        form = d.get("form", "")
        if is_annual and form not in ("10-K", "20-F"):
            continue
        if not is_annual and form not in ("10-Q", "6-K"):
            continue
        end = d.get("end", "")
        start = d.get("start", "")
        filed = d.get("filed", "")
        val = d.get("val")
        if val is None:
            continue
        # For quarterly: prefer instantaneous (balance sheet) or period values
        if is_annual:
            label = end[:4]  # FY year
        else:
            label = end  # YYYY-MM-DD
        # Prefer longer period data (avoid re-filing overwrites) -- keep latest filed
        if label not in period_map or filed > period_map[label]["filed"]:
            period_map[label] = {"val": to_millions(val, unit_filter), "filed": filed, "end": end}
    return {k: v["val"] for k, v in period_map.items()}


# ── Segment / additional data from HTML filings ────────────────────────────────

def fetch_filing_html(cik: str, accession: str, primary_doc: str) -> str | None:
    base = accession.replace("-", "")
    url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{base}/{primary_doc}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            return r.text
    except Exception:
        pass
    return None


# Patterns that indicate a table is garbage (nav, legal boilerplate, exhibit lists)
JUNK_TABLE_PATTERNS = re.compile(
    r"table of contents|exhibit index|incorporated herein by reference|"
    r"certification of chief|pursuant to rule 13a|pursuant to section 906|"
    r"instance document|taxonomy extension|inline xbrl|"
    r"ernst.*young|deloitte|kpmg|pricewaterhousecoopers|/s/ |"
    r"trading arrangement|rule 10b5|shares to be sold|expiration date|"
    r"accounting standard|fasb issued|asc 842|adoption method|"
    r"bylaws|certificate of incorporation|indenture.*trustee",
    re.IGNORECASE
)

# A table is useful if it has numeric data and looks like financial/operational content
def is_useful_table(df: pd.DataFrame) -> bool:
    if df.shape[0] < 3 or df.shape[1] < 2:
        return False
    # Check for junk content in first few cells
    sample_text = " ".join(str(v) for v in df.iloc[:3].values.flatten() if v)
    if JUNK_TABLE_PATTERNS.search(sample_text):
        return False
    # Must have at least some numeric-looking content
    all_text = " ".join(str(v) for v in df.values.flatten() if v)
    has_numbers = bool(re.search(r"\d{2,}", all_text))
    # Must have meaningful column count (not single-column prose)
    if df.shape[1] < 2:
        return False
    # Reject if >80% of cells are empty
    non_empty = sum(1 for v in df.values.flatten() if str(v).strip() not in ("", "nan", "None"))
    if non_empty / max(df.size, 1) < 0.15:
        return False
    return has_numbers


def extract_tables_from_html(html: str, keywords: list[str]) -> list[pd.DataFrame]:
    """Find HTML tables near keyword matches, with strict quality filtering."""
    soup = BeautifulSoup(html, "html.parser")
    results = []
    seen_content = set()
    for kw in keywords:
        pattern = re.compile(r"\b" + re.escape(kw) + r"\b", re.IGNORECASE)
        matches = soup.find_all(string=pattern)
        for match in matches[:5]:
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
                            # Deduplicate by content fingerprint
                            fp = str(df.values.tolist())[:300]
                            if fp in seen_content:
                                break
                            seen_content.add(fp)
                            if is_useful_table(df):
                                results.append(df)
                    except Exception:
                        pass
                    break
                parent = parent.parent
    return results


# ── Excel builder ──────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", start_color="1F3864")
SUBHEADER_FILL = PatternFill("solid", start_color="2E75B6")
ALT_ROW_FILL = PatternFill("solid", start_color="EBF3FB")
WHITE_FILL = PatternFill("solid", start_color="FFFFFF")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
SUBHEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=9)
LABEL_FONT = Font(name="Arial", size=9)
BOLD_LABEL_FONT = Font(name="Arial", bold=True, size=9)
DATA_FONT = Font(name="Arial", size=9, color="00008B")
THIN_BORDER = Border(
    bottom=Side(style="thin", color="BDD7EE"),
    top=Side(style="thin", color="BDD7EE"),
)

def style_header_row(ws, row: int, ncols: int, text: str, fill=None, font=None):
    fill = fill or HEADER_FILL
    font = font or HEADER_FONT
    ws.cell(row=row, column=1).value = text
    ws.cell(row=row, column=1).font = font
    ws.cell(row=row, column=1).fill = fill
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, ncols + 2):
        ws.cell(row=row, column=c).fill = fill


def write_statement_sheet(wb: Workbook, sheet_name: str, statements: dict, periods: list[str], company_name: str, is_annual: bool):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B3"

    period_label = "Fiscal Year" if is_annual else "Quarter Ended"
    ncols = len(periods)

    # Title row
    ws.row_dimensions[1].height = 22
    ws.cell(row=1, column=1).value = f"{company_name} — {sheet_name}"
    ws.cell(row=1, column=1).font = Font(name="Arial", bold=True, size=11, color="1F3864")
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    note_col = ncols + 2
    ws.cell(row=1, column=note_col).value = "Units: Dollar values in $MM | EPS in $ per share | Share counts in MM shares | Source: SEC EDGAR XBRL"
    ws.cell(row=1, column=note_col).font = Font(name="Arial", size=8, color="808080", italic=True)

    # Period header row
    ws.row_dimensions[2].height = 18
    ws.cell(row=2, column=1).value = period_label
    ws.cell(row=2, column=1).font = HEADER_FONT
    ws.cell(row=2, column=1).fill = HEADER_FILL
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center")
    for i, p in enumerate(periods):
        c = i + 2
        ws.cell(row=2, column=c).value = p
        ws.cell(row=2, column=c).font = HEADER_FONT
        ws.cell(row=2, column=c).fill = HEADER_FILL
        ws.cell(row=2, column=c).alignment = Alignment(horizontal="center", vertical="center")

    current_row = 3
    for section, line_items in statements.items():
        # Section header
        ws.row_dimensions[current_row].height = 16
        style_header_row(ws, current_row, ncols, section, fill=SUBHEADER_FILL, font=SUBHEADER_FONT)
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        current_row += 1

        for i, (label, period_data) in enumerate(line_items.items()):
            ws.row_dimensions[current_row].height = 14
            is_total = any(x in label.lower() for x in ["total", "gross profit", "operating income",
                                                          "net income", "ebitda", "free cash flow",
                                                          "cash from operations", "cash from investing",
                                                          "cash from financing"])
            fill = ALT_ROW_FILL if i % 2 == 0 else WHITE_FILL
            ws.cell(row=current_row, column=1).value = label
            ws.cell(row=current_row, column=1).font = BOLD_LABEL_FONT if is_total else LABEL_FONT
            ws.cell(row=current_row, column=1).fill = fill
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="left", indent=2)

            for j, p in enumerate(periods):
                c = j + 2
                val = period_data.get(p)
                cell = ws.cell(row=current_row, column=c)
                cell.fill = fill
                cell.alignment = Alignment(horizontal="right")
                if val is not None:
                    cell.value = val
                    # EPS and shares get different format
                    if "EPS" in label or "Per Share" in label:
                        cell.number_format = '#,##0.00;(#,##0.00);"-"'
                    elif "Shares" in label:
                        cell.number_format = '#,##0.1;(#,##0.1);"-"'  # MM shares, 1 decimal
                    elif "%" in label or "Margin" in label:
                        cell.number_format = '0.0%;(0.0%);"-"'
                    else:
                        cell.number_format = '#,##0.0;(#,##0.0);"-"'
                    cell.font = BOLD_LABEL_FONT if is_total else DATA_FONT
                else:
                    cell.value = "—"
                    cell.font = Font(name="Arial", size=9, color="AAAAAA")
            current_row += 1
        current_row += 1  # blank spacer

    # Column widths
    ws.column_dimensions["A"].width = 36
    for i in range(len(periods)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 13


def write_raw_table_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, title: str):
    ws = wb.create_sheet(sheet_name[:31])
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1).value = title
    ws.cell(row=1, column=1).font = Font(name="Arial", bold=True, size=10, color="1F3864")
    for c_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(row=2, column=c_idx)
        cell.value = str(col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx + 3, column=c_idx)
            cell.value = val
            cell.font = Font(name="Arial", size=9)
            cell.fill = ALT_ROW_FILL if r_idx % 2 == 0 else WHITE_FILL
    for c_idx in range(1, len(df.columns) + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 22


# ── Core data pipeline ─────────────────────────────────────────────────────────

def build_statements(facts: dict, concept_map: dict, is_annual: bool, years: int = 15) -> tuple[dict, list[str]]:
    """Returns (statement_dict, sorted_periods)."""
    all_periods = set()
    built = {}
    for label, concepts in concept_map.items():
        if not concepts:
            built[label] = {}
            continue
        # Determine correct unit filter BEFORE fetching
        if "Shares" in label:
            unit_filter = "shares"
        elif "EPS" in label:
            unit_filter = "USD/shares"
        else:
            unit_filter = "USD"

        raw = extract_concept(facts, concepts, unit_filter)
        # Fallback: EPS sometimes filed as plain USD in older filings
        if not raw and unit_filter == "USD/shares":
            raw = extract_concept(facts, concepts, "USD")
        # Fallback: shares sometimes filed under USD (weighted avg concepts)
        if not raw and unit_filter == "shares":
            raw = extract_concept(facts, concepts, "USD")

        period_data = build_period_map(raw, is_annual, unit_filter)
        built[label] = period_data
        all_periods.update(period_data.keys())

    # Filter to last N years
    if is_annual:
        sorted_periods = sorted(all_periods, reverse=True)
        cutoff = str(pd.Timestamp.now().year - years)
        sorted_periods = [p for p in sorted_periods if p >= cutoff]
    else:
        sorted_periods = sorted(all_periods, reverse=True)
        cutoff = str((pd.Timestamp.now() - pd.DateOffset(years=years)).date())
        sorted_periods = [p for p in sorted_periods if p >= cutoff]

    # Calculate Gross Profit as fallback if not directly filed
    if "Gross Profit" in built:
        gp = built["Gross Profit"]
        if not any(v is not None for v in gp.values()):
            rev = built.get("Revenue", {})
            cogs = built.get("Cost of Revenue", {})
            built["Gross Profit"] = {
                p: round((rev.get(p) or 0) - (cogs.get(p) or 0), 3)
                for p in sorted_periods
                if rev.get(p) is not None and cogs.get(p) is not None
            }

    # Calculate EBITDA — add back D&A plus amortization of intangibles
    if "EBITDA (calc)" in built:
        op = built.get("Operating Income", {})
        da = built.get("Depreciation & Amortization", {})
        amort = built.get("Amortization of Intangibles", {})
        ebitda = {}
        for p in sorted_periods:
            op_val = op.get(p)
            if op_val is None:
                continue
            da_val = da.get(p) or 0
            amort_val = amort.get(p) or 0
            ebitda[p] = round(op_val + da_val + amort_val, 3)
        built["EBITDA (calc)"] = ebitda

    if "Free Cash Flow (calc)" in built:
        cfo = built.get("Cash from Operations", {})
        capex = built.get("Capex", {})
        built["Free Cash Flow (calc)"] = {
            p: round((cfo.get(p) or 0) - abs(capex.get(p) or 0), 3)
            for p in sorted_periods
            if cfo.get(p) is not None and capex.get(p) is not None
        }

    return built, sorted_periods


def group_statements(built: dict, is_income: bool = False, is_balance: bool = False, is_cf: bool = False) -> dict:
    per_share_keys = ["EPS Basic", "EPS Diluted", "Shares Basic", "Shares Diluted"]
    if is_income:
        return {
            "Income Statement ($MM)": {k: v for k, v in built.items() if k not in per_share_keys},
            "Per Share Data (EPS in $, Shares in MM)": {k: v for k, v in built.items() if k in per_share_keys},
        }
    if is_balance:
        assets = ["Cash & Equivalents", "Short-term Investments", "Accounts Receivable", "Inventory",
                  "Other Current Assets", "Total Current Assets", "PP&E Net", "Goodwill",
                  "Intangible Assets", "Long-term Investments", "Other Long-term Assets", "Total Assets"]
        liab = ["Accounts Payable", "Accrued Liabilities", "Short-term Debt", "Deferred Revenue Current",
                "Other Current Liabilities", "Total Current Liabilities", "Long-term Debt",
                "Deferred Revenue LT", "Other Long-term Liabilities", "Total Liabilities"]
        eq = ["Common Stock & APIC", "Retained Earnings", "Treasury Stock", "Accumulated OCI",
              "Total Stockholders Equity", "Total Liabilities & Equity"]
        return {
            "Assets ($MM)": {k: built[k] for k in assets if k in built},
            "Liabilities ($MM)": {k: built[k] for k in liab if k in built},
            "Equity ($MM)": {k: built[k] for k in eq if k in built},
        }
    if is_cf:
        ops = ["Net Income", "D&A (CF)", "Amortization of Intangibles (CF)", "Stock-Based Compensation",
               "Deferred Income Taxes", "Changes in Working Capital", "Other Operating Activities",
               "Cash from Operations"]
        inv = ["Capex", "Acquisitions", "Purchases of Investments", "Sales/Maturities of Investments",
               "Cash from Investing"]
        fin = ["Debt Issuance", "Debt Repayment", "Share Repurchases", "Dividends Paid",
               "Stock Issuance", "Cash from Financing"]
        summary = ["Net Change in Cash", "Free Cash Flow (calc)"]
        return {
            "Operating Activities ($MM)": {k: built[k] for k in ops if k in built},
            "Investing Activities ($MM)": {k: built[k] for k in inv if k in built},
            "Financing Activities ($MM)": {k: built[k] for k in fin if k in built},
            "Summary ($MM)": {k: built[k] for k in summary if k in built},
        }
    return {"Data": built}


def fetch_segment_data(cik: str, filings: list[dict], max_filings: int = 8) -> list[tuple[str, pd.DataFrame]]:
    results = []
    # More targeted keywords — focus on financial/operational disclosures
    keywords = [
        "segment revenue", "segment information", "revenue by segment",
        "geographic", "revenue by region", "revenue by geography",
        "disaggregated revenue", "revenue disaggregation",
        "customer concentration", "significant customer", "major customer",
        "royalt", "product sales", "collaborative",
        "subscribers", "active users", "monthly active", "annual recurring",
        "units sold", "volume", "same.store", "comparable store",
        "backlog", "bookings", "net revenue retention",
        "key performance", "operating metric",
    ]
    seen_tables = set()
    for filing in filings[:max_filings]:
        html = fetch_filing_html(cik, filing["accession"], filing["primary_doc"])
        if not html:
            continue
        tables = extract_tables_from_html(html, keywords)
        for df in tables:
            key = str(df.values.tolist())[:300]
            if key in seen_tables:
                continue
            seen_tables.add(key)
            df = df.dropna(how="all").fillna("")
            results.append((filing["date"], df))
        time.sleep(0.3)
    return results


# ── Streamlit UI ───────────────────────────────────────────────────────────────

def main():
    st.title("📊 SEC Financial Data Extractor")
    st.markdown("Extract 15+ years of financial statements, segments, and KPIs from SEC EDGAR filings.")

    with st.sidebar:
        st.header("Settings")
        ticker = st.text_input("Stock Ticker", placeholder="e.g. AAPL, MSFT, NVDA").strip().upper()
        years = st.slider("Years of History", min_value=5, max_value=20, value=15)
        include_segments = st.checkbox("Extract Segment / KPI Tables (slower)", value=True)
        max_filings_for_segments = st.slider("# of Filings to Scan for Segments", 4, 20, 8,
                                              help="More filings = more complete segment history but slower")
        run_btn = st.button("🚀 Extract Data", type="primary", use_container_width=True)

    if not run_btn or not ticker:
        st.info("Enter a ticker in the sidebar and click **Extract Data** to begin.")
        st.markdown("""
**What this tool extracts:**
- ✅ Income Statement (quarterly + annual) — 15 years
- ✅ Balance Sheet (quarterly + annual)
- ✅ Cash Flow Statement (quarterly + annual)
- ✅ Per Share Data (EPS, share counts)
- ✅ Segment Revenue / Geographic Revenue tables (best-effort from filing HTML)
- ✅ KPI tables, customer concentration disclosures (best-effort)
- ✅ All values in $MM where applicable
        """)
        return

    progress = st.progress(0)
    status = st.empty()

    try:
        status.text("🔍 Looking up CIK for ticker...")
        cik = get_cik(ticker)
        if not cik:
            st.error(f"Could not find CIK for ticker '{ticker}'. Check the ticker and try again.")
            return
        progress.progress(5)

        company_name = get_company_name(cik)
        st.success(f"Found: **{company_name}** (CIK: {int(cik)})")

        status.text("📥 Downloading XBRL company facts (this is the big one)...")
        facts = get_xbrl_facts(cik)
        progress.progress(20)

        status.text("🔢 Processing Income Statement...")
        is_annual_built, annual_periods = build_statements(facts, INCOME_STATEMENT_CONCEPTS, True, years)
        is_qtr_built, qtr_is_periods = build_statements(facts, INCOME_STATEMENT_CONCEPTS, False, years)
        progress.progress(35)

        status.text("🔢 Processing Balance Sheet...")
        bs_annual_built, bs_annual_periods = build_statements(facts, BALANCE_SHEET_CONCEPTS, True, years)
        bs_qtr_built, bs_qtr_periods = build_statements(facts, BALANCE_SHEET_CONCEPTS, False, years)
        progress.progress(50)

        status.text("🔢 Processing Cash Flow Statement...")
        cf_annual_built, cf_annual_periods = build_statements(facts, CASH_FLOW_CONCEPTS, True, years)
        cf_qtr_built, cf_qtr_periods = build_statements(facts, CASH_FLOW_CONCEPTS, False, years)
        progress.progress(65)

        segment_tables = []
        if include_segments:
            status.text("🔍 Fetching 10-K and 10-Q filings for segment/KPI data...")
            annual_filings = get_filings_index(cik, "10-K")
            qtr_filings = get_filings_index(cik, "10-Q")
            all_filings = sorted(annual_filings + qtr_filings, key=lambda x: x["date"], reverse=True)
            segment_tables = fetch_segment_data(cik, all_filings, max_filings_for_segments)
        progress.progress(80)

        status.text("📝 Building Excel workbook...")
        wb = Workbook()
        wb.remove(wb.active)

        # Annual sheets
        write_statement_sheet(wb, "Annual — Income Statement",
                               group_statements(is_annual_built, is_income=True),
                               list(reversed(annual_periods)), company_name, True)
        write_statement_sheet(wb, "Annual — Balance Sheet",
                               group_statements(bs_annual_built, is_balance=True),
                               list(reversed(bs_annual_periods)), company_name, True)
        write_statement_sheet(wb, "Annual — Cash Flow",
                               group_statements(cf_annual_built, is_cf=True),
                               list(reversed(cf_annual_periods)), company_name, True)

        # Quarterly sheets
        write_statement_sheet(wb, "Quarterly — Income Statement",
                               group_statements(is_qtr_built, is_income=True),
                               qtr_is_periods[:60], company_name, False)
        write_statement_sheet(wb, "Quarterly — Balance Sheet",
                               group_statements(bs_qtr_built, is_balance=True),
                               bs_qtr_periods[:60], company_name, False)
        write_statement_sheet(wb, "Quarterly — Cash Flow",
                               group_statements(cf_qtr_built, is_cf=True),
                               cf_qtr_periods[:60], company_name, False)

        # Segment / KPI sheets
        if segment_tables:
            for idx, (filing_date, df) in enumerate(segment_tables[:25]):
                sheet_title = f"Seg-KPI — {filing_date}"
                write_raw_table_sheet(wb, sheet_title, df, f"Extracted Table — Filing Date: {filing_date}")

        progress.progress(95)

        # Save to buffer
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
        st.subheader("Preview — Annual Income Statement")
        preview_data = []
        for section, items in group_statements(is_annual_built, is_income=True).items():
            for label, period_data in items.items():
                row = {"Line Item": label, "Section": section}
                for p in list(reversed(annual_periods))[:10]:
                    row[p] = period_data.get(p, "—")
                preview_data.append(row)
        if preview_data:
            st.dataframe(pd.DataFrame(preview_data).set_index("Line Item"), use_container_width=True)

        if segment_tables:
            st.subheader(f"Segment / KPI Tables Found: {len(segment_tables)}")
            st.caption("These are written to individual tabs in the Excel file (Seg-KPI — YYYY-MM-DD)")
            with st.expander("Preview first segment/KPI table"):
                st.dataframe(segment_tables[0][1], use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)


if __name__ == "__main__":
    main()
