# SEC Financial Data Extractor

Extracts 15+ years of quarterly and annual financial data from SEC EDGAR for any US-listed stock.

## What It Extracts

- **Income Statement** — Revenue, Gross Profit, EBITDA, Operating Income, Net Income, EPS, Share Counts
- **Balance Sheet** — Full assets, liabilities, equity breakdown
- **Cash Flow Statement** — Operating, Investing, Financing + Free Cash Flow
- **Segment / Geographic / KPI Tables** — Best-effort extraction from filing HTML
- **Output** — Excel file with separate Quarterly and Annual tabs per statement, all values in $MM

---

## Deploy on Render (Free) — Step by Step

### Step 1: Create a GitHub repository

1. Go to [github.com](https://github.com) and sign in (create a free account if needed)
2. Click **New repository** (green button top right)
3. Name it `sec-financial-extractor`, set to **Public**, click **Create repository**
4. On your computer, install [Git](https://git-scm.com/downloads) if you don't have it
5. Open a terminal/command prompt and run:

```bash
git clone https://github.com/YOUR_USERNAME/sec-financial-extractor.git
cd sec-financial-extractor
```

6. Copy the files from this project (`app.py`, `requirements.txt`, `render.yaml`) into that folder
7. Run:

```bash
git add .
git commit -m "Initial commit"
git push origin main
```

### Step 2: Deploy on Render

1. Go to [render.com](https://render.com) and sign up for a free account
2. Click **New +** → **Web Service**
3. Connect your GitHub account when prompted
4. Select your `sec-financial-extractor` repository
5. Render will auto-detect the `render.yaml` — if not, set manually:
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`
6. Choose the **Free** instance type
7. Click **Create Web Service**

Render will build and deploy in ~3-5 minutes. You'll get a URL like `https://sec-financial-extractor.onrender.com`.

> **Note**: Free Render instances spin down after 15 minutes of inactivity. The first load after inactivity takes ~30 seconds to wake up. This is normal.

---

## How to Use

1. Open your Render URL
2. Enter a stock ticker (e.g. `AAPL`, `MSFT`, `NVDA`, `JPM`)
3. Set years of history (up to 20)
4. Toggle segment/KPI extraction on/off
5. Click **Extract Data**
6. Wait 30-90 seconds while it pulls from SEC EDGAR
7. Click **Download Excel File**

---

## Notes on Data Quality

- **Structured financials (IS/BS/CF)**: Very reliable — sourced from XBRL structured data filed directly with the SEC
- **Segment / KPI / Geographic tables**: Best-effort HTML parsing. Results vary by company and filing format. You may get some noise (unrelated tables) — scan through the Seg-KPI tabs and delete what's not relevant
- **Units**: All dollar values converted to $MM. EPS in dollars. Share counts in MM shares
- **Coverage**: Works for all US-listed companies that file 10-K/10-Q with the SEC. Does not cover foreign private issuers that file only 20-F/6-K (partial support)

---

## Running Locally (Optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```
