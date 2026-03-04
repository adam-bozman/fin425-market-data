# Market Data Downloader — FIN 425

A Streamlit app that lets students pull historical price data from Yahoo Finance
and download it as a formatted Excel workbook with three sheets:
S&P 500 (^GSPC), Risk-Free Rate (^IRX), and a firm of their choosing.

## What students can configure
- Firm ticker (e.g. LMT, AAPL, BA)
- Date range
- Frequency (Daily / Weekly / Monthly)

## Deploy to Streamlit Community Cloud (free, ~5 minutes)

### Step 1 — Push to GitHub
1. Create a **new public GitHub repository** (e.g. `fin425-market-data`)
2. Upload both files into the repo root:
   - `app.py`
   - `requirements.txt`

### Step 2 — Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
2. Click **"New app"**
3. Select your repository, branch (`main`), and set **Main file path** to `app.py`
4. Click **Deploy** — it will install dependencies automatically

### Step 3 — Share the URL
Streamlit gives you a public URL like:
`https://your-username-fin425-market-data-app-xxxx.streamlit.app`

Share that link with students. No login required on their end.

## Run locally (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- Yahoo Finance data is free but occasionally rate-limits heavy usage
- `^IRX` is the 13-week T-bill annualized yield in percent — not a price
- The Excel workbook includes a Cover sheet summarizing parameters,
  plus one data tab per ticker (Date, Open, High, Low, Close, Adj. Close, Volume)
