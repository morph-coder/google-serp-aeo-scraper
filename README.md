# Google Sheets + Apify SERP Scraper

Connect your spreadsheet to the [Google SERP & AEO Scraper](https://apify.com/morph_coder/google-serp-aeo-scraper) Actor — run keyword scans and import flat results without writing code.

---

## 1. Before you start

1. Open the Actor in Apify and **activate it** in your account:  
   https://apify.com/morph_coder/google-serp-aeo-scraper  
2. Copy your **Apify API token**:  
   https://console.apify.com/account/integrations  

You only need the Apify token in the spreadsheet. LLM API keys (ChatGPT, Gemini, etc.) stay in the Actor settings on Apify — not in Google Sheets.

---

## 2. Add the script to your spreadsheet

### Step 1 — New spreadsheet

Create a blank Google Spreadsheet in your Google Drive.

<!-- screenshot: blank spreadsheet -->

### Step 2 — Open Apps Script

Menu: **Extensions → Apps Script**

<!-- screenshot: Extensions → Apps Script -->

### Step 3 — Paste `Code.gs`

1. Delete the default `function myFunction() { ... }` code.
2. Open the script file from GitHub (raw):  
   https://raw.githubusercontent.com/morph_coder/google-serp-aeo-scraper/main/integrations/google-sheets/Code.gs  
3. Select all → Copy → Paste into Apps Script.
4. Save (Ctrl+S / Cmd+S).

<!-- screenshot: Code.gs pasted in Apps Script editor -->

### Step 4 — Reload the spreadsheet

Close the tab and open the spreadsheet again (or refresh the page).

On first open the script creates sheets automatically: **Settings**, **Keywords**, **Results**, **LLM Summary**, **Run Log**, etc.

Menu **SERP Tools** should appear in the toolbar.

<!-- screenshot: SERP Tools menu visible -->

---

## 3. Configure and run

### Apify token (one time)

**SERP Tools → Configure Apify token** → paste your Apify API token → OK.

The token is saved in **Apps Script → Script properties** (not in any cell). It is sent only to `api.apify.com`.

<!-- screenshot: Configure Apify token dialog -->

### Keywords

Open the **Keywords** sheet. Enter one search keyword per row in column **A** (from row 2).

Example: `nike`, `adidas`

### Settings (optional)

Open **Settings** to change country, how many results per keyword, LLM on/off, etc. Defaults work for a first test.

### Run

**SERP Tools → Run SERP scan**

- The Actor starts on Apify.
- Results import in about **1 minute** (automatic background check).
- First time: Google may ask to **authorize the script** — click Allow.

When finished, check **Results** and **Run Log** (`costUsd` = what Apify charged for that run).

<!-- screenshot: Results sheet with data -->

**Manual import:** **SERP Tools → Fetch last run results** if you want to pull results without starting a new run.

---

## 4. About keys and security

| What | Where it lives | Used for |
|------|----------------|----------|
| **Apify API token** | Script Properties (via menu) | Start Actor runs, download results |
| **LLM keys** (OpenAI, Gemini, …) | Apify Actor secrets | LLM add-ons — configured by the Actor owner on Apify |

**Recommendations:**

- Do not put your Apify token in spreadsheet cells or share the sheet publicly with the token configured.
- Before first use, **paste `Code.gs` into your AI assistant** (ChatGPT, Claude, Cursor, etc.) and ask it to review what the script does — especially what URLs it calls and where it stores secrets. You should be comfortable with the code before granting Google authorization.
- The script only talks to `api.apify.com` using your token. It does not send data to other third-party APIs.

---

## 5. Troubleshooting

| Problem | What to try |
|---------|-------------|
| No **SERP Tools** menu | Reload spreadsheet; check Apps Script saved without errors |
| Run fails immediately | Activate the Actor in Apify Console; check token |
| No results after 1 min | **SERP Tools → Fetch last run results**; check Run Log status |
| Authorization popup | Normal on first run — review permissions, then Allow |

---

## Links

- **Apify Actor:** https://apify.com/morph_coder/google-serp-aeo-scraper  
- **Code.gs (raw):** https://raw.githubusercontent.com/morph_coder/google-serp-aeo-scraper/main/integrations/google-sheets/Code.gs  
