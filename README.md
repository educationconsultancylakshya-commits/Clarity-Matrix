# AI Wire & Cable Technical Extraction Platform

A ready-to-use Streamlit web platform for extracting and comparing technical specifications from wire/cable documents, product pages, and a target brand ecosystem.

## What it does

- Upload up to **100** technical files at once.
- Paste up to **50** product/spec URLs.
- Extract technical attributes and values, including voltage rating, jacket material, compliance standards, conductor material, insulation, shielding, AWG size, temperature rating, fire rating, and more.
- Use optional AI extraction and AI validation against common wire/cable industry terminology.
- Show dashboards for:
  - Most common attributes
  - Attribute value breakdown
  - Advanced attribute search with a **30% minimum appearance threshold**, plus a troubleshooting **Show all** option
- Run a bottom-page brand comparison module that extracts common attributes and values from the listed brands, then exports Excel reports.
- Use the separate **Review** tab to see all extracted data in a grouped review table matching the requested format: Values Correct/Missed/False and Attributes Spec Sheet NA/AI Returned False NA/False Positives/Omissions.
- New: brand analysis has a **Product focus / cable family** field and an **Auto-fill if AI returns empty** safety option so dashboards do not come back blank if the model/web search returns no usable JSON.

## Files in this repository

```text
app.py                         Main Streamlit application
requirements.txt               Python dependencies
.streamlit/config.toml          Streamlit UI and upload settings
sample_secrets.toml             Example secrets; do not upload a real key
sample_brand_urls.csv           Optional fallback template for brand URL analysis
.github/workflows/smoke-test.yml GitHub Actions compile test
README.md                       This guide
```

## No-code deployment using GitHub + Streamlit Cloud

### Step 1: Create a GitHub repository

1. Go to GitHub.
2. Click **New repository**.
3. Name it something like `ai-wire-cable-platform`.
4. Keep it private if you want to protect your code.
5. Click **Create repository**.

### Step 2: Upload the project files

1. Download and unzip the project package.
2. In your GitHub repository, click **Add file > Upload files**.
3. Drag all files and folders from the unzipped project into GitHub.
4. Click **Commit changes**.

### Step 3: Deploy on Streamlit Community Cloud

1. Go to Streamlit Community Cloud.
2. Click **Create app**.
3. Choose **Deploy a public app from GitHub** or connect your GitHub account.
4. Select your repository.
5. Set the main file path to:

```text
app.py
```

6. Open **Advanced settings** and paste your secrets:

```toml
OPENAI_API_KEY = "paste-your-openai-api-key-here"
OPENAI_MODEL = "gpt-4o-mini"
```

7. Click **Deploy**.

You will get a web URL ending in `.streamlit.app`. Open that URL and start using the platform.

## Important security note

Do **not** paste your real OpenAI API key into any GitHub file. Use Streamlit Cloud secrets only.

## Local testing, optional

If someone technical wants to test it locally:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL shown in the terminal.

## How to use the app

### Document & URL Analysis

1. Open the app.
2. Upload PDFs, DOCX files, CSV/XLSX files, or text files.
3. Paste product/spec URLs, one per line.
4. Click **Analyze uploaded documents and URLs**.
5. Review the dashboards.
6. Download the Excel report.
7. Open the **Review** tab to see the grouped review table.

### Review Tab

The app now includes a separate **Review** tab. After you run document/URL analysis or brand analysis, open this tab and choose the data source. The table is displayed in the requested format with grouped headers:

- **Values:** Correct, Missed, False
- **Attributes:** Spec Sheet NA, AI Returned False NA, False Positives, Omissions

Each category shows a count and percentage. You can also download the review table as an Excel report.

### Brand Comparison / Top Brand Attribute Extraction

This section now appears **at the bottom of the main platform page**, underneath the document and URL analysis area.

There are two options:

**Option A: Automatic AI brand analysis with blank-result protection**

- Turn on AI extraction in the sidebar.
- Add your OpenAI API key in Streamlit Cloud secrets.
- Scroll to the bottom section named **Brand Comparison / Top Brand Attribute Extraction**.
- Enter the product family, for example `Fire alarm cable`.
- Keep **Auto-fill if AI returns empty** turned ON. This prevents empty brand dashboards when the selected AI model cannot browse the web or returns no valid JSON.
- Click **Get common attributes and values from the Top 50 Brands**.
- Check the **Raw Brand Records** table. Rows marked `built_in_industry_fallback` came from the built-in fallback, while rows marked `ai_web_brand_analysis` came from the AI response.

**Option B: Official brand URL CSV analysis**

- Fill in `sample_brand_urls.csv` with `brand,url` rows using official product/spec pages.
- Open the **Alternative: analyze official brand URLs from CSV** expander in the bottom brand section.
- Upload the CSV and click **Analyze provided brand URLs**.
- Use this option when you need source-grounded brand results instead of fallback-assisted results.

## Notes

- Dashboard 3 follows the required 30% threshold by default. If no values meet the threshold, the app now shows a helpful message and the closest extracted values so it does not look broken.
- The brand module now includes a **Brand Comparison Matrix** where each row is an attribute/value pair and each brand column shows whether that value was found.
- The separate **Review** tab presents review results in a grouped table layout like the provided screenshot.
- If brand analysis previously showed `Brands analyzed: 48` but `Unique attributes: 0`, update to this version and keep **Auto-fill if AI returns empty** turned ON.
- The target brand list is loaded directly into `app.py`.
- The uploaded requirement calls this the Top 50 Brands list; the provided named list contains 48 entries. The app includes an input box where you can add extra brands without editing code.
- AI output should still be reviewed by a domain expert before using it for compliance or purchasing decisions.
