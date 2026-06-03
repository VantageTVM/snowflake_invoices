# snowflake_invoices

Tools to convert Snowflake usage statement PDFs into Excel spreadsheets.

## Files

- [app.py](app.py) - Streamlit web app for uploading a PDF and downloading the resulting Excel file. Supports appending to an existing workbook and auto-derives the output filename from the input.
- [snowflake_usage.py](snowflake_usage.py) - CLI script for the same conversion. Run directly when a GUI is not needed.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

### Streamlit app

```bash
streamlit run app.py
```

Upload a Snowflake usage PDF, optionally upload an existing `.xlsx` to append to, then download the result.

### CLI

```bash
python snowflake_usage.py input.pdf [output.xlsx]
```

Defaults to `snowflake_usage.xlsx` if no output path is given.

## Output format

Both tools produce an Excel sheet (`Monthly Usage`) with four columns:

| USAGE MONTH | USAGE CATEGORY | UNITS CONSUMED | TOTAL USAGE (USD) |
|-------------|---------------|----------------|-------------------|

Months are formatted as `MMM-YY`. Category total rows are bolded. Numbers use accounting-style formatting with parentheses for negatives.
