# Summarize Amount by Period

This utility reads an Excel file with a `Period` column and an `Amount` column,
aggregates Amount by month, computes month-over-month flux and percent change,
and writes the results to an Excel workbook with a `Summary` sheet and one sheet per month.

Usage examples:

```bash
# install dependencies (recommended inside virtualenv)
python3 -m pip install -r requirements.txt

# run the summarizer (defaults assume `68100 details.xlsx` exists)
python3 summarize_amount_by_period.py

# or specify a file and output name
python3 summarize_amount_by_period.py "68100 details.xlsx" --output-excel summary_flux.xlsx
```

Files produced:

 Optional OpenAI analysis
 - The script includes an opt-in OpenAI analysis that summarizes month-over-month fluctuations using transaction `Amount` and `memo` columns.
 - To enable it, set your API key in the `OPENAI_API_KEY` environment variable (do not hard-code keys into files):

 ```bash
 export OPENAI_API_KEY="sk-..."
 # Default model is gpt-4.1-2025-04-14; override with --openai-model or OPENAI_MODEL
 python3 summarize_amount_by_period.py "68100 details.xlsx" --output-excel summary_flux.xlsx --openai-analyze
 # Or explicitly select model
 python3 summarize_amount_by_period.py "68100 details.xlsx" --openai-analyze --openai-model gpt-4.1-2025-04-14
 ```

 The analysis will be written to `openai_analysis.txt` and appended as an `AI_Analysis` sheet in the Excel workbook (when possible).
