# P&L Flux Analyzer

A comprehensive Python utility for analyzing Profit & Loss (P&L) financial data with advanced flux analysis, department/class breakdowns, and AI-powered insights. This tool reads Excel or CSV files containing financial transactions and provides detailed month-over-month and quarter-over-quarter analysis.

## Features

- **Multi-format Support**: Reads both Excel (.xlsx, .xls, .xlsm) and CSV files
- **Flexible Column Mapping**: Automatically detects and maps common column names (Period, Amount, Memo, Department, Class)
- **Comprehensive Analysis**: 
  - Month-over-Month (MoM) flux and percentage changes
  - Quarter-over-Quarter (QoQ) analysis with configurable fiscal year end
  - Department and Class-level breakdowns
- **Interactive Interface**: User-friendly prompts for configuration
- **AI-Powered Insights**: Optional OpenAI integration for intelligent analysis
- **Rich Output**: Detailed Excel workbooks with multiple sheets and analysis

## Installation

```bash
# Install dependencies (recommended inside virtualenv)
python3 -m pip install -r requirements.txt
```

## Usage

The tool runs interactively and will prompt you for configuration:

```bash
python3 PL_Flux.py
```

### Interactive Prompts

1. **Input File**: Enter the Excel/CSV filename to analyze
2. **AI Analysis**: Choose whether to run AI analysis (y/N)
3. **Department Analysis**: Enable department-level breakdowns (y/N)
4. **Class Analysis**: Enable class-level breakdowns (y/N)
5. **Fiscal Year End**: Set fiscal year end month (1-12, default: 12)
6. **Analysis Mode**: Choose MoM or QoQ analysis (if AI enabled)

### Required Columns

- **Period**: Date column (supports various date formats)
- **Amount**: Financial amount column (handles currency symbols, commas, parentheses)
- **Memo/Description**: Transaction description column (required for AI analysis)

### Optional Columns

- **Department**: For department-level analysis
- **Class**: For class-level analysis

## Output Files

### Excel Workbook (`summary_flux.xlsx`)

- **Summary Sheet**: Monthly totals with MoM flux and percentage changes
- **Quarterly Sheet**: Fiscal quarter summaries with QoQ analysis
- **Monthly Sheets**: Individual transaction details for each month (YYYY-MM format)
- **Department/Class Sheets**: Breakdowns by department and class (if enabled)
- **AI_Analysis Sheet**: AI-generated insights (if enabled)

### Text Analysis (`openai_analysis.txt`)

Detailed AI analysis of financial fluctuations and trends.

## AI Analysis Features

The optional AI analysis provides:

- **Intelligent Insights**: Explains month-over-month or quarter-over-quarter fluctuations
- **Context-Aware Analysis**: Considers department, class, and memo information
- **Trend Identification**: Identifies patterns and potential causes of changes
- **Detailed Breakdowns**: Top increases/decreases by various dimensions

### AI Configuration

The tool automatically manages OpenAI API configuration:

1. Creates/updates a `.env` file with your API key
2. Supports custom base URLs for different OpenAI-compatible services
3. Uses configurable models (default: gpt-5-2025-08-07)

## Advanced Features

### Fiscal Year Configuration

- Configurable fiscal year end month (1-12)
- Automatic fiscal quarter calculation
- Quarterly analysis sheets

### Data Processing

- **Smart Column Detection**: Handles case variations and whitespace
- **Currency Formatting**: Processes various currency formats ($, commas, parentheses)
- **Date Parsing**: Supports multiple date formats and period strings
- **Error Handling**: Graceful handling of malformed data

### Analysis Modes

- **Month-over-Month (MoM)**: Detailed monthly trend analysis
- **Quarter-over-Quarter (QoQ)**: Fiscal quarter analysis
- **Multi-dimensional**: Department and class breakdowns
- **Pairwise Analysis**: Detailed comparison between adjacent periods

## Example Workflow

1. Prepare your financial data with Period, Amount, and Memo columns
2. Run `python3 PL_Flux.py`
3. Follow the interactive prompts
4. Review the generated Excel workbook and AI analysis
5. Use insights for financial planning and analysis

## Contact

For issues or questions, contact Ray Sang.
