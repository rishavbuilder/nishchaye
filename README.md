# Nikshay Data Entry Bot

Automated data entry bot for Nikshay (TB Control Programme) - fills enrollment forms from Excel data.

## Features

- Auto-login to Nikshay portal
- Bulk enrollment from Excel (.xlsx) files
- Fills: Name, Father/Husband Name, Age, Gender, Mobile, Symptoms, Key Population, Village
- Auto-selects Health Facility (Ab HWC Gulni) and Village
- Handles Key Population checkboxes (unchecks "Not Applicable", checks correct option)
- Smart symptom matching (aliases like "night sweat" → "Night sweat")
- Failed records saved to `failed_records.xlsx`
- Successful records deleted from original Excel
- Retry logic for failed records (3 attempts)
- Default fallbacks: Village → "Bermi", Key Population → "Other"

## Setup

1. Install Python 3.8+
2. Install dependencies:
   ```bash
   pip install python-dotenv openpyxl selenium webdriver-manager
   ```
3. Create `config.env` from example:
   ```bash
   cp config.env.example config.env
   ```
4. Edit `config.env` with your Nikshay credentials

## Excel Format

Your Excel file should have these columns:

| Name | Mobile | Father/Husband Name | Age | Gender | Marital Status | Symptoms | Key Population | Village |
|------|--------|-------------------|-----|--------|---------------|----------|---------------|---------|

## Usage

```bash
python nikshay_data_entry.py
```

## How it Works

1. Logs in to Nikshay portal
2. Reads records from Excel
3. For each record:
   - Fills enrollment form
   - Selects appropriate options
   - Submits the form
4. After completion:
   - Successful records removed from Excel
   - Failed records saved to `failed_records.xlsx`
   - Summary printed in terminal

## Config Options

| Option | Default | Description |
|--------|---------|-------------|
| NIKSHAY_USERNAME | - | Nikshay login username |
| NIKSHAY_PASSWORD | - | Nikshay login password |
| NIKSHAY_HEADLESS | false | Run browser in headless mode |
| NIKSHAY_START_ROW | 1 | Start from which row in Excel |
| NIKSHAY_ADDRESS | Bari Gulni | Default address |
| NIKSHAY_PINCODE | 805124 | Default pincode |
| NIKSHAY_HEALTH_FACILITY | Ab HWC Gulni | Health facility name |

## Notes

- Run `start.sh` for quick start
- Check `debug_*.png` screenshots if any issues occur
- Failed records are saved with error reasons for debugging
