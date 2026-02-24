# Gold & Silver Price Tracker for Google Sheets

A Google Apps Script automation that fetches and logs gold and silver prices from Nepali jewelry association websites directly into a Google Sheet. Features automatic fallback between data sources, price change visualization, and scheduled daily updates.

## Overview

This project automates the collection of precious metal prices from multiple Nepali jewelry association websites:
- **FENEGOSIDA** (Federation of Nepal Gold and Silver Dealers' Association)
- **FNGSGJA** (Federation of Nepal Gold Silver and Jewellery Associations)

The script scrapes current rates for gold and silver (per tola and per 10 grams), logs them to a Google Sheet, calculates the gold-to-silver ratio, and visually indicates price changes with color coding.

## Features

- **Multi-Source Data Fetching**: Automatically tries multiple URLs with intelligent fallback
- **Resilient Architecture**: Retry logic and caching ensure data collection even when sources are temporarily unavailable
- **Price Change Visualization**: 
  - Green text = Price increased
  - Red text = Price decreased
  - Grey text = Price unchanged
- **Automatic Calculations**: Computes gold-to-silver ratio automatically
- **Scheduled Execution**: Daily automatic updates via time-based triggers (default: 11:15 AM)
- **Customizable Configuration**: Easy-to-edit settings for headers, formats, colors, and schedule

## Installation

### 1. Create a Google Sheet
1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name your spreadsheet (e.g., "Gold Silver Price Tracker")

### 2. Open Apps Script
1. Click **Extensions** → **Apps Script**
2. Delete any code in the default `Code.gs` file
3. Copy the entire contents of `Import Gold and Silver Price.gs` into the script editor
4. Click **Save** (💾) and name your project (e.g., "Gold Silver Price Tracker")

### 3. Authorize the Script
1. Click **Run** → **Run function** → Select `update_gold_silver_prices`
2. Google will request authorization:
   - Click through the authorization prompts
   - Review and allow the requested permissions (spreadsheet access, external URL fetching, triggers)
3. The script will run and create the "Rates" sheet with your first data entry

## Configuration

All customizable settings are in the `CONFIG` object at the top of the script:

```javascript
var CONFIG = {
  // Data source URLs (script tries them in order)
  source_urls: [
    "https://www.fngsgja.org.np/",
    "https://www.fenegosida.org/",
    "https://fenegosida.org/",
    "https://www.fenegosida.com/"
  ],

  // Sheet settings
  sheet_name: "Rates",           // Name of the sheet tab
  header_row_number: 1,          // Row containing headers
  insert_at_row_number: 2,       // Where new data is inserted

  // Schedule settings
  trigger: {
    hour: 11,      // 24-hour format
    minute: 15     // Near this minute
  },

  // Column headers (customize as needed)
  headers: [
    "Run Datetime",
    "Nepali Date",
    "Gold per tola",
    "Silver per tola",
    "Gold per 10 grm",
    "Silver per 10 grm",
    "Gold to Silver Ratio"
  ],

  // Visual settings
  change_colors: {
    up: "#1a7f37",    // Green for price increase
    down: "#d1242f",  // Red for price decrease
    same: "#6e7781"   // Grey for no change
  }
};
```

### Customization Options

| Setting | Description | Default |
|---------|-------------|---------|
| `source_urls` | Array of URLs to fetch data from | 4 Nepali jewelry sites |
| `sheet_name` | Name of the sheet tab | "Rates" |
| `trigger.hour` | Hour for daily execution (24h) | 11 |
| `trigger.minute` | Minute for daily execution | 15 |
| `max_retries_per_url` | Retry attempts per URL | 3 |
| `cache_minutes` | How long to cache successful fetches | 120 |

## How It Works

### Data Flow

1. **Trigger Check**: Ensures a daily time-based trigger exists (creates one if missing)
2. **HTML Fetching**: Attempts to fetch data from configured URLs in order
3. **Caching**: Stores successful fetches in CacheService for fallback
4. **Parsing**: Uses regex-based parsers to extract prices from HTML
5. **Data Insertion**: Inserts new row at the top (below headers)
6. **Formatting**: Applies number formats and color-codes price changes
7. **Logging**: Records success/failure in the Apps Script execution log

### HTML Parsers

The script includes two parsers for different website formats:

1. **FENEGOSIDA Parser** (`parse_fenegosida_rate_wrap_`): Parses the `rate-content-wrap` div structure
2. **FNGSGJA Parser** (`parse_fngsgja_home_`): Parses text-based layouts

The script automatically tries parsers in order until one succeeds.

### Error Handling

- **Retry Logic**: Each URL gets up to 3 attempts with 1.2s delays
- **Cache Fallback**: If all URLs fail, uses last cached successful response
- **Debug Logging**: Logs HTML snippets when parsing fails for troubleshooting

## Usage

### Manual Execution
Run `update_gold_silver_prices()` from the Apps Script editor for immediate data fetch.

### Automatic Updates
Once configured, the script automatically runs daily at the scheduled time. The trigger is created automatically on first run.

### Viewing Data
Open your Google Sheet and navigate to the "Rates" tab to see:
- Timestamp of each data collection
- Nepali calendar date (from the source website)
- Gold and silver prices (per tola and per 10 grams)
- Gold-to-silver ratio
- Color-coded price changes

## Troubleshooting

### Script fails to fetch data
1. Check **View** → **Executions** in the Apps Script editor for error details
2. Verify the source websites are accessible
3. Check if the website HTML structure has changed

### Price parsing fails
1. The script logs debug snippets when parsing fails
2. Check **View** → **Executions** → Click on failed run → View logs
3. Look for HTML structure changes on the source websites

### Trigger not working
1. Go to **Triggers** (⏰ icon) in the Apps Script editor
2. Verify a time-based trigger exists for `update_gold_silver_prices`
3. Check if the trigger was disabled due to errors

### Authorization errors
1. Go to **Apps Script** → **Project Settings**
2. Click the Google Cloud Platform project link
3. Ensure the project has proper OAuth consent screen configuration

## Data Format

The script creates a sheet with the following columns:

| Column | Header | Format | Description |
|--------|--------|--------|-------------|
| A | Run Datetime | Date/Time | When the data was collected |
| B | Nepali Date | Text | Date as shown on source website |
| C | Gold per tola | Number | Gold price per tola (NPR) |
| D | Silver per tola | Number | Silver price per tola (NPR) |
| E | Gold per 10 grm | Number | Gold price per 10 grams (NPR) |
| F | Silver per 10 grm | Number | Silver price per 10 grams (NPR) |
| G | Gold to Silver Ratio | Number | Ratio of gold to silver prices |

## Technical Details

### Architecture
- **Language**: Google Apps Script (JavaScript)
- **Services Used**: 
  - `SpreadsheetApp` - Google Sheets integration
  - `UrlFetchApp` - HTTP requests
  - `CacheService` - Temporary caching
  - `ScriptApp` - Trigger management
  - `Utilities` - Sleep/delay functions

### Key Functions

| Function | Purpose |
|----------|---------|
| `update_gold_silver_prices()` | Main orchestration function |
| `fetch_html_with_fallback_()` | Fetches HTML with retries and caching |
| `parse_fenegosida_rate_wrap_()` | Parser for FENEGOSIDA format |
| `parse_fngsgja_home_()` | Parser for FNGSGJA format |
| `apply_change_font_colors_()` | Color-codes price changes |
| `ensure_daily_trigger_()` | Manages automatic execution |

## License

This project is provided as-is for personal and educational use. Respect the terms of service of the data source websites when using this script.

## Contributing

To modify or extend the script:

1. Edit the `CONFIG` object for settings changes
2. Add new parsers if source websites change their HTML structure
3. Update `source_urls` if URLs change

## Support

For issues or questions:
1. Check the execution logs in the Apps Script editor
2. Verify your Google account has the necessary permissions
3. Ensure the source websites are accessible from your region
