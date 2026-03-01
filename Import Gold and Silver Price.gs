/***** CONFIG (EDIT THIS ONLY) *****/
var CONFIG = {
  // ⬇️ IN THIS LINE — paste your actual spreadsheet ID
  spreadsheet_id: "PASTE_YOUR_SPREADSHEET_ID_HERE",
  
  source_urls: [
    "https://www.fngsgja.org.np/",
    "https://www.fenegosida.org/",
    "https://fenegosida.org/",
    "https://www.fenegosida.com/"
  ],

  sheet_name: "Rates",

  header_row_number: 1,
  insert_at_row_number: 2,

  trigger: {
    hour: 10,
    minute: 45
  },

  // Sheet headers (customizable)
  headers: [
    "Run Datetime",
    "Nepali Date",
    "Gold per tola",
    "Silver per tola",
    "Gold per 10 grm",
    "Silver per 10 grm",
    "Gold to Silver Ratio"
  ],

  // Number formats (no decimals)
  number_formats: {
    price_no_decimals: "#,##0",
    ratio_no_decimals: "#,##0.00"
  },

  // Font colors to indicate change vs last run
  // Increased = green, Decreased = red, Same = grey (edit colors if you want)
  change_colors: {
    up: "#1a7f37",
    down: "#d1242f",
    same: "#6e7781"
  },

  max_retries_per_url: 3,
  sleep_ms_between_retries: 1200,
  user_agent: "Mozilla/5.0 (Google Apps Script; compatible)",
  cache_minutes: 120
};
/***** END CONFIG *****/


function update_gold_silver_prices() {
  // Ensure daily trigger exists (auto-creates if missing)
  ensure_daily_trigger_();
  // ✅ Works from both manual runs AND time-based triggers
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(CONFIG.spreadsheet_id);
  }
  if (!ss) {
    Logger.log("ERROR: Could not open spreadsheet. Check CONFIG.spreadsheet_id.");
    return;
  }

  var sheet = ensure_rates_sheet_(ss, CONFIG.sheet_name);

  var fetched = fetch_html_with_fallback_(CONFIG.source_urls);
  if (!fetched || !fetched.html) {
    Logger.log("Failed to fetch HTML from all sources and cache is empty.");
    return;
  }

  // Parse (try FENEGOSIDA first, then FNGSGJA)
  var parsed = parse_fenegosida_rate_wrap_(fetched.html);
  if (!parsed) parsed = parse_fngsgja_home_(fetched.html);

  if (!parsed) {
    Logger.log("Could not parse rates. Source used: " + fetched.url_used);
    log_debug_snippet_(fetched.html);
    return;
  }

  // Calculated field: ratio (10 grm)
  var ratio_10g = "";
  if (parsed.silver_10g && Number(parsed.silver_10g) !== 0) {
    ratio_10g = Number(parsed.gold_10g) / Number(parsed.silver_10g);
  }

  // Insert newest row at top (below header)
  var insert_row = CONFIG.insert_at_row_number;
  sheet.insertRowBefore(insert_row);

  // Write values: keep as numbers (so number formatting works)
  var row_values = [
    new Date(),
    parsed.nepali_date,
    parsed.gold_tola,
    parsed.silver_tola,
    parsed.gold_10g,
    parsed.silver_10g,
    ratio_10g
  ];
  sheet.getRange(insert_row, 1, 1, row_values.length).setValues([row_values]);

  // Apply number formats (no decimals)
  apply_number_formats_(sheet, insert_row);

  // Color-code changes vs previous run (row 3 after insertion)
  var prev_row = insert_row + 1;
  if (sheet.getLastRow() >= prev_row) {
    apply_change_font_colors_(sheet, insert_row, prev_row);
  } else {
    // First run: set neutral color for all numeric fields
    set_neutral_font_colors_(sheet, insert_row);
  }

  Logger.log("Success. Source used: " + fetched.url_used);
}


function apply_number_formats_(sheet, row) {
  // Columns: 3..7 contain numeric fields (prices + ratio)
  sheet.getRange(row, 3).setNumberFormat(CONFIG.number_formats.price_no_decimals); // gold tola
  sheet.getRange(row, 4).setNumberFormat(CONFIG.number_formats.price_no_decimals); // silver tola
  sheet.getRange(row, 5).setNumberFormat(CONFIG.number_formats.price_no_decimals); // gold 10g
  sheet.getRange(row, 6).setNumberFormat(CONFIG.number_formats.price_no_decimals); // silver 10g
  sheet.getRange(row, 7).setNumberFormat(CONFIG.number_formats.ratio_no_decimals); // ratio
}


function apply_change_font_colors_(sheet, current_row, previous_row) {
  // Compare these columns: gold_tola=3, silver_tola=4, gold_10g=5, silver_10g=6, ratio=7
  var cols = [3, 4, 5, 6, 7];

  for (var i = 0; i < cols.length; i++) {
    var col = cols[i];

    var current_val = sheet.getRange(current_row, col).getValue();
    var previous_val = sheet.getRange(previous_row, col).getValue();

    var current_num = Number(current_val);
    var previous_num = Number(previous_val);

    var color = CONFIG.change_colors.same;

    // If previous is missing or non-numeric, use neutral
    if (!isNaN(current_num) && !isNaN(previous_num)) {
      if (current_num > previous_num) color = CONFIG.change_colors.up;
      else if (current_num < previous_num) color = CONFIG.change_colors.down;
      else color = CONFIG.change_colors.same;
    }

    sheet.getRange(current_row, col).setFontColor(color);
  }
}


function set_neutral_font_colors_(sheet, row) {
  // Set numeric columns to neutral color
  sheet.getRange(row, 3, 1, 5).setFontColor(CONFIG.change_colors.same);
}


function ensure_rates_sheet_(ss, sheet_name) {
  var sheet = ss.getSheetByName(sheet_name);
  if (!sheet) sheet = ss.insertSheet(sheet_name);

  var header_row = CONFIG.header_row_number;

  // If empty, write headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange(header_row, 1, 1, CONFIG.headers.length).setValues([CONFIG.headers]);
    return sheet;
  }

  // If header size differs, rewrite headers to match config
  var existing = sheet.getRange(header_row, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existing.length !== CONFIG.headers.length) {
    sheet.getRange(header_row, 1, 1, CONFIG.headers.length).setValues([CONFIG.headers]);
  }

  return sheet;
}


function fetch_html_with_fallback_(urls) {
  var cache = CacheService.getScriptCache();
  var cache_key = "rates_last_good_html";

  for (var i = 0; i < urls.length; i++) {
    var url = urls[i];

    for (var attempt = 1; attempt <= CONFIG.max_retries_per_url; attempt++) {
      try {
        var response = UrlFetchApp.fetch(url, {
          method: "get",
          followRedirects: true,
          muteHttpExceptions: true,
          headers: { "User-Agent": CONFIG.user_agent }
        });

        var code = response.getResponseCode();
        var html = response.getContentText();

        if (code >= 200 && code < 300 && html && html.length > 2000) {
          cache.put(cache_key, html, CONFIG.cache_minutes * 60);
          return { html: html, url_used: url };
        }

        Logger.log("Fetch not usable. url=" + url + ", attempt=" + attempt + ", code=" + code);
      } catch (e) {
        Logger.log("Fetch exception. url=" + url + ", attempt=" + attempt + ", error=" + e);
      }

      Utilities.sleep(CONFIG.sleep_ms_between_retries);
    }
  }

  var cached = cache.get(cache_key);
  if (cached) {
    Logger.log("Using cached HTML because all sources failed.");
    return { html: cached, url_used: "CACHE" };
  }

  return null;
}


/* -------- PARSER 1: FENEGOSIDA rate-content-wrap -------- */
function parse_fenegosida_rate_wrap_(html) {
  if (html.indexOf('class="rate-content-wrap"') === -1) return null;

  var wrap = extract_first_div_block_(html, /<div\s+class="rate-content-wrap">/i);
  if (!wrap) return null;

  var day = extract_text_(wrap, /<div\s+class="rate-date-day">\s*([^<]+)\s*<\/div>/i);
  var month = extract_text_(wrap, /<div\s+class="rate-date-month">\s*([^<]+)\s*<\/div>/i);
  var year = extract_text_(wrap, /<div\s+class="rate-date-year">\s*([^<]+)\s*<\/div>/i);
  var nepali_date = (day && month && year) ? (day.trim() + " " + month.trim() + " " + year.trim()) : "";

  var header_blocks = extract_all_div_blocks_(wrap, /<div\s+id="header-rate">/ig);
  if (!header_blocks || header_blocks.length < 2) return null;

  var gms_block = header_blocks[0];
  var tola_block = header_blocks[1];

  var gold_10g = extract_price_from_header_(gms_block, "FINE GOLD (9999)");
  var silver_10g = extract_price_from_header_(gms_block, "SILVER");
  var gold_tola = extract_price_from_header_(tola_block, "FINE GOLD (9999)");
  var silver_tola = extract_price_from_header_(tola_block, "SILVER");

  if (gold_10g === null || silver_10g === null || gold_tola === null || silver_tola === null) return null;

  return { nepali_date: nepali_date, gold_tola: gold_tola, silver_tola: silver_tola, gold_10g: gold_10g, silver_10g: silver_10g };
}


function extract_price_from_header_(header_html, label_text) {
  var safe_label = escape_regex_(label_text);
  var regex = new RegExp(
    "<p>\\s*" + safe_label + "[\\s\\S]*?<b>\\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\\.[0-9]+)?|[0-9]+(?:\\.[0-9]+)?)\\s*<\\/b>",
    "i"
  );

  var match = header_html.match(regex);
  if (!match || !match[1]) return null;

  var raw = match[1].trim().replace(/,/g, "");
  var value = Number(raw);
  return isNaN(value) ? null : value;
}


/* -------- PARSER 2: FNGSGJA text layout -------- */
function parse_fngsgja_home_(html) {
  var nepali_date = "";
  var date_match = html.match(/(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})\s*,/);
  if (date_match) nepali_date = date_match[1] + " " + date_match[2] + " " + date_match[3];

  var gold_tola = extract_fngsgja_rate_(html, "FINE GOLD", "Per 1 tola");
  var gold_10g = extract_fngsgja_rate_(html, "FINE GOLD", "Per 10 gms");
  var silver_tola = extract_fngsgja_rate_(html, "SILVER", "Per 1 tola");
  var silver_10g = extract_fngsgja_rate_(html, "SILVER", "Per 10 gms");

  if (gold_tola === null || gold_10g === null || silver_tola === null || silver_10g === null) return null;

  return { nepali_date: nepali_date, gold_tola: gold_tola, silver_tola: silver_tola, gold_10g: gold_10g, silver_10g: silver_10g };
}


function extract_fngsgja_rate_(html, metal_label, unit_label) {
  var safe_metal = escape_regex_(metal_label);
  var safe_unit = escape_regex_(unit_label);

  var regex = new RegExp(
    safe_metal + "[\\s\\S]*?" + safe_unit + "\\s*:?[\\s\\S]*?NRs\\.?\\s*([0-9,]+(?:\\.[0-9]+)?)",
    "i"
  );

  var match = html.match(regex);
  if (!match || !match[1]) return null;

  var raw = match[1].trim().replace(/,/g, "");
  var value = Number(raw);
  return isNaN(value) ? null : value;
}


/* -------- HELPERS -------- */
function extract_first_div_block_(html, start_regex) {
  var start_match = start_regex.exec(html);
  if (!start_match) return null;

  var start_index = start_match.index;
  var slice = html.substring(start_index);

  var open_div = /<div\b/ig;
  var close_div = /<\/div>/ig;

  var depth = 0;
  var pos = 0;

  while (true) {
    open_div.lastIndex = pos;
    close_div.lastIndex = pos;

    var o = open_div.exec(slice);
    var c = close_div.exec(slice);

    if (!o && !c) break;

    if (o && (!c || o.index < c.index)) {
      depth++;
      pos = o.index + 4;
    } else if (c) {
      depth--;
      pos = c.index + 6;
      if (depth === 0) return slice.substring(0, pos);
    }
  }

  return null;
}


function extract_all_div_blocks_(html, start_regex) {
  var blocks = [];
  var match;

  while ((match = start_regex.exec(html)) !== null) {
    var start_index = match.index;
    var slice = html.substring(start_index);

    var open_div = /<div\b/ig;
    var close_div = /<\/div>/ig;

    var depth = 0;
    var pos = 0;

    while (true) {
      open_div.lastIndex = pos;
      close_div.lastIndex = pos;

      var o = open_div.exec(slice);
      var c = close_div.exec(slice);

      if (!o && !c) break;

      if (o && (!c || o.index < c.index)) {
        depth++;
        pos = o.index + 4;
      } else if (c) {
        depth--;
        pos = c.index + 6;
        if (depth === 0) {
          blocks.push(slice.substring(0, pos));
          break;
        }
      }
    }
  }

  return blocks;
}


function extract_text_(html, regex) {
  var match = html.match(regex);
  return (match && match[1]) ? match[1] : null;
}


function escape_regex_(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}


function log_debug_snippet_(html) {
  var lower = html.toLowerCase();
  var idx = lower.indexOf("fine gold");
  if (idx === -1) idx = lower.indexOf("silver");
  if (idx === -1) idx = lower.indexOf("rate-content-wrap");

  if (idx === -1) {
    Logger.log("No obvious keywords found in HTML.");
    return;
  }

  Logger.log(html.substring(Math.max(0, idx - 500), Math.min(html.length, idx + 2000)));
}


/* -------- TRIGGER -------- */
function ensure_daily_trigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "update_gold_silver_prices") {
      triggerExists = true;
      break;
    }
  }
  
  if (!triggerExists) {
    create_daily_trigger_();
  }
}


function create_daily_trigger_() {
  // Deletes existing triggers for update_gold_silver_prices, then creates a new one
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "update_gold_silver_prices") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("update_gold_silver_prices")
    .timeBased()
    .atHour(CONFIG.trigger.hour)
    .nearMinute(CONFIG.trigger.minute)
    .everyDays(1)
    .create();
}
