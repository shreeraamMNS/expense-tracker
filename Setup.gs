/**
 * One-time sheet setup. Run setupSheets() (or Dashboard menu > Setup sheets).
 * Safe to re-run: existing tabs and their data are left untouched; only missing tabs are created.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const accounts = getOrCreate_(ss, ACCOUNTS_TAB, ["Account", "Kind"], [
    ["SBI", "Bank"], ["Axis_CC", "Credit Card"], ["Cash", "Cash"]
  ]);
  accounts.setColumnWidth(1, 160);

  const keywords = getOrCreate_(ss, KEYWORDS_TAB, ["Keyword", "Category"], [
    ["zomato", "FOOD"], ["swiggy", "FOOD"], ["uber", "TRANSPORT"], ["sip", "INVESTMENT"], ["salary", "SALARY"]
  ]);
  keywords.setColumnWidth(1, 160);
  keywords.setColumnWidth(2, 160);

  let tx = ss.getSheetByName(TX_TAB);
  if (!tx) {
    tx = ss.insertSheet(TX_TAB, 0);
    tx.getRange(1, 1, 1, TX_HEADERS.length).setValues([TX_HEADERS]);
    tx.setFrozenRows(1);
    tx.getRange("A1:I1").setFontWeight("bold").setBackground("#6366f1").setFontColor("#ffffff");
    tx.getRange("A2:A").setNumberFormat("yyyy-mm-dd");
    tx.getRange("D2:D").setNumberFormat("#,##0.00");
    [90, 90, 240, 100, 120, 120, 140, 200, 200].forEach(function (w, i) { tx.setColumnWidth(i + 1, w); });
  }

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(false).build();
  const accRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(ACCOUNTS_TAB).getRange("A2:A"), true).setAllowInvalid(false).build();
  tx.getRange("B2:B").setDataValidation(typeRule);
  tx.getRange("E2:F").setDataValidation(accRule);
  tx.getRange("H1").setNote("Add tags by hand, comma-separated, e.g. goa-trip, impulse. The dashboard Tag filter reads this column.");

  SpreadsheetApp.getUi().alert("Setup done. Edit the Accounts and Keywords tabs, then run Dashboard > Rebuild dashboard.");
}

function getOrCreate_(ss, name, headers, sampleRows) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#6366f1").setFontColor("#ffffff");
  sh.setFrozenRows(1);
  if (sampleRows && sampleRows.length) sh.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  return sh;
}
