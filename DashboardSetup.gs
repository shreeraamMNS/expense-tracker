/**
 * Builds a live "Dashboard" tab inside this spreadsheet.
 * Run setupDashboard() once (after setupSheets()). Not needed again when you add rows, accounts or tags.
 *
 * How it works:
 *  - A hidden "_Data" tab pulls the EXPENSE rows of the "Transactions" tab into one clean table (formulas).
 *    INCOME and TRANSFER rows are not counted as spending.
 *  - The "Dashboard" tab reads from _Data with formulas + native charts, so it updates as rows are added.
 */
const DASH_TAB = "Dashboard";
const DATA_TAB = "_Data";
const TREND_MONTHS = 6;      // bars in the monthly trend chart
const MAX_ROWS = 3000;       // max expense rows the dashboard reads
const PALETTE = ["#6366f1", "#f59e0b", "#10b981", "#ef4444", "#06b6d4", "#ec4899", "#8b5cf6", "#84cc16", "#f97316", "#14b8a6", "#64748b"];
const INR_CELL = '[>=10000000]"₹"##\\,##\\,##\\,##0;[>=100000]"₹"##\\,##\\,##0;"₹"##,##0';
const INR_TEXT = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0';

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Dashboard")
    .addItem("Setup sheets", "setupSheets")
    .addItem("Rebuild dashboard", "setupDashboard")
    .addToUi();
}

function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  if (!ss.getSheetByName(TX_TAB)) throw new Error('No "' + TX_TAB + '" tab found. Run Dashboard > Setup sheets first.');

  const data = resetSheet_(ss, DATA_TAB);
  const dash = resetSheet_(ss, DASH_TAB);

  buildData_(data);
  buildDashboard_(dash, data, tz);

  data.hideSheet();
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

function resetSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.showSheet();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clear().clearDataValidations();
  sh.setConditionalFormatRules([]);
  return sh;
}

// ---------------------------------------------------------------- helper tab
function buildData_(data) {
  if (data.getMaxRows() < MAX_ROWS + 1) data.insertRowsAfter(data.getMaxRows(), MAX_ROWS + 1 - data.getMaxRows());
  const R = MAX_ROWS + 1;
  const D = "Dashboard";
  const T = TX_TAB + "!";

  // Transactions columns: A Date, B Type, C Description, D Amount, E Account, G Category, H Tags
  const isExpense = T + 'A1:A<>"", ' + T + 'B1:B="EXPENSE"';
  data.getRange("A1:H1").setValues([["Date", "Description", "Amount", "Category", "Account", "Month", "DateNum", "TagKey"]]);
  data.getRange("A2").setFormula(
    "=IFERROR(FILTER({" + T + "A1:A, " + T + "C1:C, " + T + "D1:D, " + T + "G1:G, " + T + "E1:E}, " + isExpense + '),"")'
  );
  // DateNum works whether the sheet stored a real date or a text date
  data.getRange("G2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(ISNUMBER(A2:A),A2:A,DATEVALUE(A2:A&""))))');
  data.getRange("F2").setFormula('=ARRAYFORMULA(IF(G2:G="","",TEXT(G2:G,"yyyy-mm")))');
  // ",goa-trip,impulse," so a tag filter can match whole tags only; rows line up with A:E (same FILTER)
  data.getRange("H2").setFormula(
    '=IFERROR(LET(t,FILTER(' + T + 'H1:H, ' + isExpense + '),ARRAYFORMULA(IF(t="","",","&REGEXREPLACE(LOWER(t&""),"\\s*,\\s*",",")&","))),"")'
  );

  // dropdown sources
  data.getRange("I1").setValue("All");
  data.getRange("I2").setFormula('=IFERROR(SORT(UNIQUE(FILTER({F2:F' + R + ';AC2:AC' + R + '},{F2:F' + R + ';AC2:AC' + R + '}<>"")),1,FALSE))');
  data.getRange("K1").setValue("All");
  data.getRange("K2").setFormula('=IFERROR(SORT(UNIQUE(FILTER(E2:E' + R + ',E2:E' + R + '<>""))))');
  data.getRange("W1").setValue("All");
  data.getRange("W2").setFormula('=IFERROR(SORT(UNIQUE(TRANSPOSE(SPLIT(TEXTJOIN(",",TRUE,H2:H' + R + ',AE2:AE' + R + '),",")))))');

  // INCOME rows (Y:AE) - Date, Amount, Category, Account, Month, DateNum, TagKey
  const isIncome = T + 'A1:A<>"", ' + T + 'B1:B="INCOME"';
  data.getRange("Y1:AE1").setValues([["Date", "Amount", "Category", "Account", "Month", "DateNum", "TagKey"]]);
  data.getRange("Y2").setFormula(
    "=IFERROR(FILTER({" + T + "A1:A, " + T + "D1:D, " + T + "G1:G, " + T + "E1:E}, " + isIncome + '),"")'
  );
  data.getRange("AD2").setFormula('=ARRAYFORMULA(IF(Y2:Y="","",IF(ISNUMBER(Y2:Y),Y2:Y,DATEVALUE(Y2:Y&""))))');
  data.getRange("AC2").setFormula('=ARRAYFORMULA(IF(AD2:AD="","",TEXT(AD2:AD,"yyyy-mm")))');
  data.getRange("AE2").setFormula(
    '=IFERROR(LET(t,FILTER(' + T + 'H1:H, ' + isIncome + '),ARRAYFORMULA(IF(t="","",","&REGEXREPLACE(LOWER(t&""),"\\s*,\\s*",",")&","))),"")'
  );

  // criteria cells (driven by the Month / Account / Tag dropdowns on the Dashboard)
  const labels = ["where: month+account+tag", "where: month+tag", "where: account+tag", "selected month", "days in selected month", "previous month", "days for daily avg"];
  data.getRange("L1:L7").setValues(labels.map(function (l) { return [l]; }));
  const tagPart = 'IF(' + D + '!$I$4="All",""," and H contains \',"&LOWER(TRIM(' + D + '!$I$4))&",\'")';
  const monthPart = 'IF(' + D + '!$C$4="All","", " and F = \'"&' + D + '!$C$4&"\'")';
  const accPart = 'IF(' + D + '!$F$4="All","", " and E = \'"&' + D + '!$F$4&"\'")';
  const f = [
    '=' + monthPart + '&' + accPart + '&' + tagPart,
    '=' + monthPart + '&' + tagPart,
    '=' + accPart + '&' + tagPart,
    '=IF(' + D + '!$C$4="All",INDEX(I2:I,1),' + D + '!$C$4)',
    '=DAY(EOMONTH(DATEVALUE(M4&"-01"),0))',
    '=IF(' + D + '!$C$4="All","",TEXT(EDATE(DATEVALUE(' + D + '!$C$4&"-01"),-1),"yyyy-mm"))',
    '=IF(' + D + '!$C$4="All",MAX(G2:G' + R + ')-MIN(G2:G' + R + ')+1,IF(' + D + '!$C$4=TEXT(TODAY(),"yyyy-mm"),DAY(TODAY()),DAY(EOMONTH(DATEVALUE(' + D + '!$C$4&"-01"),0))))'
  ];
  f.forEach(function (formula, i) { data.getRange(i + 1, 13).setFormula(formula); }); // M1:M7

  // income criteria (M8): QUERY letters are relative to the range Y:AE -> Month=E, Account=D, Tag=G
  data.getRange("L8").setValue("income where: month+account+tag");
  data.getRange("M8").setFormula('=' + monthPart.replace(" F = ", " E = ") + '&' + accPart.replace(" E = ", " D = ") + '&' + tagPart.replace(" H contains", " G contains"));

  // monthly trend (last N calendar months, filtered by the account and tag dropdowns)
  const N = TREND_MONTHS;
  const accCond = '(((E2:E' + R + '=' + D + '!$F$4)+(' + D + '!$F$4="All"))>0)*' +
    '(((ISNUMBER(SEARCH(","&LOWER(TRIM(' + D + '!$I$4))&",",H2:H' + R + ")))+(" + D + '!$I$4="All"))>0)';
  const incCond = '(((AB2:AB' + R + '=' + D + '!$F$4)+(' + D + '!$F$4="All"))>0)*' +
    '(((ISNUMBER(SEARCH(","&LOWER(TRIM(' + D + '!$I$4))&",",AE2:AE' + R + ")))+(" + D + '!$I$4="All"))>0)';
  data.getRange("O1:S1").setValues([["Month start", "Month", "Income", "Spend", "Key"]]);
  data.getRange("O2").setFormula('=ARRAYFORMULA(DATE(YEAR(TODAY()),MONTH(TODAY())+SEQUENCE(' + N + ',1,-' + (N - 1) + ',1),1))');
  data.getRange("P2").setFormula('=ARRAYFORMULA(TEXT(O2:O' + (N + 1) + ',"mmm yy"))');
  data.getRange("S2").setFormula('=ARRAYFORMULA(TEXT(O2:O' + (N + 1) + ',"yyyy-mm"))');
  data.getRange("Q2").setFormula('=BYROW(S2:S' + (N + 1) + ',LAMBDA(m,SUMPRODUCT((AC2:AC' + R + '=m)*' + incCond + '*Z2:Z' + R + ')))');
  data.getRange("R2").setFormula('=BYROW(S2:S' + (N + 1) + ',LAMBDA(m,SUMPRODUCT((F2:F' + R + '=m)*' + accCond + '*C2:C' + R + ')))');

  // daily spend for the selected month (latest month when "All")
  data.getRange("T1:U1").setValues([["Day", "Spend"]]);
  data.getRange("T2").setFormula("=SEQUENCE(31)");
  data.getRange("U2").setFormula(
    '=BYROW(T2:T32,LAMBDA(d,IF(d>$M$5,"",SUMPRODUCT((G2:G' + R + '=DATE(VALUE(LEFT($M$4,4)),VALUE(RIGHT($M$4,2)),d))*' + accCond + '*C2:C' + R + '))))'
  );
}

// ---------------------------------------------------------------- dashboard tab
function buildDashboard_(dash, data, tz) {
  const DT = "'" + DATA_TAB + "'!";
  const R = MAX_ROWS + 1;
  const SRC = DT + "A1:H" + R;

  dash.setHiddenGridlines(true);
  dash.getRange("A1:J80").setBackground("#f4f5fa").setFontFamily("Roboto").setFontColor("#1e2230");
  dash.setColumnWidth(1, 16);
  for (let c = 2; c <= 9; c++) dash.setColumnWidth(c, 112);
  dash.setColumnWidth(10, 16);

  // title + subtitle
  dash.setRowHeight(2, 38);
  dash.getRange("B2:F2").merge().setValue("Finance Dashboard").setFontSize(22).setFontWeight("bold");
  dash.getRange("B3:F3").merge().setFontColor("#7a8194")
    .setFormula('=IF(C4="All","All time",TEXT(DATEVALUE(C4&"-01"),"mmmm yyyy"))&"  ·  "&IF(F4="All","All accounts",F4)&IF(I4="All",""," · #"&I4)');

  // filters
  dash.setRowHeight(4, 34);
  dash.getRange("B4").setValue("Month").setHorizontalAlignment("right").setFontColor("#7a8194");
  dash.getRange("E4").setValue("Account").setHorizontalAlignment("right").setFontColor("#7a8194");
  dash.getRange("H4").setValue("Tag").setHorizontalAlignment("right").setFontColor("#7a8194");
  const monthRule = SpreadsheetApp.newDataValidation().requireValueInRange(data.getRange("I1:I60"), true).setAllowInvalid(true).build();
  const accRule = SpreadsheetApp.newDataValidation().requireValueInRange(data.getRange("K1:K30"), true).setAllowInvalid(true).build();
  const tagRule = SpreadsheetApp.newDataValidation().requireValueInRange(data.getRange("W1:W200"), true).setAllowInvalid(true).build();
  dash.getRange("C4").setNumberFormat("@").setValue(Utilities.formatDate(new Date(), tz, "yyyy-MM")).setDataValidation(monthRule);
  dash.getRange("F4").setNumberFormat("@").setValue("All").setDataValidation(accRule);
  dash.getRange("I4").setNumberFormat("@").setValue("All").setDataValidation(tagRule);
  ["C4", "F4", "I4"].forEach(function (a) {
    dash.getRange(a).setBackground("#ffffff").setFontWeight("bold").setHorizontalAlignment("center")
      .setBorder(true, true, true, true, false, false, "#c9cde0", SpreadsheetApp.BorderStyle.SOLID);
  });
  dash.getRange("B4:I4").setVerticalAlignment("middle");

  // KPI cards (rows 6-8)
  dash.setRowHeights(6, 1, 24); dash.setRowHeights(7, 1, 40); dash.setRowHeights(8, 1, 26);
  [["B", "C"], ["D", "E"], ["F", "G"], ["H", "I"]].forEach(function (p) {
    for (let r = 6; r <= 8; r++) dash.getRange(p[0] + r + ":" + p[1] + r).merge();
    dash.getRange(p[0] + "6:" + p[1] + "8").setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("left")
      .setBorder(true, true, true, true, false, false, "#e3e5ee", SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(p[0] + "6").setFontSize(9).setFontColor("#7a8194").setFontWeight("bold");
    dash.getRange(p[0] + "7").setFontSize(20).setFontWeight("bold");
    dash.getRange(p[0] + "8").setFontSize(9).setFontColor("#7a8194");
  });
  dash.getRange("B6").setValue("INCOME");
  dash.getRange("D6").setValue("TOTAL SPENT");
  dash.getRange("F6").setValue("NET SAVINGS");
  dash.getRange("H6").setValue("TOP CATEGORY");

  // SUM() ignores the header text QUERY may or may not return, so this works either way
  const q1 = function (agg) {
    return '=IFERROR(SUM(QUERY(' + SRC + ',"select ' + agg + '(C) where C is not null"&' + DT + 'M1,1)),0)';
  };
  const INC = DT + "Y1:AE" + R;
  dash.getRange("B7").setFormula('=IFERROR(SUM(QUERY(' + INC + ',"select sum(B) where B is not null"&' + DT + 'M8,1)),0)').setNumberFormat(INR_CELL);
  dash.getRange("B8").setFormula('=IFERROR(COUNT(QUERY(' + INC + ',"select B where B is not null"&' + DT + 'M8,1))&" entr"&IF(COUNT(QUERY(' + INC + ',"select B where B is not null"&' + DT + 'M8,1))=1,"y","ies"),"")');
  dash.getRange("D7").setFormula(q1("sum")).setNumberFormat(INR_CELL);
  dash.getRange("D8").setFormula(
    '="₹"&TEXT(IFERROR(D7/' + DT + 'M7,0),"' + INR_TEXT + '")&"/day"&IF($C$4="All","",IFERROR(LET(p,SUM(QUERY(' + SRC + ',"select sum(C) where C is not null and F = \'"&' + DT + 'M6&"\'"&' + DT + 'M3,1)),' +
    'IF(p=0,""," · "&IF(D7>=p,"▲ ","▼ ")&TEXT(ABS(D7-p)/p,"0%")&" vs last month")),""))'
  );
  dash.getRange("F7").setFormula('=B7-D7').setNumberFormat(INR_CELL);
  dash.getRange("F8").setFormula('=IF(B7=0,"No income in view","Savings rate "&TEXT((B7-D7)/B7,"0%"))');
  dash.getRange("H7").setFormula('=IF(B28="","-",B28)').setFontSize(16);
  dash.getRange("H8").setFormula('=IFERROR("₹"&TEXT(C28,"' + INR_TEXT + '")&" · "&TEXT(C28/D7,"0%")&" of spend","")');

  dash.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#ef4444").setRanges([dash.getRange("F7:G7")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor("#10b981").setRanges([dash.getRange("F7:G7")]).build()
  ]);

  // section titles
  [["B10", "INCOME VS SPEND"], ["G10", "SPEND BY CATEGORY"], ["B26", "CATEGORY BREAKDOWN"], ["G26", "BY ACCOUNT"],
   ["B42", "DAILY SPEND"], ["B59", "RECENT TRANSACTIONS"]].forEach(function (t) {
    dash.getRange(t[0]).setValue(t[1]).setFontSize(10).setFontWeight("bold").setFontColor("#7a8194");
  });

  // category table (B27:D40) + inline bars (E)
  dash.getRange("B27").setFormula(
    '=IFERROR(ARRAY_CONSTRAIN(QUERY(' + SRC + ',"select D, sum(C) where D is not null"&' + DT + 'M1&" group by D order by sum(C) desc label D \'Category\', sum(C) \'Amount\'",1),14,2),{"Category","Amount"})'
  );
  dash.getRange("D27").setValue("Share");
  dash.getRange("D28").setFormula('=ARRAYFORMULA(IF(C28:C40="","",IFERROR(C28:C40/$D$7,0)))');
  const bars = [];
  for (let r = 28; r <= 40; r++) bars.push(['=IF(D' + r + '="","",SPARKLINE(D' + r + ',{"charttype","bar";"max",1;"color1","#6366f1"}))']);
  dash.getRange("E28:E40").setFormulas(bars);
  dash.getRange("C28:C40").setNumberFormat(INR_CELL);
  dash.getRange("D28:D40").setNumberFormat("0%");

  // account table (G27:I40)
  dash.getRange("G27").setFormula(
    '=IFERROR(ARRAY_CONSTRAIN(QUERY(' + SRC + ',"select E, sum(C) where E is not null"&' + DT + 'M2&" group by E order by sum(C) desc label E \'Account\', sum(C) \'Amount\'",1),14,2),{"Account","Amount"})'
  );
  dash.getRange("I27").setValue("Share");
  dash.getRange("I28").setFormula('=ARRAYFORMULA(IF(H28:H40="","",IFERROR(H28:H40/SUM(H28:H40),0)))');
  dash.getRange("H28:H40").setNumberFormat(INR_CELL);
  dash.getRange("I28:I40").setNumberFormat("0%");

  // recent transactions (B60:F75)
  dash.getRange("B60").setFormula(
    '=IFERROR(QUERY(' + SRC + ',"select G, D, E, C, B where B is not null"&' + DT + 'M1&" order by G desc limit 15 label G \'Date\', D \'Category\', E \'Account\', C \'Amount\', B \'Description\'",1),"No transactions yet")'
  );
  dash.getRange("B61:B75").setNumberFormat("dd mmm yyyy");
  dash.getRange("E61:E75").setNumberFormat(INR_CELL);

  // table styling
  [["B27:E40", "B27:E27"], ["G27:I40", "G27:I27"], ["B60:I75", "B60:I60"]].forEach(function (t) {
    dash.getRange(t[0]).setBackground("#ffffff").setBorder(true, true, true, true, false, false, "#e3e5ee", SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(t[1]).setFontWeight("bold").setFontColor("#7a8194").setFontSize(10).setBackground("#fafbfe");
  });

  // charts (native, they follow the dropdowns)
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(data.getRange("P1:R" + (TREND_MONTHS + 1))).setNumHeaders(1).setPosition(11, 2, 0, 0)
    .setOption("title", "Income vs spend, last " + TREND_MONTHS + " months (follows account + tag filters)").setOption("width", 556).setOption("height", 300)
    .setOption("legend", { position: "bottom" }).setOption("colors", ["#10b981", "#6366f1"]).setOption("backgroundColor", "#ffffff").build());

  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(dash.getRange("B27:C40")).setNumHeaders(1).setPosition(11, 7, 4, 0)
    .setOption("title", "By category").setOption("width", 332).setOption("height", 300)
    .setOption("pieHole", 0.55).setOption("colors", PALETTE).setOption("legend", { position: "bottom" })
    .setOption("backgroundColor", "#ffffff").build());

  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(data.getRange("T1:U32")).setNumHeaders(1).setPosition(43, 2, 0, 0)
    .setOption("title", "Spend per day (selected month, latest month when Month = All)").setOption("width", 892).setOption("height", 300)
    .setOption("legend", { position: "none" }).setOption("colors", ["#06b6d4"]).setOption("backgroundColor", "#ffffff").build());
}
