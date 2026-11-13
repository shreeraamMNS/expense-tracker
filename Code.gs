/**
 * Personal finance tracker - API (web app).
 * POST JSON: { type: "EXPENSE"|"INCOME"|"TRANSFER", amount, description, account, toAccount (TRANSFER only) }
 * New rows are inserted at the top of the "Transactions" tab (newest first).
 */
const TX_TAB = "Transactions";
const ACCOUNTS_TAB = "Accounts";
const KEYWORDS_TAB = "Keywords";
const TX_HEADERS = ["Date", "Type", "Description", "Amount", "Account", "To Account", "Category", "Tags", "Notes"];
const TYPES = ["EXPENSE", "INCOME", "TRANSFER"];
const FALLBACK_CATEGORY = { EXPENSE: "OTHER EXPENSE", INCOME: "OTHER INCOME", TRANSFER: "TRANSFER" };

function getCategoryForDescription(description, type) {
  const keywordsTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYWORDS_TAB);
  if (keywordsTab) {
    const data = keywordsTab.getDataRange().getValues();
    const lowerDesc = description.toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const keyword = data[i][0].toString().trim().toLowerCase();
      if (keyword && lowerDesc.includes(keyword)) return data[i][1];
    }
  }
  return FALLBACK_CATEGORY[type];
}

function getAccounts_() {
  const tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_TAB);
  if (!tab || tab.getLastRow() < 2) return [];
  return tab.getRange(2, 1, tab.getLastRow() - 1, 1).getValues()
    .map(function (r) { return r[0].toString().trim(); })
    .filter(function (a) { return a; });
}

// returns the canonical account name from the Accounts tab (case-insensitive), or null
function resolveAccount_(name, accounts) {
  const wanted = (name || "").toString().trim().toLowerCase();
  for (let i = 0; i < accounts.length; i++) {
    if (accounts[i].toLowerCase() === wanted) return accounts[i];
  }
  return null;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(message) {
  return jsonResponse_({ status: "error", message: message });
}

function doPost(request) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    const body = JSON.parse(request.postData.contents);

    const type = (body.type || "").toString().trim().toUpperCase();
    if (TYPES.indexOf(type) === -1) return fail_("Invalid type: " + body.type + " (use " + TYPES.join("/") + ")");

    const amount = Number((body.amount === undefined || body.amount === null ? "" : body.amount).toString().replace(/,/g, ""));
    if (!isFinite(amount) || amount <= 0) return fail_("Invalid amount: " + body.amount);

    const accounts = getAccounts_();
    const account = resolveAccount_(body.account, accounts);
    if (!account) return fail_("Unknown account: " + body.account);

    let toAccount = "";
    if (type === "TRANSFER") {
      toAccount = resolveAccount_(body.toAccount, accounts);
      if (!toAccount) return fail_("Unknown to-account: " + body.toAccount);
      if (toAccount === account) return fail_("Transfer needs two different accounts");
    }

    const description = (body.description || "").toString().trim() || (type === "TRANSFER" ? "Transfer" : "");
    if (!description) return fail_("Description is required");

    const category = getCategoryForDescription(description, type);

    const tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TX_TAB);
    if (!tab) return fail_("Tab not found: " + TX_TAB);

    const today = new Date();
    const date = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    lock.waitLock(10000);
    locked = true;
    tab.insertRowsAfter(1, 1);
    const row = tab.getRange(2, 1, 1, TX_HEADERS.length);
    // the new row inherits the header's look; take format + validation from the row below instead
    if (tab.getLastRow() >= 3) {
      const below = tab.getRange(3, 1, 1, TX_HEADERS.length);
      below.copyTo(row, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      below.copyTo(row, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    row.setValues([[date, type, description, amount, account, toAccount, category, "", ""]]);
    tab.getRange(2, 1).setNumberFormat("yyyy-mm-dd");
    SpreadsheetApp.flush();

    return jsonResponse_({ status: "success", type: type, category: category });
  } catch (error) {
    return fail_(error.message);
  } finally {
    if (locked) lock.releaseLock();
  }
}

function testDoPost() {
  const mock = function (payload) { return { postData: { contents: JSON.stringify(payload) } }; };
  Logger.log(doPost(mock({ type: "EXPENSE", amount: 450, description: "Zomato dinner", account: "Axis_CC" })).getContent());
  Logger.log(doPost(mock({ type: "INCOME", amount: 50000, description: "Salary", account: "SBI" })).getContent());
  Logger.log(doPost(mock({ type: "TRANSFER", amount: 10000, description: "CC bill", account: "SBI", toAccount: "Axis_CC" })).getContent());
  Logger.log(doPost(mock({ type: "EXPENSE", amount: 1, description: "x", account: "Nope" })).getContent());
}
