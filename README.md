# Personal Expense Tracker

A Google Sheet + Apps Script + iPhone Shortcuts setup for logging expenses, income and transfers, with a live dashboard.

## How it works

- **iPhone Shortcut** → POSTs JSON to an Apps Script web app → a new row is inserted at the top of the `Transactions` tab (newest first).
- **Types:** `EXPENSE` (spending, including SIPs and EMIs), `INCOME`, `TRANSFER` (between your own accounts, e.g. a credit card bill payment; not counted as spending).
- **Accounts** are a fixed list in the `Accounts` tab. An unknown account is rejected by the API.
- **Category** is picked from the `Keywords` tab (first keyword contained in the description). Fallbacks: `OTHER EXPENSE`, `OTHER INCOME`, `TRANSFER`.
- **Tags** are typed by hand in the `Tags` column (comma-separated, e.g. `goa-trip, impulse`). The API never sets them.
- **Dashboard** shows income, spend, net savings, top category, income vs spend trend, category/account breakdown, daily spend and recent transactions, with Month / Account / Tag filters. TRANSFER rows are ignored.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Web app API (`doPost`) and `testDoPost` |
| `Setup.gs` | `setupSheets()` creates the `Transactions`, `Accounts` and `Keywords` tabs |
| `DashboardSetup.gs` | Builds the `Dashboard` and hidden `_Data` tabs, plus the **Dashboard** menu |

## Setup

### 1. Create the sheet and scripts
1. Create a new Google Sheet in your personal Google account.
2. Open **Extensions > Apps Script**.
3. Replace the contents of the default `Code.gs` with this repo's `Code.gs`.
4. Add two script files (**+ > Script**) named `Setup` and `DashboardSetup`, and paste in the matching contents.
5. Save the project.

### 2. Create the tabs
1. In the editor, select `setupSheets` in the function dropdown and click **Run**.
2. Approve the permission prompt (Advanced > Go to project).
3. Check the sheet now has `Transactions`, `Accounts` and `Keywords`.

### 3. Fill in the lists
- **`Accounts` tab:** one row per account or card, column A is the name (for example `SBI`, `Axis_CC`, `Cash`). Replace the samples. These names are what the shortcuts must send; case is ignored.
- **`Keywords` tab:** `Keyword | Category`, for example `zomato | FOOD`, `sip | INVESTMENT`. Order matters: the first match wins.

### 4. Build the dashboard
1. Reload the sheet. A **Dashboard** menu appears.
2. Click **Dashboard > Rebuild dashboard**.
3. Run it again only if you want to reset the dashboard layout. Adding rows, accounts or tags needs no rebuild.

### 5. Deploy the web app
1. In Apps Script: **Deploy > New deployment > type Web app**.
2. Execute as: **Me**. Who has access: **Anyone**.
3. Click **Deploy** and copy the `/exec` URL.
4. After any later code change: **Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy**. The URL stays the same.

### 6. Test the API
1. In the editor, run `testDoPost`, then check the execution log and the `Transactions` tab.
2. It writes 3 real rows (expense, income, transfer) and 1 rejected call. Delete the 3 rows afterwards.
3. Or test from a terminal (PowerShell):

```powershell
$url = "PASTE_EXEC_URL_HERE"
Invoke-RestMethod -Method Post -Uri $url -ContentType "application/json" `
  -Body '{"type":"EXPENSE","amount":450,"description":"Zomato dinner","account":"SBI"}'
```

Expected: `status: success` and the picked category.

## API

`POST <exec URL>` with a JSON body:

| Field | Required | Notes |
|---|---|---|
| `type` | yes | `EXPENSE`, `INCOME` or `TRANSFER` |
| `amount` | yes | A positive number; commas are allowed |
| `description` | yes (optional for TRANSFER) | Used for the category match |
| `account` | yes | Must exist in the `Accounts` tab; for a TRANSFER this is the source |
| `toAccount` | TRANSFER only | Destination; must differ from `account` |

Response: `{"status":"success","type":"EXPENSE","category":"FOOD"}` or `{"status":"error","message":"..."}`.

The date is always today, in the script's time zone. Set it under **Project Settings** in Apps Script (for IST, `Asia/Kolkata`).

## iPhone Shortcuts

Create three shortcuts: **Log Expense**, **Log Income** and **Log Transfer**. Each uses the same actions:

1. **Ask for Input** (Number): "How much?"
2. **Ask for Input** (Text): "What was it for?" (skip for the transfer shortcut)
3. **Choose from List**: your account names, exactly as in the `Accounts` tab. The transfer shortcut asks twice: "From" and "To".
4. **Get Contents of URL**:
   - URL: the `/exec` URL
   - Method: `POST`
   - Request Body: `JSON`, with fields `type` (`EXPENSE`, `INCOME` or `TRANSFER`), `amount`, `description`, `account`, and `toAccount` for the transfer
5. **Get Dictionary Value** for `status` and `category`, then **Show Notification**.

## Daily use

- Tap a shortcut, then enter the amount, description and account.
- Add tags later by typing them in the `Tags` column.
- Use the dashboard filters (cells C4, F4, I4) for month, account and tag.

## Notes

- Paying a credit card bill must be logged as a TRANSFER. If you log it as an EXPENSE you count the spending twice.
- The web app URL has no authentication, so keep it private.
- The old `Expenses_<account>` tabs are not read and not migrated.
- If a dashboard cell shows an error, check that `Transactions` has data and that the header row is intact.
