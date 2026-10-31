# Cashbook Combiner

One Power Query that turns a bank cashbook spread over dozens of sheets into a single table you can filter by branch, bank, currency and status.

**Live write-up:** https://chrismukonowatsauka.github.io/cashbook-combiner/

## The problem

A multi-branch retailer keeps one cashbook workbook with a sheet per bank account (USD and ZiG accounts at several banks, mobile cash deposits, rand): 50 or more sheets and over 25,000 lines a month. Answering "what did this branch bank across all accounts?" meant opening sheet after sheet, and the sheets are not laid out the same way.

## What the query does

- Reads the cashbook path from one named cell (`CashbookPath`).
- Lists every sheet and skips the non-bank ones (`Org`, `Bank charges & Imtt`). New bank sheets are picked up automatically.
- Finds the "Opening balance" row on each sheet. Its position tells the query which of the two layouts the sheet uses:
  - Layout A: Date | Posting Ref | Statement Ref | Description | Branch | DR | CR | Balance
  - Layout B: Date | Posting Ref | Description | Branch | DR | CR | Balance
- Keeps the opening line and every line with a date and an amount; converts text dates (`2/9/2026 02:00:00`) and amounts typed with commas.
- Adds Bank Sheet, Currency (from the sheet name), Status (Opening, Posted, X, No ref) and Net, keeping each sheet's original order.

Output columns: Date, Bank Sheet, Currency, Posting Ref, Status, Statement Ref, Description, Branch, DR, CR, Net, Balance.

## Files

| File | What it is |
|---|---|
| `Cashbook_Combiner_Demo.xlsx` | Setup steps, the combined table, a Summary that proves every sheet (opening + net = closing), and the query code |
| `Demo_Cashbook_Sep2026.xlsx` | Fictional cashbook with seven bank sheets in both layouts |
| `CashbookCombiner.pq` | The Power Query (M) code |
| `index.html` | The write-up page |

## Try it

1. Download both workbooks. Put the full path of `Demo_Cashbook_Sep2026.xlsx` in the yellow cell on the Start sheet.
2. Data > Get Data > From Other Sources > Blank Query > Advanced Editor, and paste `CashbookCombiner.pq`.
3. Name it `All Transactions`, then Close & Load To the existing `All Transactions` sheet (clear the preview first).
4. Add slicers for Branch, Bank Sheet, Currency and Status. Data > Refresh All each month.

All data in this repository is fictional (Kestrel Distributors, September 2026). No employer data is used.

Built by Chris Mukonowatsauka · [linkedin.com/in/cmkn](https://www.linkedin.com/in/cmkn)
