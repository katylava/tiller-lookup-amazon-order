# Tiller Amazon Order Lookup

> **Note:** This code was written entirely by Claude (Anthropic) and has not been
> reviewed by a human. It has only been QA'd by a human.

A Google Apps Script add-on for [Tiller](https://www.tillerhq.com/) spreadsheets
that matches Amazon transactions to order details (date and items) from an
exported Amazon order history.

## Features

- **Lookup selected amount** — Select a cell, and the add-on searches for the
  absolute dollar value in "tmp-amazon". Matching order date and items are shown
  in a pop-up dialog and added as a note to the cell.

- **Lookup all uncategorized Amazon transactions** — Bulk processes the
  "Transactions" sheet. Finds all rows with a Description starting with
  "Amazon", skips rows that already have notes, and stops when it hits a row
  with a Category value. Adds order notes to each matching Amount cell and shows
  a summary when done.

## Everyday Usage

### Preparing the tmp-amazon Sheet

1. Export your Amazon order history as CSV using the
   [Amazon Order History Reporter](https://chromewebstore.google.com/detail/amazon-order-history-repo/mgkilgclilajckgnedgjgnfdokkgnibi)
   Chrome extension
2. Create a new tab in the Tiller spreadsheet named **tmp-amazon**
3. Paste the CSV file contents into the sheet
4. Select the pasted column, then use **Data → Split text to columns** to
   split on commas
5. Verify the resulting columns include `total`, `date`, and `items`

When you're done looking up orders, delete the tmp-amazon tab or at least clear
its contents — it's not needed after the notes have been added.

## Initial Setup

1. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
2. Log in: `clasp login`
3. Enable the Apps Script API at https://script.google.com/home/usersettings
4. Create the standalone Apps Script project:
   ```
   clasp create --type standalone --title "Amazon Order Lookup"
   ```
5. Push the code: `clasp push`
6. Open the script project in the browser: `clasp open-script`
7. Go to **Deploy → Test deployments**
8. Set the application type to **Editor Add-on**
9. Select your Tiller spreadsheet as the test document
10. Click **Execute** — this will open the spreadsheet with the add-on active

## Development

### Making Changes

1. Edit `Code.gs` locally
2. `clasp push --force`
3. Open the Apps Script project (`clasp open-script`) and execute a new **test
   deployment** — reloading the spreadsheet alone is not enough

### Installing in a New Spreadsheet

1. Open the Apps Script project: `clasp open-script`
2. Go to **Deploy → Test deployments**
3. Select the new spreadsheet as the test document
4. Click **Execute**
