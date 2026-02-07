# Lookup Amazon Order From Cell Value

A Google Apps Script add-on for Tiller spreadsheets that looks up Amazon order
details from a "tmp-amazon" sheet (columns: total, date, items).

## Features

- **Lookup selected amount** — Select a cell, and the add-on searches for the
  absolute dollar value in "tmp-amazon". Matching order date and items are shown
  in a pop-up dialog and added as a note to the cell.

- **Lookup all uncategorized Amazon transactions** — Bulk processes the
  "Transactions" sheet. Finds all rows with a Description starting with
  "Amazon", skips rows that already have notes, and stops when it hits a row
  with a Category value. Adds order notes to each matching Amount cell and shows
  a summary when done.

## Deployment

1. `clasp push --force`
2. Open the Apps Script project and execute a new **test deployment**
   (reloading the spreadsheet alone is not enough)

