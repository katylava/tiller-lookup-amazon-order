function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Amazon Lookup')
    .addItem('Lookup selected amount', 'lookupAmazonOrder')
    .addItem('Lookup all uncategorized Amazon transactions', 'lookupAllAmazon')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function findAmazonMatches(amazonData, headers, searchAmount) {
  var totalCol = headers.indexOf('total');
  var dateCol = headers.indexOf('date');
  var itemsCol = headers.indexOf('items');

  var matches = [];
  for (var i = 1; i < amazonData.length; i++) {
    var rowTotal = parseFloat(amazonData[i][totalCol]);
    if (!isNaN(rowTotal) && Math.abs(Math.abs(rowTotal) - searchAmount) < 0.005) {
      var date = amazonData[i][dateCol];
      if (date instanceof Date) {
        date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      matches.push({
        date: date,
        items: amazonData[i][itemsCol]
      });
    }
  }
  return matches;
}

function formatNoteFromMatches(matches) {
  var noteLines = [];
  for (var j = 0; j < matches.length; j++) {
    var m = matches[j];
    noteLines.push('Date: ' + m.date + '\nItems: ' + m.items);
  }
  return noteLines.join('\n---\n');
}

function lookupAmazonOrder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = sheet.getActiveCell();
  var cellValue = cell.getValue();

  if (cellValue === '' || cellValue === null) {
    SpreadsheetApp.getUi().alert('Selected cell is empty.');
    return;
  }

  var searchAmount = Math.abs(parseFloat(cellValue));
  if (isNaN(searchAmount)) {
    SpreadsheetApp.getUi().alert('Selected cell does not contain a numeric value.');
    return;
  }

  var amazonSheet = sheet.getSheetByName('tmp-amazon');
  if (!amazonSheet) {
    SpreadsheetApp.getUi().alert('Sheet "tmp-amazon" not found.');
    return;
  }

  var data = amazonSheet.getDataRange().getValues();
  var headers = data[0];

  if (headers.indexOf('total') === -1) {
    SpreadsheetApp.getUi().alert('Column "total" not found in "tmp-amazon" sheet.');
    return;
  }
  if (headers.indexOf('date') === -1) {
    SpreadsheetApp.getUi().alert('Column "date" not found in "tmp-amazon" sheet.');
    return;
  }
  if (headers.indexOf('items') === -1) {
    SpreadsheetApp.getUi().alert('Column "items" not found in "tmp-amazon" sheet.');
    return;
  }

  var matches = findAmazonMatches(data, headers, searchAmount);

  if (matches.length === 0) {
    SpreadsheetApp.getUi().alert('No matching order found for $' + searchAmount.toFixed(2));
    return;
  }

  cell.setNote(formatNoteFromMatches(matches));

  var htmlLines = [];
  for (var j = 0; j < matches.length; j++) {
    var m = matches[j];
    htmlLines.push('<p><strong>Date:</strong> ' + m.date + '<br><strong>Items:</strong> ' + m.items + '</p>');
  }

  var html = HtmlService
    .createHtmlOutput(
      '<div style="font-family: Arial, sans-serif; padding: 8px;">' +
      '<h3>Amazon Order' + (matches.length > 1 ? 's' : '') + ' for $' + searchAmount.toFixed(2) + '</h3>' +
      htmlLines.join('<hr>') +
      '</div>'
    )
    .setWidth(400)
    .setHeight(200 + (matches.length - 1) * 80);

  SpreadsheetApp.getUi().showModalDialog(html, 'Amazon Order Lookup');
}

function lookupAllAmazon() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var txnSheet = ss.getSheetByName('Transactions');
  if (!txnSheet) {
    ui.alert('Sheet "Transactions" not found.');
    return;
  }

  var amazonSheet = ss.getSheetByName('tmp-amazon');
  if (!amazonSheet) {
    ui.alert('Sheet "tmp-amazon" not found.');
    return;
  }

  var amazonData = amazonSheet.getDataRange().getValues();
  var amazonHeaders = amazonData[0];

  if (amazonHeaders.indexOf('total') === -1 || amazonHeaders.indexOf('date') === -1 || amazonHeaders.indexOf('items') === -1) {
    ui.alert('Required columns (total, date, items) not found in "tmp-amazon" sheet.');
    return;
  }

  var txnHeaders = txnSheet.getRange(1, 1, 1, txnSheet.getLastColumn()).getValues()[0];
  var descCol = txnHeaders.indexOf('Description');
  var categoryCol = txnHeaders.indexOf('Category');
  var amountCol = txnHeaders.indexOf('Amount');

  if (descCol === -1) {
    ui.alert('Column "Description" not found in "Transactions" sheet.');
    return;
  }
  if (categoryCol === -1) {
    ui.alert('Column "Category" not found in "Transactions" sheet.');
    return;
  }
  if (amountCol === -1) {
    ui.alert('Column "Amount" not found in "Transactions" sheet.');
    return;
  }

  // Use TextFinder on the Description column to find cells starting with "Amazon"
  var descRange = txnSheet.getRange(2, descCol + 1, txnSheet.getLastRow() - 1, 1);
  var finder = descRange.createTextFinder('^Amazon').useRegularExpression(true);

  var matched = 0;
  var noMatch = 0;
  var skipped = 0;
  var processed = 0;

  var found = finder.findNext();
  while (found) {
    var row = found.getRow();
    processed++;

    // Check if this row has a category — if so, stop entirely
    var categoryCell = txnSheet.getRange(row, categoryCol + 1);
    if (categoryCell.getValue() !== '' && categoryCell.getValue() !== null) {
      break;
    }

    // Check if the Amount cell already has a note — if so, skip
    var amountCell = txnSheet.getRange(row, amountCol + 1);
    if (amountCell.getNote() !== '') {
      skipped++;
      found = finder.findNext();
      continue;
    }

    // Look up the absolute amount
    var amountValue = Math.abs(parseFloat(amountCell.getValue()));
    if (isNaN(amountValue)) {
      found = finder.findNext();
      continue;
    }

    var matches = findAmazonMatches(amazonData, amazonHeaders, amountValue);
    if (matches.length > 0) {
      amountCell.setNote(formatNoteFromMatches(matches));
      matched++;
    } else {
      noMatch++;
    }

    found = finder.findNext();
  }

  ui.alert(
    'Bulk Amazon Lookup Complete',
    'Processed ' + processed + ' transactions.\n' +
    matched + ' matched.\n' +
    noMatch + ' had no match.\n' +
    skipped + ' skipped (already had notes).',
    ui.ButtonSet.OK
  );
}
