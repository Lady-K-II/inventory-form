// ============================================================
// LADY K II INVENTORY - Apps Script backend, v2
// Paste the entire Sheet ID between /d/ and /edit into SPREADSHEET_ID below
// before pushing. Then run setupInventorySheet() once from the Apps Script
// editor (Run menu, pick setupInventorySheet), THEN redeploy.
// ============================================================

var SPREADSHEET_ID = '1uV-A0lJKZY-k0koHOqfjcxZoqvPU9qnrCxMNCqBbicM';
var SHEET_NAME = 'Inventory';
var HEADER_ROW = 4;
var FIRST_DATA_ROW = 5;

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Tools')
    .addItem('Get Add-Item Form Link', 'showFormLink')
    .addToUi();
}

function showFormLink() {
  var url = ScriptApp.getService().getUrl();
  var msg;
  if (!url) {
    msg = '<p style="font-family:Arial">Not deployed yet.</p>';
  } else {
    msg = '<p style="font-family:Arial">Open or bookmark this on your phone:</p>'
        + '<p><a href="' + url + '" target="_blank">' + url + '</a></p>';
  }
  var html = HtmlService.createHtmlOutput(msg).setWidth(420).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add-Item Form Link');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  if (action === 'getFormData') {
    try {
      return jsonResponse(getFormData());
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }
  return HtmlService.createHtmlOutputFromFile('Form')
      .setTitle('Lady K II - Add Inventory Item')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'submitItem') {
      return jsonResponse(submitItem(body.item));
    }
    if (body.action === 'uploadPhoto') {
      return jsonResponse(uploadPhoto(body.base64Data, body.mimeType, body.filename));
    }
    if (body.action === 'analyzePhoto') {
      return jsonResponse(analyzePhoto(body.base64Data, body.mimeType));
    }
    return jsonResponse({ error: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function analyzePhoto(base64Data, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { error: 'No Gemini API key configured in Script Properties.' };
  }

  var categoryList = ['Engine', 'Generator', 'Electrical', 'HVAC', 'Watermaker', 'Plumbing',
    'Deck/Crane', 'Safety Equipment', 'Tender/Outboard', 'Galley',
    'Navigation/Electronics', 'Consumables', 'Other'].join(', ');

  var prompt = 'You are helping catalog a spare part or piece of equipment on a boat. ' +
    'Look at this photo and identify what you can. Respond with ONLY a JSON object, ' +
    'no markdown formatting, no code fences, no explanation, in exactly this shape: ' +
    '{"manufacturer":"","partNumber":"","partName":"","description":"","category":"","itemType":""}. ' +
    'manufacturer: the brand name if visible, else empty string. ' +
    'partNumber: any part, model, or article number visible, else empty string. ' +
    'partName: a short 2-4 word name for the item, e.g. "Oil Filter" or "Fuel Filter". ' +
    'description: one brief sentence describing the item and what it fits, using anything readable on the label. ' +
    'category: pick exactly one from this list, matching spelling exactly: ' + categoryList + '. ' +
    'itemType: either "Consumable" (used up and replaced, like a filter, oil, or fluid) or "Tool" (durable equipment kept and reused, like a drill or wrench). ' +
    'If you genuinely cannot tell something, use an empty string rather than guessing.';

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent';
  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: mimeType, data: base64Data } }
      ]
    }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.error) {
    return { error: json.error.message || 'Gemini API error' };
  }
  var text = json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts && json.candidates[0].content.parts[0].text;
  if (!text) {
    return { error: 'No response from Gemini.' };
  }
  var cleaned = text.replace(/```json/gi, '').replace(/```/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e2) {
    return { error: 'Could not parse Gemini response: ' + cleaned.substring(0, 200) };
  }
}

function setupInventorySheet() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    throw new Error('An "Inventory" tab already exists. Delete it first if you want a clean rebuild, or skip this step if it already has your data.');
  }
  sheet = ss.insertSheet(SHEET_NAME);

  sheet.getRange('A1').setValue('LADY K II - PARTS & CONSUMABLES INVENTORY')
    .setFontWeight('bold').setFontSize(14).setFontColor('#1F3864').setFontFamily('Arial');
  sheet.getRange('A2').setValue('Add new items using the Add-Item form, Inventory Tools menu')
    .setFontStyle('italic').setFontSize(9).setFontColor('#595959').setFontFamily('Arial');

  var headers = ['Photo', 'Photo Filename', 'Part Name', 'System / Machinery (used on)',
    'Category', 'Description', 'Location Onboard', 'Manufacturer / Part No.',
    'Qty On Hand', 'Min Stock Required', 'Reorder Status', 'Unit Cost', 'Total Value',
    'Supplier / Source', 'Purchase Date', 'Expiration Date', 'Days to Expiry',
    'Expiry Status', 'Critical Spare (Y/N)', 'Notes'];

  var headerRange = sheet.getRange(HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F3864')
    .setFontFamily('Arial').setFontSize(10).setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(HEADER_ROW, 40);

  var widths = [110, 220, 150, 150, 120, 200, 150, 160, 80, 90, 95, 80, 90, 150, 95, 95, 80, 100, 90, 200];
  for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);

  sheet.setFrozenRows(HEADER_ROW);

  var categories = ['Engine', 'Generator', 'Electrical', 'HVAC', 'Watermaker', 'Plumbing',
    'Deck/Crane', 'Safety Equipment', 'Tender/Outboard', 'Galley',
    'Navigation/Electronics', 'Consumables', 'Other'];
  var catRule = SpreadsheetApp.newDataValidation().requireValueInList(categories, true).setAllowInvalid(true).build();
  var critRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).setAllowInvalid(true).build();
  var LAST_ROW = FIRST_DATA_ROW + 199;
  sheet.getRange(FIRST_DATA_ROW, 5, LAST_ROW - FIRST_DATA_ROW + 1, 1).setDataValidation(catRule);
  sheet.getRange(FIRST_DATA_ROW, 19, LAST_ROW - FIRST_DATA_ROW + 1, 1).setDataValidation(critRule);

  var rules = sheet.getConditionalFormatRules();
  var kRange = sheet.getRange(FIRST_DATA_ROW, 11, LAST_ROW - FIRST_DATA_ROW + 1, 1);
  var rRange = sheet.getRange(FIRST_DATA_ROW, 18, LAST_ROW - FIRST_DATA_ROW + 1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('REORDER').setBackground('#F8CBCB').setRanges([kRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('EXPIRING SOON').setBackground('#FFF2B2').setRanges([rRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('EXPIRED').setBackground('#E06666').setRanges([rRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.getRange(FIRST_DATA_ROW, 15, LAST_ROW - FIRST_DATA_ROW + 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(FIRST_DATA_ROW, 16, LAST_ROW - FIRST_DATA_ROW + 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(FIRST_DATA_ROW, 12, LAST_ROW - FIRST_DATA_ROW + 1, 1).setNumberFormat('"EUR "#,##0.00;("EUR "#,##0.00);"-"');
  sheet.getRange(FIRST_DATA_ROW, 13, LAST_ROW - FIRST_DATA_ROW + 1, 1).setNumberFormat('"EUR "#,##0.00;("EUR "#,##0.00);"-"');

  return 'Inventory tab created. You can now use the Add-Item form.';
}

function updateCategoryValidation() {
  var sheet = getSS().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Inventory tab not found. Run setupInventorySheet() first.');
  }
  var categories = ['Engine', 'Generator', 'Electrical', 'HVAC', 'Watermaker', 'Plumbing',
    'Deck/Crane', 'Safety Equipment', 'Tender/Outboard', 'Galley',
    'Navigation/Electronics', 'Consumables', 'Other'];
  var catRule = SpreadsheetApp.newDataValidation().requireValueInList(categories, true).setAllowInvalid(true).build();
  var LAST_ROW = FIRST_DATA_ROW + 199;
  sheet.getRange(FIRST_DATA_ROW, 5, LAST_ROW - FIRST_DATA_ROW + 1, 1).setDataValidation(catRule);
  return 'Category dropdown updated.';
}

function getFormData() {
  var sheet = getSS().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Inventory tab not found. Run setupInventorySheet() first from the Apps Script editor.');
  }
  var lastRow = sheet.getLastRow();
  var systems = {}, locations = {};
  if (lastRow >= FIRST_DATA_ROW) {
    var data = sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 20).getValues();
    data.forEach(function (row) {
      var sys = String(row[3] || '').trim();
      var loc = String(row[6] || '').trim();
      if (sys) systems[sys] = true;
      if (loc) locations[loc] = true;
    });
  }
  return {
    systems: Object.keys(systems).sort(),
    locations: Object.keys(locations).sort(),
    categories: ['Engine', 'Generator', 'Electrical', 'HVAC', 'Watermaker', 'Plumbing',
      'Deck/Crane', 'Safety Equipment', 'Tender/Outboard', 'Galley',
      'Navigation/Electronics', 'Consumables', 'Other']
  };
}

function uploadPhoto(base64Data, mimeType, filename) {
  var folder = getOrCreatePhotoFolder();
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    url: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    name: filename
  };
}

function getOrCreatePhotoFolder() {
  var name = 'Lady K II Inventory Photos';
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function submitItem(item) {
  var sheet = getSS().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Inventory tab not found. Run setupInventorySheet() first from the Apps Script editor.');
  }
  var newRow = Math.max(sheet.getLastRow() + 1, FIRST_DATA_ROW);

  var rowValues = [
    '', item.filename || '', item.partName || '', item.system || '', item.category || '',
    item.description || '', item.location || '', item.manufacturer || '',
    item.qty === '' ? '' : Number(item.qty),
    item.minStock === '' ? '' : Number(item.minStock),
    '', item.unitCost === '' ? '' : Number(item.unitCost), '',
    item.supplier || '', item.purchaseDate || '', item.expirationDate || '',
    '', '', item.critical || '', item.notes || ''
  ];

  sheet.getRange(newRow, 1, 1, rowValues.length).setValues([rowValues]);

  if (item.photoUrl) {
    sheet.getRange(newRow, 1).setFormula('=IMAGE("' + item.photoUrl + '")');
  }
  sheet.getRange(newRow, 11).setFormula(
    '=IF(I' + newRow + '="","",IF(I' + newRow + '<J' + newRow + ',"REORDER","OK"))');
  sheet.getRange(newRow, 13).setFormula(
    '=IF(OR(I' + newRow + '="",L' + newRow + '=""),"",I' + newRow + '*L' + newRow + ')');
  sheet.getRange(newRow, 17).setFormula(
    '=IF(P' + newRow + '="","",P' + newRow + '-TODAY())');
  sheet.getRange(newRow, 18).setFormula(
    '=IF(P' + newRow + '="","N/A",IF(Q' + newRow + '<0,"EXPIRED",' +
    'IF(Q' + newRow + '<=30,"EXPIRING SOON","OK")))');

  return { row: newRow };
}
