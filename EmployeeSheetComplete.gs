/**
 * Generalized Google Apps Script for syncing data, tracking updates,
 * incentive calculation, and onEdit workflow.
 *
 * Replace placeholder IDs and sheet names in SETTINGS as needed.
 */

const SETTINGS = {
  sourceSpreadsheetId: 'ENTER_SOURCE_SPREADSHEET_ID',
  targetSpreadsheetId: 'ENTER_TARGET_SPREADSHEET_ID',
  sourceSheetName: 'SourceData',
  targetSheetName: 'Tracking',
  courierCodesSheetName: 'CourierCodes',
  leadSheetName: 'LeadCallMsg',
  historySheetName: 'History',
  taskSheetName: 'Task Scheduller',
  incentiveSheetName: 'Incentive_Report',
  incentiveCalcSheetName: 'Incentive_Report',
  whatsappBaseUrl: 'https://web.whatsapp.com/send',
  trackingBaseUrl: 'https://t.17track.net/en#nums=',
  phoneCountryCode: '91',
  domesticCountryKeyword: 'india',
  domesticDeliveryText: 'Usual delivery time is 3–5 days',
  internationalDeliveryText: 'Usual delivery time is 15–20 days (international shipping)',
  unknownDeliveryText: 'Delivery timeline will be shared soon',
  callStatusColumn: 19,
  messageColumn: 13,
  messageStatusColumn: 24,
};

function runAllOperations() {
  Logger.log('🚀 STARTING ALL OPERATIONS');

  try {
    Logger.log('📊 Step 1: Running Full Sync...');
    runFullSync();

    Logger.log('📦 Step 2: Updating Tracking Details...');
    updateTrackingDetails();

    Logger.log('💰 Step 3: Calculating Incentives...');
    updateVrindaIncentiveBox();

    Logger.log('✅ ALL OPERATIONS COMPLETED SUCCESSFULLY!');
    SpreadsheetApp.getActiveSpreadsheet().toast('All operations completed successfully!');
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error occurred: ' + error.message);
  }
}

function updateTrackingDetails() {
  const sourceSS = SpreadsheetApp.openById(SETTINGS.sourceSpreadsheetId);
  const targetSS = SpreadsheetApp.openById(SETTINGS.targetSpreadsheetId);

  const sourceSheet = sourceSS.getSheetByName(SETTINGS.sourceSheetName) || sourceSS.getSheets()[0];
  const targetSheet = targetSS.getSheetByName(SETTINGS.targetSheetName) || targetSS.getSheets()[0];
  const courierSheet = targetSS.getSheetByName(SETTINGS.courierCodesSheetName);

  if (!sourceSheet || !targetSheet || !courierSheet) {
    SpreadsheetApp.getUi().alert('⚠️ Please verify the configured sheet names and spreadsheet IDs.');
    return;
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();
  const courierData = courierSheet.getDataRange().getValues();

  const courierMap = buildCourierMap(courierData);
  const trackingMap = buildTrackingMap(sourceData);

  for (let i = 1; i < targetData.length; i++) {
    const code = normalizeValue(targetData[i][9]);
    if (!code) {
      ensureDefaultCallStatus(targetData, i);
      continue;
    }

    const match = trackingMap.get(code);
    if (!match) {
      ensureDefaultCallStatus(targetData, i);
      continue;
    }

    const status = normalizeValue(match[27]);
    const prevStatus = normalizeValue(targetData[i][24]);
    const courierNameRaw = normalizeValue(match[25]);
    const country = normalizeValue(match[23]);

    const trackingLink = buildTrackingLink(code, courierNameRaw, courierMap);
    targetData[i][10] = trackingLink;
    targetData[i][11] = status;

    const deliveryText = buildDeliveryText(country);
    const mobile = normalizeValue(targetData[i][4]);
    const name = normalizeValue(targetData[i][2]) || 'Sir / Ma’am';
    const paymentMethod = normalizeValue(targetData[i][3]);

    if (mobile) {
      const cleanMobile = SETTINGS.phoneCountryCode + mobile.replace(/\D/g, '');
      targetData[i][13] = buildWhatsAppHyperlink(cleanMobile, buildTransitMessage(name, code, trackingLink, deliveryText), 'Send Message');

      targetData[i][14] = paymentsMessage(status, paymentMethod, name, cleanMobile);
      targetData[i][15] = returnMessage(status, name, cleanMobile);
      targetData[i][17] = feedbackMessage(status, name, cleanMobile);

      if (status !== prevStatus) {
        targetData[i][22] = buildWhatsAppHyperlink(cleanMobile, buildDynamicMessage(status, name, code, trackingLink, deliveryText), 'Send WhatsApp');
        targetData[i][23] = 'Pending';
        targetData[i][24] = status;
      }
    }

    ensureDefaultCallStatus(targetData, i);
  }

  targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
  Logger.log('✅ FINAL SYSTEM WITH DELIVERY + CALL DEFAULT LIVE');
}

function buildCourierMap(courierData) {
  const map = new Map();
  for (let i = 1; i < courierData.length; i++) {
    const name = normalizeValue(courierData[i][0]);
    const code = normalizeValue(courierData[i][1]);
    if (name && code) map.set(name, code);
  }
  return map;
}

function buildTrackingMap(sourceData) {
  const map = new Map();
  for (let i = 1; i < sourceData.length; i++) {
    const trackingCode = normalizeValue(sourceData[i][26]);
    if (trackingCode) map.set(trackingCode, sourceData[i]);
  }
  return map;
}

function buildTrackingLink(trackingCode, courierNameRaw, courierMap) {
  let trackingLink = SETTINGS.trackingBaseUrl + trackingCode;
  if (courierNameRaw) {
    const courierName = courierNameRaw.toLowerCase().trim();
    const fcCode = courierMap.get(courierName);
    if (fcCode) trackingLink += '&fc=' + fcCode;
  }
  return trackingLink;
}

function buildDeliveryText(country) {
  if (!country) return SETTINGS.unknownDeliveryText;
  return country.toLowerCase().includes(SETTINGS.domesticCountryKeyword)
    ? SETTINGS.domesticDeliveryText
    : SETTINGS.internationalDeliveryText;
}

function buildWhatsAppHyperlink(phone, text, label) {
  const url = `${SETTINGS.whatsappBaseUrl}?phone=${phone}&text=${encodeURIComponent(text)}`;
  return `=HYPERLINK("${url}","${label}")`;
}

function buildTransitMessage(name, trackingCode, trackingLink, deliveryText) {
  return `Hello ${name} 👋,\n\nYour order is in transit 🚚\n\n📦 ${deliveryText}\n\nTracking Code: ${trackingCode}\nTrack here: ${trackingLink}\n\nPlease pick up the courier call 📞 so delivery happens smoothly 😊\n\nThanks!`;
}

function paymentsMessage(status, paymentMethod, name, phone) {
  if (paymentMethod.toUpperCase() !== 'COD') return '';
  if (!status.toLowerCase().includes('deliver')) return '';

  const codMsg = `Hello ${name} 😊,\n\nYour order has been delivered 🎉\nKindly confirm COD payment 💰\n\nThank you for shopping with us 🙏`;
  return buildWhatsAppHyperlink(phone, codMsg, 'Send COD Msg');
}

function returnMessage(status, name, phone) {
  if (!status.toLowerCase().includes('return')) return '';

  const rtoMsg = `Hello ${name},\n\nYour return has been completed 🔁\nIf you need any help, feel free to contact us 😊`;
  return buildWhatsAppHyperlink(phone, rtoMsg, 'Send RTO Msg');
}

function feedbackMessage(status, name, phone) {
  if (!status.toLowerCase().includes('deliver')) return '';

  const feedbackMsg = `Hi ${name} 😊,\n\nWe hope you had a great experience with us ✨\n\nIf you would like, please share your feedback.\n\nThanks for choosing us 🙏`;
  return buildWhatsAppHyperlink(phone, feedbackMsg, 'Send Feedback');
}

function buildDynamicMessage(status, name, trackingCode, trackingLink, deliveryText) {
  const normalizedStatus = status.toLowerCase();

  if (normalizedStatus.includes('transit')) {
    return `Hello ${name} 👋,\n\nYour order is in transit 🚚\n\n📦 ${deliveryText}\n\nTracking Code: ${trackingCode}\nTrack here: ${trackingLink}\n\nPlease pick up the courier call 📞\n\nThanks!`;
  }

  if (normalizedStatus.includes('deliver')) {
    return `Hello ${name} 😊,\n\n🎉 Your order has been delivered!\n\nTracking Code: ${trackingCode}\n\nWe hope you loved it ❤️\n\nIf you would like, please share your feedback.\n\nThanks!`;
  }

  if (normalizedStatus.includes('return')) {
    return `Hello ${name},\n\nYour order return has been completed 🔁\n\nIf you need any help, feel free to contact us 😊`;
  }

  return `Hello ${name},\n\nYour order status is: ${status}\n\nTrack here:\n${trackingLink}`;
}

function ensureDefaultCallStatus(data, rowIndex) {
  if (!data[rowIndex][18] || data[rowIndex][18] === '') {
    data[rowIndex][18] = 'Pending';
  }
}

function normalizeValue(value) {
  return value == null ? '' : String(value).trim();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < 2) return;

  if (['Sheet1', 'Tracking'].includes(sheet.getName())) {
    if (col !== SETTINGS.callStatusColumn && col !== SETTINGS.messageColumn && col !== SETTINGS.messageStatusColumn) return;
    handleTrackingSheetEdit(e, row, col);
    return;
  }

  if (sheet.getName() === SETTINGS.leadSheetName) {
    handleLeadSheetEdit(e, ss);
  }
}

function handleTrackingSheetEdit(e, row, col) {
  const sheet = e.source.getActiveSheet();
  const cell = sheet.getRange(row, col);
  const newValue = normalizeValue(e.value);
  const history = PropertiesService.getDocumentProperties();

  if (col === SETTINGS.callStatusColumn) {
    const callKey = 'call_row_' + row;
    if (newValue !== 'Completed' && history.getProperty(callKey) === 'Completed') {
      cell.setValue('Completed');
      SpreadsheetApp.getUi().alert('❌ Locked: Already set to Complete');
      return;
    }
    if (newValue === 'Completed') history.setProperty(callKey, 'Completed');
    setNormalizedStatus(cell, newValue, ['Pending', 'Unreachable', 'Completed']);
  }

  if (col === SETTINGS.messageColumn) {
    const msgKey = 'msg_row_' + row;
    if (newValue !== 'Sent' && history.getProperty(msgKey) === 'Sent') {
      cell.setValue('Sent');
      SpreadsheetApp.getUi().alert('❌ Locked: Already set to Sent');
      return;
    }
    if (newValue === 'Sent') history.setProperty(msgKey, 'Sent');
    setNormalizedStatus(cell, newValue, ['Pending', 'Sent']);
  }

  if (col === SETTINGS.messageStatusColumn) {
    const msgKey = 'wa_msg_' + row;
    if (newValue !== 'Sent' && history.getProperty(msgKey) === 'Sent') {
      cell.setValue('Sent');
      SpreadsheetApp.getUi().alert('❌ Already Sent. Locked.');
      return;
    }
    if (newValue === 'Sent') history.setProperty(msgKey, 'Sent');
    setNormalizedStatus(cell, newValue, ['Pending', 'Sent']);
  }
}

function setNormalizedStatus(cell, value, validValues) {
  const normalized = normalizeValue(value).toLowerCase();
  validValues.forEach((option) => {
    if (normalized === option.toLowerCase()) cell.setValue(option);
  });
}

function handleLeadSheetEdit(e, ss) {
  const sheet = ss.getSheetByName(SETTINGS.leadSheetName);
  const historySheet = ss.getSheetByName(SETTINGS.historySheetName);
  if (!sheet || !historySheet) {
    SpreadsheetApp.getUi().alert('⚠️ Lead or History sheet not found.');
    return;
  }

  const START_ROW = 3;
  const STATUS_COL = 10;
  const PHONE_COL = 5;
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col !== STATUS_COL || row < START_ROW) return;

  const newValue = normalizeValue(e.value);
  const oldValue = normalizeValue(e.oldValue);

  const lastCol = sheet.getLastColumn();
  const rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const formulas = sheet.getRange(row, 1, 1, lastCol).getFormulas()[0];
  const phone = normalizeValue(rowData[PHONE_COL - 1]);

  if (newValue === 'Completed') {
    for (let c = 1; c <= 9; c++) {
      if (!rowData[c - 1]) {
        e.range.setValue(oldValue || '');
        SpreadsheetApp.getUi().alert('❌ Fill all fields from A to I before marking Completed.');
        return;
      }
    }

    const phoneList = historySheet.getRange(START_ROW, PHONE_COL, Math.max(0, historySheet.getLastRow() - START_ROW + 1), 1).getValues().flat().map(normalizeValue);
    if (phoneList.includes(phone)) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Already saved in History');
      return;
    }

    const targetRow = historySheet.getLastRow() + 1;
    historySheet.getRange(targetRow, 1, 1, lastCol).setValues([rowData]);
    formulas.forEach((formula, index) => {
      if (formula) historySheet.getRange(targetRow, index + 1).setFormula(formula);
    });
    historySheet.getRange(targetRow, 1, 1, lastCol).clearDataValidations();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Saved to History');
    return;
  }

  if (oldValue === 'Completed' && newValue !== 'Completed') {
    const phoneList = historySheet.getRange(START_ROW, PHONE_COL, Math.max(0, historySheet.getLastRow() - START_ROW + 1), 1).getValues().flat().map(normalizeValue);
    for (let i = phoneList.length - 1; i >= 0; i--) {
      if (phoneList[i] === phone) {
        historySheet.deleteRow(i + START_ROW);
        SpreadsheetApp.getActiveSpreadsheet().toast('↩ Removed from History');
        return;
      }
    }
  }
}

function updateVrindaIncentiveBox(sheetName = SETTINGS.incentiveCalcSheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ Incentive sheet not found: ' + sheetName);
    return;
  }

  const startRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const values = sheet.getRange(startRow, 9, lastRow - startRow + 1, 1).getValues();
  let totalSales = 0;
  values.forEach((row) => {
    totalSales += Number(row[0]) || 0;
  });

  let earning1 = 0;
  let earning2 = 0;
  let bonus = 0;

  if (totalSales <= 50000) {
    earning1 = totalSales * 0.025;
  } else {
    earning1 = 50000 * 0.025;
    earning2 = (totalSales - 50000) * 0.03;
  }

  if (totalSales > 150000) bonus = 500;
  const totalPayable = earning1 + earning2 + bonus;

  sheet.getRange('AC1').setValue('Incentive Summary');
  sheet.getRange('AA2').setValue(totalSales);

  sheet.getRange('AA5').setValue(1);
  sheet.getRange('AB5').setValue(Math.min(totalSales, 50000));
  sheet.getRange('AC5').setValue('2.50%');
  sheet.getRange('AD5').setValue(Math.round(earning1));

  sheet.getRange('AA6').setValue(2);
  sheet.getRange('AB6').setValue(totalSales > 50000 ? totalSales - 50000 : 0);
  sheet.getRange('AC6').setValue('3%');
  sheet.getRange('AD6').setValue(Math.round(earning2));

  sheet.getRange('AA7').setValue(3);
  sheet.getRange('AB7').setValue('Bonus');
  sheet.getRange('AD7').setValue(bonus);

  sheet.getRange('AC8').setValue('Total Payable');
  sheet.getRange('AD8').setValue(Math.round(totalPayable));
}

function runFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(SETTINGS.taskSheetName);
  const leadSheet = ss.getSheetByName(SETTINGS.leadSheetName);
  const incentiveSheet = ss.getSheetByName(SETTINGS.incentiveSheetName);

  if (!taskSheet || !leadSheet || !incentiveSheet) {
    SpreadsheetApp.getUi().alert('⚠️ Missing sheet for full sync.');
    return;
  }

  Logger.log('🚀 MASTER SYNC STARTED');

  const finalData = [];
  const columnGLinks = [];

  const leadLastRow = leadSheet.getLastRow();
  if (leadLastRow >= 3) {
    const leadData = leadSheet.getRange(3, 1, leadLastRow - 2, 10).getValues();
    leadData.forEach((row) => {
      const date = row[0] || new Date();
      const name = normalizeValue(row[3]);
      let phone = normalizeValue(row[4]).replace(/\D/g, '');
      if (phone.length === 10) phone = SETTINGS.phoneCountryCode + phone;
      if (!phone) return;

      const status = normalizeValue(row[5]).toLowerCase() === 'sent' ? 'Sent' : 'Pending';
      finalData.push([date, SETTINGS.leadSheetName, name, phone, status, '']);
      columnGLinks.push([`=HYPERLINK("${SETTINGS.whatsappBaseUrl}?phone=${phone}","Whatsapp")`]);
    });
    Logger.log('✅ LeadCallMsg Loaded');
  }

  const incLastRow = incentiveSheet.getLastRow();
  if (incLastRow >= 2) {
    const incValues = incentiveSheet.getRange(2, 1, incLastRow - 1, 25).getValues();
    const incFormulas = incentiveSheet.getRange(2, 1, incLastRow - 1, 25).getFormulas();

    incValues.forEach((row, index) => {
      const date = row[0];
      const name = normalizeValue(row[2]);
      let phone = normalizeValue(row[4]).replace(/\D/g, '');
      if (phone.length === 10) phone = SETTINGS.phoneCountryCode + phone;
      if (!phone) return;

      const status = normalizeValue(row[23]) || 'Pending';
      const originalFormula = normalizeValue(incFormulas[index][22]);
      const originalValue = normalizeValue(row[22]);
      const finalMessage = originalFormula || originalValue;
      finalData.push([date, SETTINGS.incentiveSheetName, name, phone, status, finalMessage]);

      let newFormula = '';
      if (originalFormula) {
        const match = originalFormula.match(/"(https?:\/\/[^\"]+)"/);
        if (match && match[1]) {
          newFormula = `=HYPERLINK("${match[1]}","Whatsapp")`;
        }
      }
      columnGLinks.push([newFormula]);
    });
    Logger.log('✅ Incentive_Report Loaded');
  }

  if (taskSheet.getLastRow() > 1) {
    taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 7).clearContent();
  }

  if (finalData.length > 0) {
    taskSheet.getRange(2, 1, finalData.length, 6).setValues(finalData);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Sent'], true).build();
    taskSheet.getRange(2, 5, finalData.length, 1).setDataValidation(rule);
    SpreadsheetApp.flush();
    taskSheet.getRange(2, 7, columnGLinks.length, 1).setFormulas(columnGLinks);
    SpreadsheetApp.flush();
    Logger.log('🎉 FINAL SYNC COMPLETE');
  } else {
    Logger.log('❌ No data found');
  }
}
