/**
 * كود ربط نظام "حساب البيت" بـ Google Sheets
 * مع دعم:
 * - توقعات الشهر القادم (المرتب والإيرادات المتوقعة)
 * - سجل التعديلات مع خاصية CHECKBOX للحذف والتراجع من الشيت مباشرة
 * - خصم إيجار بدر من الحساب البنكي لحساب فلوسك
 */

var SHEETS = {
  exp: 'المصاريف',
  badrE: 'مصاريف بدر',
  badrI: 'إيرادات بدر',
  expected: 'إيرادات متوقعة للشهر القادم',
  expectedExp: 'مصاريف متوقعة للشهر القادم',
  settings: 'الأرصدة والدهب',
  logs: 'سجل التعديلات'
};

// إنشاء قائمة مخصصة في شريط أدوات Google Sheets
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 حساب البيت')
    .addItem('☑️ تفعيل مربعات الاختيار (Checkboxes) لجميع التعديلات', 'setupLogSheetWithCheckboxes')
    .addItem('↩️ التراجع وحذف التعديل المحدد', 'undoSelectedRow')
    .addItem('↩️ التراجع وحذف آخر تعديل تم', 'undoLatestEdit')
    .addSeparator()
    .addItem('🧹 مسح سجل التعديلات بالكامل', 'clearAllLogs')
    .addToUi();
}

// مراقبة النقر على الـ CHECKBOX في صفحة "سجل التعديلات"
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() === SHEETS.logs) {
    var col = e.range.getColumn();
    var row = e.range.getRow();
    var val = e.value;
    if (col === 1 && row > 1 && (val === 'TRUE' || val === true || val === 'true')) {
      undoAndDeleteLogAtRow(sheet, row);
    }
  }
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'get';
  var result = {};
  
  try {
    if (action === 'get') {
      result = { success: true, data: getAllData() };
    } else {
      result = { success: false, error: 'أمر غير معروف' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result = {};
  try {
    var contents = e.postData ? e.postData.contents : '{}';
    var req = JSON.parse(contents);
    var action = req.action;
    
    if (action === 'get') {
      result = { success: true, data: getAllData() };
    } else if (action === 'uploadAll') {
      uploadAllData(req.data);
      setupLogSheetWithCheckboxes();
      result = { success: true, message: 'تم رفع كافة البيانات وتحديث الشيت بنجاح' };
    } else if (action === 'add') {
      addRecord(req.list, req.item);
      result = { success: true, message: 'تمت الإضافة بنجاح' };
    } else if (action === 'update') {
      updateRecord(req.list, req.item, req.oldItem);
      result = { success: true, message: 'تم التعديل وتسجيله بنجاح' };
    } else if (action === 'delete') {
      deleteRecord(req.list, req.id, req.oldItem);
      result = { success: true, message: 'تم الحذف وتسجيله بنجاح' };
    } else if (action === 'setSettings') {
      saveSettings(req.settings, req.oldSettings);
      result = { success: true, message: 'تم تحديث الأرصدة وتسجيلها بنجاح' };
    } else if (action === 'deleteLog') {
      deleteLog(req.id);
      result = { success: true, message: 'تم حذف هذا التعديل من الشيت' };
    } else if (action === 'clearLogs') {
      clearAllLogs();
      result = { success: true, message: 'تم مسح سجل التعديلات بالكامل' };
    } else {
      result = { success: false, error: 'أمر غير مدعوم: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateTime(d) {
  var dt = d ? (d instanceof Date ? d : new Date(d)) : new Date();
  var yr = dt.getFullYear();
  var mo = ('0' + (dt.getMonth() + 1)).slice(-2);
  var da = ('0' + dt.getDate()).slice(-2);
  var hr = dt.getHours();
  var mi = ('0' + dt.getMinutes()).slice(-2);
  var sc = ('0' + dt.getSeconds()).slice(-2);
  var ampm = hr >= 12 ? 'م' : 'ص';
  hr = hr % 12;
  hr = hr ? hr : 12;
  var hrStr = ('0' + hr).slice(-2);
  return yr + '-' + mo + '-' + da + ' ' + hrStr + ':' + mi + ':' + sc + ' ' + ampm;
}

function getOrCreateSheet(sheetName, headers, headerColor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var color = headerColor || '#0f766e';
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground(color);
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function setupLogSheetWithCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.logs);
  var headers = [
    '☑️ حذف وتراجع',
    'اسم البند',
    'القسم',
    'القيمة القديمة',
    'القيمة بعد التعديل',
    'تاريخ ووقت التعديل',
    'نوع العملية',
    'ID البند',
    'التفاصيل'
  ];
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.logs);
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight('bold');
    hRange.setBackground('#6b21a8');
    hRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return;
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var hRange2 = sheet.getRange(1, 1, 1, headers.length);
  hRange2.setFontWeight('bold');
  hRange2.setBackground('#6b21a8');
  hRange2.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    range.insertCheckboxes();
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('تم تجهيز الـ Checkboxes لجميع التعديلات بنجاح!', '✅ حساب البيت', 5);
}

function addLog(actionType, section, itemId, itemName, oldVal, newVal, details) {
  var headers = [
    '☑️ حذف وتراجع',
    'اسم البند',
    'القسم',
    'القيمة القديمة',
    'القيمة بعد التعديل',
    'تاريخ ووقت التعديل',
    'نوع العملية',
    'ID البند',
    'التفاصيل'
  ];
  var sheet = getOrCreateSheet(SHEETS.logs, headers, '#6b21a8');
  var timeStr = formatDateTime(new Date());
  
  sheet.appendRow([
    false,
    itemName || '',
    section || '',
    oldVal !== undefined && oldVal !== null ? oldVal : '',
    newVal !== undefined && newVal !== null ? newVal : '',
    timeStr,
    actionType || '',
    String(itemId || ''),
    details || ''
  ]);
  
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).insertCheckboxes();
}

function undoAndDeleteLogAtRow(sheet, row) {
  var numCols = Math.max(9, sheet.getLastColumn());
  var data = sheet.getRange(row, 1, 1, numCols).getValues()[0];
  
  var itemName = String(data[1] || '').trim();
  var section = String(data[2] || '').trim();
  var oldVal = data[3];
  var newVal = data[4];
  var actionType = String(data[6] || '').trim();
  var itemId = String(data[7] || '').trim();

  if (!section && data[4]) {
    if (String(data[4]).indexOf('بيت') !== -1 || String(data[4]).indexOf('بدر') !== -1 || String(data[4]).indexOf('أرصدة') !== -1 || String(data[4]).indexOf('توقعات') !== -1) {
      section = String(data[4]).trim();
      itemName = String(data[5] || '').trim();
      oldVal = data[6];
      itemId = String(data[1] || '').trim();
      actionType = String(data[3] || '').trim();
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var success = false;
  var msg = '';
  var timeStr = formatDateTime(new Date());

  var targetSheetName = '';
  if (section.indexOf('مصاريف البيت') !== -1 || section.indexOf('المصاريف') !== -1) targetSheetName = SHEETS.exp;
  else if (section.indexOf('مصاريف بدر') !== -1) targetSheetName = SHEETS.badrE;
  else if (section.indexOf('إيرادات بدر') !== -1) targetSheetName = SHEETS.badrI;
  else if (section.indexOf('مصاريف متوقعة') !== -1) targetSheetName = SHEETS.expectedExp;
  else if (section.indexOf('توقعات') !== -1 || section.indexOf('إيرادات متوقعة') !== -1) targetSheetName = SHEETS.expected;

  if (targetSheetName) {
    var tSheet = ss.getSheetByName(targetSheetName) || (targetSheetName === SHEETS.expected ? ss.getSheetByName('توقعات الشهر القادم') : null);
    if (tSheet) {
      var tRows = tSheet.getDataRange().getValues();
      var foundRow = -1;

      if (itemId) {
        for (var i = 1; i < tRows.length; i++) {
          if (String(tRows[i][0]).trim() === itemId) {
            foundRow = i + 1;
            break;
          }
        }
      }

      if (foundRow === -1 && itemName) {
        for (var k = 1; k < tRows.length; k++) {
          if (String(tRows[k][1]).trim().toLowerCase() === itemName.toLowerCase()) {
            foundRow = k + 1;
            break;
          }
        }
      }

      if (foundRow !== -1) {
        tSheet.getRange(foundRow, 2).setValue(itemName);
        tSheet.getRange(foundRow, 3).setValue(parseFloat(oldVal) || 0);
        var isDated = (targetSheetName === SHEETS.exp);
        if (isDated) tSheet.getRange(foundRow, 5).setValue(timeStr + ' (مسترجع)');
        else tSheet.getRange(foundRow, 4).setValue(timeStr + ' (مسترجع)');
        success = true;
        msg = 'تم حذف التعديل واسترجاع قيمة "' + itemName + '" الأصلية (' + oldVal + ' ج) بنجاح!';
      } else {
        if (actionType.indexOf('حذف') !== -1 && itemName) {
          var isDated2 = (targetSheetName === SHEETS.exp);
          var newRow = isDated2 
            ? [itemId || Date.now(), itemName, parseFloat(oldVal) || 0, new Date().toISOString().slice(0, 10), timeStr + ' (مسترجع)']
            : [itemId || Date.now(), itemName, parseFloat(oldVal) || 0, timeStr + ' (مسترجع)'];
          tSheet.appendRow(newRow);
          success = true;
          msg = 'تم التراجع وإعادة إدراج البند المحذوف "' + itemName + '" بمبلغ (' + oldVal + ' ج)!';
        }
      }
    }
  } else if (section.indexOf('أرصدة') !== -1 || section.indexOf('الدهب') !== -1 || itemName.indexOf('بنك') !== -1 || itemName.indexOf('بيت') !== -1 || itemName.indexOf('أختي') !== -1 || itemName.indexOf('عيار') !== -1) {
    var setSheet = ss.getSheetByName(SHEETS.settings);
    if (setSheet) {
      var sRows = setSheet.getDataRange().getValues();
      for (var j = 1; j < sRows.length; j++) {
        var key = String(sRows[j][0]);
        var desc = String(sRows[j][1]);
        if (desc.indexOf(itemName) !== -1 || key.indexOf(itemName) !== -1 || itemName.indexOf(desc) !== -1) {
          setSheet.getRange(j + 1, 3).setValue(parseFloat(oldVal) || 0);
          setSheet.getRange(j + 1, 4).setValue(timeStr + ' (مسترجع)');
          success = true;
          msg = 'تمت استعادة ' + desc + ' إلى (' + oldVal + ')';
          break;
        }
      }
    }
  }

  if (success) {
    sheet.deleteRow(row);
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '✅ تم الحذف والاسترجاع', 5);
  } else {
    sheet.getRange(row, 1).setValue(false);
    SpreadsheetApp.getUi().alert('⚠️ تعذر العثور على البند "' + itemName + '" في جدول ' + section + '. تأكد أن البند موجود في الجدول.');
  }
}

function undoSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== SHEETS.logs) {
    SpreadsheetApp.getUi().alert('يرجى فتح صفحة "سجل التعديلات" وتحديد سطر التعديل أولاً.');
    return;
  }
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('يرجى تحديد سطر تعديل صالح أسفل شريط العناوين.');
    return;
  }
  undoAndDeleteLogAtRow(sheet, row);
}

function undoLatestEdit() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.logs);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('لا توجد تعديلات مسجلة للتراجع عنها.');
    return;
  }
  undoAndDeleteLogAtRow(sheet, sheet.getLastRow());
}

function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  getOrCreateSheet(SHEETS.exp, ['ID', 'البند', 'المبلغ', 'التاريخ', 'تاريخ ووقت آخر تعديل']);
  getOrCreateSheet(SHEETS.badrE, ['ID', 'البند', 'المبلغ', 'تاريخ ووقت آخر تعديل']);
  getOrCreateSheet(SHEETS.badrI, ['ID', 'البند', 'المبلغ', 'تاريخ ووقت آخر تعديل']);
  getOrCreateSheet(SHEETS.expected, ['ID', 'البند المتوقع', 'المبلغ المتوقع', 'تاريخ ووقت التعديل'], '#0284c7');
  getOrCreateSheet(SHEETS.settings, ['المفتاح', 'الوصف', 'القيمة', 'تاريخ ووقت آخر تعديل']);
  
  var data = {
    bank: 0,
    fund: 0,
    sister: 0,
    exp: [],
    badrE: [],
    badrI: [],
    expected: [],
    expectedExp: [],
    g21: { g: 0, p: 0 },
    g24: { g: 0, p: 0 },
    logs: []
  };
  
  var setSheet = ss.getSheetByName(SHEETS.settings);
  if (setSheet) {
    var setRows = setSheet.getDataRange().getValues();
    for (var i = 1; i < setRows.length; i++) {
      var k = String(setRows[i][0]).trim();
      var v = parseFloat(setRows[i][2]) || 0;
      if (k === 'bank') data.bank = v;
      else if (k === 'fund') data.fund = v;
      else if (k === 'sister') data.sister = v;
      else if (k === 'g21_g') data.g21.g = v;
      else if (k === 'g21_p') data.g21.p = v;
      else if (k === 'g24_g') data.g24.g = v;
      else if (k === 'g24_p') data.g24.p = v;
    }
  }
  
  data.exp = readListWithDate(ss.getSheetByName(SHEETS.exp));
  data.badrE = readListSimple(ss.getSheetByName(SHEETS.badrE));
  data.badrI = readListSimple(ss.getSheetByName(SHEETS.badrI));
  data.expected = readListSimple(ss.getSheetByName(SHEETS.expected) || ss.getSheetByName('توقعات الشهر القادم'));
  data.expectedExp = readListSimple(ss.getSheetByName(SHEETS.expectedExp));
  data.logs = readLogs(ss.getSheetByName(SHEETS.logs));
  
  return data;
}

function readListWithDate(sheet) {
  var list = [];
  if (!sheet) return list;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var id = rows[i][0];
    var name = rows[i][1];
    var amt = rows[i][2];
    var dt = rows[i][3];
    var lastMod = rows[i][4] || '';
    if (name !== '' && name !== undefined) {
      if (dt instanceof Date) {
        dt = Utilities.formatDate(dt, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
      } else {
        dt = String(dt || '');
      }
      list.push({
        id: id || (Date.now() + i),
        n: String(name),
        a: parseFloat(amt) || 0,
        d: dt,
        updatedAt: String(lastMod)
      });
    }
  }
  return list;
}

function readListSimple(sheet) {
  var list = [];
  if (!sheet) return list;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var id = rows[i][0];
    var name = rows[i][1];
    var amt = rows[i][2];
    var lastMod = rows[i][3] || '';
    if (name !== '' && name !== undefined) {
      list.push({
        id: id || (Date.now() + i),
        n: String(name),
        a: parseFloat(amt) || 0,
        updatedAt: String(lastMod)
      });
    }
  }
  return list;
}

function readLogs(sheet) {
  var list = [];
  if (!sheet) return list;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var checked = rows[i][0];
    var n = rows[i][1];
    var sec = rows[i][2];
    var oldV = rows[i][3];
    var newV = rows[i][4];
    var dt = rows[i][5];
    var type = rows[i][6];
    var itemId = rows[i][7];
    var desc = rows[i][8];
    
    if (dt || n || itemId) {
      list.push({
        id: i,
        itemId: String(itemId || ''),
        checked: (checked === true || checked === 'TRUE'),
        dt: String(dt || ''),
        type: String(type || ''),
        sec: String(sec || ''),
        n: String(n || ''),
        oldV: oldV,
        newV: newV,
        desc: String(desc || '')
      });
    }
  }
  return list.reverse();
}

function uploadAllData(data) {
  if (!data) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeStr = formatDateTime(new Date());
  
  var oldInc = ss.getSheetByName('الإيرادات');
  if (oldInc) {
    try { ss.deleteSheet(oldInc); } catch(e){}
  }
  
  saveSettings({
    bank: data.bank,
    fund: data.fund,
    sister: data.sister,
    g21: data.g21,
    g24: data.g24
  });
  
  var expSheet = getOrCreateSheet(SHEETS.exp, ['ID', 'البند', 'المبلغ', 'التاريخ', 'تاريخ ووقت آخر تعديل']);
  clearSheetData(expSheet);
  if (data.exp && data.exp.length > 0) {
    var expRows = data.exp.map(function(x) {
      return [x.id, x.n, x.a, x.d || '', x.updatedAt || timeStr];
    });
    expSheet.getRange(2, 1, expRows.length, 5).setValues(expRows);
  }
  
  var badrESheet = getOrCreateSheet(SHEETS.badrE, ['ID', 'البند', 'المبلغ', 'تاريخ ووقت آخر تعديل']);
  clearSheetData(badrESheet);
  if (data.badrE && data.badrE.length > 0) {
    var bERows = data.badrE.map(function(x) {
      return [x.id, x.n, x.a, x.updatedAt || timeStr];
    });
    badrESheet.getRange(2, 1, bERows.length, 4).setValues(bERows);
  }
  
  var badrISheet = getOrCreateSheet(SHEETS.badrI, ['ID', 'البند', 'المبلغ', 'تاريخ ووقت آخر تعديل']);
  clearSheetData(badrISheet);
  if (data.badrI && data.badrI.length > 0) {
    var bIRows = data.badrI.map(function(x) {
      return [x.id, x.n, x.a, x.updatedAt || timeStr];
    });
    badrISheet.getRange(2, 1, bIRows.length, 4).setValues(bIRows);
  }

  var expctSheet = getOrCreateSheet(SHEETS.expected, ['ID', 'البند المتوقع (إيراد)', 'المبلغ المتوقع', 'تاريخ ووقت التعديل'], '#0284c7');
  clearSheetData(expctSheet);
  if (data.expected && data.expected.length > 0) {
    var expctRows = data.expected.map(function(x) {
      return [x.id, x.n, x.a, x.updatedAt || timeStr];
    });
    expctSheet.getRange(2, 1, expctRows.length, 4).setValues(expctRows);
  }

  var expctExpSheet = getOrCreateSheet(SHEETS.expectedExp, ['ID', 'المصروف المتوقع', 'المبلغ المتوقع', 'تاريخ ووقت التعديل'], '#b42318');
  clearSheetData(expctExpSheet);
  if (data.expectedExp && data.expectedExp.length > 0) {
    var expctExpRows = data.expectedExp.map(function(x) {
      return [x.id, x.n, x.a, x.updatedAt || timeStr];
    });
    expctExpSheet.getRange(2, 1, expctExpRows.length, 4).setValues(expctExpRows);
  }
}

function clearSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function saveSettings(s, oldSettings) {
  var sheet = getOrCreateSheet(SHEETS.settings, ['المفتاح', 'الوصف', 'القيمة', 'تاريخ ووقت آخر تعديل']);
  var timeStr = formatDateTime(new Date());
  var rows = [
    ['bank', 'إجمالي الحساب البنكي', s.bank !== undefined ? s.bank : 0, timeStr],
    ['fund', 'فلوس البيت الإجمالية', s.fund !== undefined ? s.fund : 0, timeStr],
    ['sister', 'فلوس أختي', s.sister !== undefined ? s.sister : 0, timeStr],
    ['g21_g', 'جرامات دهب عيار 21', s.g21 ? s.g21.g : 0, timeStr],
    ['g21_p', 'سعر جرام دهب عيار 21', s.g21 ? s.g21.p : 0, timeStr],
    ['g24_g', 'جرامات دهب عيار 24', s.g24 ? s.g24.g : 0, timeStr],
    ['g24_p', 'سعر جرام دهب عيار 24', s.g24 ? s.g24.p : 0, timeStr]
  ];
  
  clearSheetData(sheet);
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  
  if (oldSettings) {
    if (s.bank !== undefined && s.bank !== oldSettings.bank) {
      addLog('💰 تعديل رصيد', 'الأرصدة', 'bank', 'الحساب البنكي', oldSettings.bank, s.bank, 'تعديل رصيد البنك من ' + oldSettings.bank + ' إلى ' + s.bank);
    }
    if (s.fund !== undefined && s.fund !== oldSettings.fund) {
      addLog('💰 تعديل رصيد', 'الأرصدة', 'fund', 'فلوس البيت الإجمالية', oldSettings.fund, s.fund, 'تعديل فلوس البيت من ' + oldSettings.fund + ' إلى ' + s.fund);
    }
    if (s.sister !== undefined && s.sister !== oldSettings.sister) {
      addLog('💰 تعديل رصيد', 'الأرصدة', 'sister', 'فلوس أختي', oldSettings.sister, s.sister, 'تعديل فلوس أختي من ' + oldSettings.sister + ' إلى ' + s.sister);
    }
    if (s.g21 && oldSettings.g21 && (s.g21.p !== oldSettings.g21.p || s.g21.g !== oldSettings.g21.g)) {
      addLog('🪙 تعديل دهب', 'الدهب', 'g21', 'عيار 21', oldSettings.g21.p, s.g21.p, 'تحديث دهب عيار 21');
    }
    if (s.g24 && oldSettings.g24 && (s.g24.p !== oldSettings.g24.p || s.g24.g !== oldSettings.g24.g)) {
      addLog('🪙 تعديل دهب', 'الدهب', 'g24', 'عيار 24', oldSettings.g24.p, s.g24.p, 'تحديث دهب عيار 24');
    }
  }
}

function addRecord(listName, item) {
  var sheetName = SHEETS[listName];
  if (!sheetName) throw new Error('قائمة غير صالحة');
  
  var isDated = (listName === 'exp');
  var headers = isDated 
    ? ['ID', 'البند', 'المبلغ', 'التاريخ', 'تاريخ ووقت آخر تعديل'] 
    : ['ID', 'البند', 'المبلغ', 'تاريخ ووقت آخر تعديل'];
  var sheet = getOrCreateSheet(sheetName, headers);
  var timeStr = formatDateTime(new Date());
  
  var row = isDated
    ? [item.id, item.n, item.a, item.d || '', timeStr]
    : [item.id, item.n, item.a, timeStr];
    
  sheet.appendRow(row);
  
  var secName = listName === 'exp' ? 'مصاريف البيت' : 
                listName === 'badrE' ? 'مصاريف بدر' : 
                listName === 'badrI' ? 'إيرادات بدر' : 
                listName === 'expected' ? 'إيرادات متوقعة للشهر القادم' : 
                listName === 'expectedExp' ? 'مصاريف متوقعة للشهر القادم' : 'أخرى';
  addLog('➕ إضافة بند', secName, item.id, item.n, '', item.a, 'إضافة ' + item.n + ' بمبلغ ' + item.a + ' ج');
}

function updateRecord(listName, item, oldItem) {
  var sheetName = SHEETS[listName];
  if (!sheetName) throw new Error('قائمة غير صالحة');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || (listName === 'expected' ? ss.getSheetByName('توقعات الشهر القادم') : null);
  if (!sheet) throw new Error('الصفحة غير موجودة');
  
  var data = sheet.getDataRange().getValues();
  var targetId = String(item.id);
  var rowIndex = -1;
  var oldName = oldItem ? oldItem.n : '';
  var oldAmt = oldItem ? oldItem.a : '';
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === targetId) {
      rowIndex = i + 1;
      if (!oldName) oldName = String(data[i][1]);
      if (!oldAmt) oldAmt = data[i][2];
      break;
    }
  }
  
  var timeStr = formatDateTime(new Date());
  
  if (rowIndex === -1) {
    addRecord(listName, item);
    return;
  }
  
  var isDated = (listName === 'exp');
  sheet.getRange(rowIndex, 2).setValue(item.n);
  sheet.getRange(rowIndex, 3).setValue(item.a);
  if (isDated) {
    sheet.getRange(rowIndex, 4).setValue(item.d || '');
    sheet.getRange(rowIndex, 5).setValue(timeStr);
  } else {
    sheet.getRange(rowIndex, 4).setValue(timeStr);
  }
  
  var secName = listName === 'exp' ? 'مصاريف البيت' : 
                listName === 'badrE' ? 'مصاريف بدر' : 
                listName === 'badrI' ? 'إيرادات بدر' : 
                listName === 'expected' ? 'إيرادات متوقعة للشهر القادم' : 
                listName === 'expectedExp' ? 'مصاريف متوقعة للشهر القادم' : 'أخرى';
  var details = 'تعديل البند من "' + oldName + ' (' + oldAmt + ' ج)" إلى "' + item.n + ' (' + item.a + ' ج)"';
  addLog('✏️ تعديل بند', secName, item.id, oldName, oldAmt, item.a, details);
}

function deleteRecord(listName, id, oldItem) {
  var sheetName = SHEETS[listName];
  if (!sheetName) throw new Error('قائمة غير صالحة');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || (listName === 'expected' ? ss.getSheetByName('توقعات الشهر القادم') : null);
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  var targetId = String(id);
  var oldName = oldItem ? oldItem.n : '';
  var oldAmt = oldItem ? oldItem.a : '';
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === targetId) {
      if (!oldName) oldName = String(data[i][1]);
      if (!oldAmt) oldAmt = data[i][2];
      sheet.deleteRow(i + 1);
      break;
    }
  }
  
  var secName = listName === 'exp' ? 'مصاريف البيت' : 
                listName === 'badrE' ? 'مصاريف بدر' : 
                listName === 'badrI' ? 'إيرادات بدر' : 
                listName === 'expected' ? 'إيرادات متوقعة للشهر القادم' : 
                listName === 'expectedExp' ? 'مصاريف متوقعة للشهر القادم' : 'أخرى';
  addLog('🗑️ حذف بند', secName, id, oldName, oldAmt, 0, 'حذف بند "' + oldName + '" نهائياً');
}

function deleteLog(logRowId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.logs);
  if (!sheet) return;
  var row = parseInt(logRowId);
  if (row > 1 && row <= sheet.getLastRow()) {
    sheet.deleteRow(row);
  }
}

function clearAllLogs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.logs);
  if (!sheet) return;
  clearSheetData(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast('تم مسح سجل التعديلات بالكامل.', 'حساب البيت', 4);
}
