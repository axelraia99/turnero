const SLOTS = ['09:00','09:45','10:30','11:15','12:00','12:45','13:30','14:15','15:00','15:45','16:30','17:15','18:00','18:45','19:30','20:15','21:00'];
const DAYS_ES = ['Dom','Lun','Mar','Mie','Jue','Vie','Sab'];
const MONTHS_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function getSS() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Sin planilla. Ejecuta setupSS primero.');
  return SpreadsheetApp.openById(id);
}

function setupSS() {
  PropertiesService.getScriptProperties().setProperty('SS_ID', '13_J1zgZT5RS9PnqJNoKWxZhQDSsqO1yb4fZnavB2IBw');
  var ss = getSS();
  var now = new Date();
  var next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  getOrCreateMonthSheet(ss, now);
  getOrCreateMonthSheet(ss, next);
  Logger.log('Listo. URL: ' + ss.getUrl());
}

function resetSheets() {
  var ss = getSS();
  var sheets = ss.getSheets();
  for (var i = sheets.length - 1; i > 0; i--) ss.deleteSheet(sheets[i]);
  sheets[0].clear();
  var now = new Date();
  var next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  getOrCreateMonthSheet(ss, now);
  getOrCreateMonthSheet(ss, next);
  Logger.log('✓ Sheets recreados con nuevos estilos');
}

function generateAllMonths() {
  var ss = getSS();
  var now = new Date();
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth();

  for (var m = currentMonth; m <= 11; m++) {
    var date = new Date(currentYear, m, 1);
    var sheetName = monthSheetName(date);
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) ss.deleteSheet(sheet);
    getOrCreateMonthSheet(ss, date);
  }
  Logger.log('✓ Todos los meses generados hasta diciembre ' + currentYear);
}

function monthSheetName(date) {
  return MONTHS_ES[date.getMonth()] + ' ' + date.getFullYear();
}

function getOrCreateMonthSheet(ss, date) {
  var name = monthSheetName(date);
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); buildMonthSheet(sheet, date); }
  return sheet;
}

function buildMonthSheet(sheet, date) {
  var year = date.getFullYear();
  var month = date.getMonth();
  var daysInMonth = new Date(year, month + 1, 0).getDate();
  var header = ['Dia'].concat(SLOTS);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#FFFFFF').setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  for (var c = 2; c <= SLOTS.length + 1; c++) sheet.setColumnWidth(c, 145);
  for (var d = 1; d <= daysInMonth; d++) {
    var dow = new Date(year, month, d).getDay();
    var label = DAYS_ES[dow] + ' ' + (d < 10 ? '0'+d : d);
    var row = d + 1;
    sheet.getRange(row, 1).setValue(label).setFontWeight('bold').setBackground('#FFFFFF').setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (dow === 1) {
      sheet.getRange(row, 2, 1, SLOTS.length).setValue('CERRADO').setBackground('#FFFFFF').setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle');
    } else {
      sheet.getRange(row, 2, 1, SLOTS.length).setBackground('#FFFFFF').setFontColor('#000000').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    }
    sheet.setRowHeight(row, 60);
  }
}

function midnightTask() {
  var ss = getSS();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  getOrCreateMonthSheet(ss, tomorrow);
  var lastDay = new Date(tomorrow.getFullYear(), tomorrow.getMonth() + 1, 0).getDate();
  if (tomorrow.getDate() === lastDay) {
    getOrCreateMonthSheet(ss, new Date(tomorrow.getFullYear(), tomorrow.getMonth() + 1, 1));
  }
}

function dateRow(fecha) {
  return parseInt(fecha.split('-')[2]) + 1;
}

function slotCol(hora) {
  var i = SLOTS.indexOf(hora);
  return i === -1 ? -1 : i + 2;
}

function doGet(e) {
  if (e.parameter.reset === 'true') {
    resetSheets();
    return jsonOut({ ok: true, msg: 'Sheets recreados' });
  }
  if (e.parameter.generate === 'true') {
    generateAllMonths();
    return jsonOut({ ok: true, msg: 'Todos los meses generados hasta diciembre' });
  }
  var fecha = e.parameter.fecha;
  var parts = fecha.split('-');
  var y = parseInt(parts[0]), m = parseInt(parts[1]);
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName(monthSheetName(new Date(y, m - 1, 1)));
    if (!sheet) return jsonOut({ nico: [], david: [] });
    var row = dateRow(fecha);
    var vals = sheet.getRange(row, 2, 1, SLOTS.length).getValues()[0];
    var nicoOcupados = [];
    var davidOcupados = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i]).trim();
      if (v === 'CERRADO') { nicoOcupados.push(SLOTS[i]); davidOcupados.push(SLOTS[i]); continue; }
      var entries = v.split('─────').filter(function(x){ return x.trim() !== ''; });
      for (var j = 0; j < entries.length; j++) {
        var entry = entries[j].trim().toLowerCase();
        if (entry.indexOf('nico:') === 0) nicoOcupados.push(SLOTS[i]);
        if (entry.indexOf('david:') === 0) davidOcupados.push(SLOTS[i]);
      }
    }
    return jsonOut({ nico: nicoOcupados, david: davidOcupados });
  } catch(err) { return jsonOut({ nico: [], david: [], error: err.message }); }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (!data.fecha || !data.hora || !data.nombre || !data.telefono || !data.silla)
    return jsonOut({ ok: false, error: 'Faltan datos.' });
  var sillaNombre = data.silla === 'nico' ? 'Nico' : 'David';
  var parts = data.fecha.split('-');
  var y = parseInt(parts[0]), m = parseInt(parts[1]);
  try {
    var ss = getSS();
    var sheetName = monthSheetName(new Date(y, m - 1, 1));
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheet = ss.insertSheet(sheetName); buildMonthSheet(sheet, new Date(y, m - 1, 1)); }
    var row = dateRow(data.fecha);
    var col = slotCol(data.hora);
    if (col === -1) return jsonOut({ ok: false, error: 'Horario invalido.' });
    var cell = sheet.getRange(row, col);
    var current = String(cell.getValue()).trim();
    if (current === 'CERRADO') return jsonOut({ ok: false, error: 'Dia cerrado.' });
    var SEP = '─────';
    var entries = current === '' ? [] : current.split(SEP).map(function(x){ return x.trim(); }).filter(Boolean);
    var sillaOcupada = entries.some(function(en) {
      return en.toLowerCase().indexOf(sillaNombre.toLowerCase() + ':') === 0;
    });
    if (sillaOcupada) return jsonOut({ ok: false, error: 'Silla ' + sillaNombre + ' ya ocupada en este horario.' });
    var fullName = data.nombre + (data.apellido ? ' ' + data.apellido : '');
    var phone = data.telefono.replace(/\D/g, '');
    var whatsappUrl = 'https://wa.me/54' + phone;
    var displayText = sillaNombre + ': ' + fullName + ' · ' + data.telefono;
    if (entries.length === 0) {
      cell.setFormula('=HYPERLINK("' + whatsappUrl + '","' + displayText + '")');
    } else {
      cell.setValue(entries[0] + '\n' + SEP + '\n' + displayText);
    }
    var newEntries = entries.concat([displayText]);
    var ambosOcupados = newEntries.some(function(en){ return en.toLowerCase().indexOf('nico:') === 0; }) &&
                        newEntries.some(function(en){ return en.toLowerCase().indexOf('david:') === 0; });
    cell.setBackground(ambosOcupados ? '#90EE90' : '#FFFFFF').setFontColor('#000000');
    return jsonOut({ ok: true, silla: sillaNombre });
  } catch(err) { return jsonOut({ ok: false, error: err.message }); }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function repaintSheets() {
  var ss = getSS();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var maxRow = sheet.getLastRow();
    var maxCol = sheet.getLastColumn();
    if (maxRow < 2) continue;
    sheet.getRange(1, 1, 1, maxCol).setBackground('#FFFFFF').setFontColor('#000000');
    for (var r = 2; r <= maxRow; r++) {
      var cell = sheet.getRange(r, 1);
      var val = String(cell.getValue()).trim();
      if (val === 'CERRADO') {
        sheet.getRange(r, 2, 1, SLOTS.length).setBackground('#F0F0F0').setFontColor('#CCCCCC');
      } else {
        sheet.getRange(r, 2, 1, maxCol - 1).setBackground('#FFFFFF').setFontColor('#000000');
      }
      cell.setBackground('#FFFFFF').setFontColor('#000000');
    }
  }
  Logger.log('Sheets repainted');
}
