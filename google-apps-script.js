// =====================================================
// CÓDIGO PARA GOOGLE APPS SCRIPT - NUEVA ESTRUCTURA CALENDARIO
// =====================================================
//
// PLANILLA CREADA: https://docs.google.com/spreadsheets/d/13_J1zgZT5RS9PnqJNoKWxZhQDSsqO1yb4fZnavB2IBw/edit
//
// PARA ACTUALIZAR EL DEPLOYMENT:
// 1. Ir a: https://script.google.com/home/projects/1witZS0YSoXUnYWxGpFRlbM_WevvZKQ5c6kFJowLgfWD8AR_zzG2VbbsY/edit
// 2. Cerrar el aviso de Rhino (botón "Cerrar")
// 3. Rápidamente: Implementar → Gestionar implementaciones
// 4. Click en el ícono lápiz (editar)
// 5. Versión: "Nueva versión"
// 6. Click en "Implementar"
// La URL del web app queda igual.
//
// ESTRUCTURA DE LA PLANILLA:
// - Una pestaña por mes (ej: "Junio 2026")
// - Columna A: día ("Mar 16")
// - Columnas B-R: slots de horario (09:00 a 21:00 cada 45 min)
// - Cada celda admite 2 reservas (2 sillas)
// - Lunes aparecen como CERRADO
// - Trigger a medianoche para crear el próximo mes automáticamente
//
// CANCELAR UN TURNO: borrar el nombre de la celda en la planilla.
// Si la celda tenía 2 nombres y borrás uno, ese slot vuelve a aparecer disponible.
// =====================================================

const SLOTS = ['09:00','09:45','10:30','11:15','12:00','12:45','13:30','14:15','15:00','15:45','16:30','17:15','18:00','18:45','19:30','20:15','21:00'];
const DAYS_ES = ['Dom','Lun','Mar','Mie','Jue','Vie','Sab'];
const MONTHS_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Sin planilla configurada.');
  return SpreadsheetApp.openById(id);
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
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#111111').setFontColor('#E6E628');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  for (var c = 2; c <= SLOTS.length + 1; c++) sheet.setColumnWidth(c, 145);
  for (var d = 1; d <= daysInMonth; d++) {
    var dow = new Date(year, month, d).getDay();
    var label = DAYS_ES[dow] + ' ' + (d < 10 ? '0'+d : d);
    var row = d + 1;
    sheet.getRange(row, 1).setValue(label).setFontWeight('bold').setBackground('#1c1c1c');
    if (dow === 1) {
      sheet.getRange(row, 2, 1, SLOTS.length).setValue('CERRADO').setBackground('#2a2020').setFontColor('#555555').setHorizontalAlignment('center');
    } else {
      sheet.getRange(row, 2, 1, SLOTS.length).setBackground('#1a1a1a').setFontColor('#cccccc').setVerticalAlignment('top').setWrap(true);
    }
    sheet.setRowHeight(row, 60);
  }
}

// Ejecutar manualmente UNA VEZ si SS_ID no está configurado:
function setupSS() {
  // La planilla ya fue creada. Solo setear el ID si se perdió:
  PropertiesService.getScriptProperties().setProperty('SS_ID', '13_J1zgZT5RS9PnqJNoKWxZhQDSsqO1yb4fZnavB2IBw');
  Logger.log('SS_ID configurado.');
}

// Trigger: crear en Activadores → basado en tiempo → cada día a las 00:00
function midnightTask() {
  var ss = getSpreadsheet();
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
  var fecha = e.parameter.fecha;
  var parts = fecha.split('-');
  var y = parseInt(parts[0]), m = parseInt(parts[1]);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(monthSheetName(new Date(y, m - 1, 1)));
    if (!sheet) return jsonOut({ ocupados: [] });
    var row = dateRow(fecha);
    var vals = sheet.getRange(row, 2, 1, SLOTS.length).getValues()[0];
    var ocupados = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i]).trim();
      if (v === 'CERRADO') { ocupados.push(SLOTS[i]); continue; }
      var entries = v.split('─────').filter(function(x){ return x.trim() !== ''; });
      if (entries.length >= 2) ocupados.push(SLOTS[i]);
    }
    return jsonOut({ ocupados: ocupados });
  } catch(err) { return jsonOut({ ocupados: [], error: err.message }); }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (!data.fecha || !data.hora || !data.nombre || !data.telefono)
    return jsonOut({ ok: false, error: 'Faltan datos.' });
  var parts = data.fecha.split('-');
  var y = parseInt(parts[0]), m = parseInt(parts[1]);
  try {
    var ss = getSpreadsheet();
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
    if (entries.length >= 2) return jsonOut({ ok: false, error: 'Horario completo (2 turnos).' });
    var newEntry = data.nombre + (data.apellido ? ' ' + data.apellido : '') + ' · ' + data.telefono;
    var newVal = entries.length === 0 ? newEntry : entries[0] + '\n' + SEP + '\n' + newEntry;
    cell.setValue(newVal);
    cell.setBackground(entries.length === 1 ? '#2a1a1a' : '#1a2a1a');
    return jsonOut({ ok: true, silla: entries.length + 1 });
  } catch(err) { return jsonOut({ ok: false, error: err.message }); }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
