// =====================================================
// CÓDIGO PARA GOOGLE APPS SCRIPT (no va en el sitio web)
// =====================================================
// PASO 1: Entrá a https://sheets.google.com y creá una planilla nueva.
//         Llamala "Turnos Fernández Salón".
//
// PASO 2: En la hoja, en la fila 1, escribí estos encabezados (uno por columna):
//         Fecha | Hora | Servicio | Nombre | Apellido | Telefono | CreadoEn
//
// PASO 3: Arriba en el menú: Extensiones > Apps Script.
//         Borrá lo que haya y pegá TODO este archivo. Guardá (Ctrl+S).
//
// PASO 4: Arriba a la derecha, botón "Implementar" > "Nueva implementación".
//         Tipo: "Aplicación web". Ejecutar como: "Yo". Quién tiene acceso: "Cualquier usuario".
//         Hacé clic en "Implementar" y AUTORIZÁ los permisos que pida Google.
//
// PASO 5: Copiá la URL que te da ("URL de la aplicación web", termina en /exec)
//         y pegala en index.html donde dice SHEETS_URL.
//
// PARA CANCELAR UN TURNO: simplemente abrí la planilla y borrá la fila de ese turno.
// Apenas la borrás, ese horario vuelve a aparecer disponible en la web.
//
// Si en el futuro cambiás algo en este código, tenés que volver a "Implementar" >
// "Gestionar implementaciones" > ícono de lápiz > "Nueva versión" > Implementar.
// =====================================================

const SHEET_NAME = 'Turnos';

// Columnas: Fecha(0) Hora(1) Servicio(2) Nombre(3) Apellido(4) Telefono(5) Confirmado(6) CreadoEn(7)
// Un turno está ocupado si existe la fila Y Confirmado != "No confirmó"

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const fecha = e.parameter.fecha;
  const rows = sheet.getDataRange().getValues();
  const ocupados = [];
  for (let i = 1; i < rows.length; i++) {
    const confirmado = String(rows[i][6]).trim();
    if (String(rows[i][0]) === fecha && confirmado !== 'No confirmó') {
      ocupados.push(String(rows[i][1]));
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ocupados }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = JSON.parse(e.postData.contents);

  if (!data.fecha || !data.hora || !data.nombre || !data.telefono) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Faltan datos obligatorios.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const confirmado = String(rows[i][6]).trim();
    if (String(rows[i][0]) === data.fecha && String(rows[i][1]) === data.hora && confirmado !== 'No confirmó') {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Ese horario ya fue reservado.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow([
    data.fecha,
    data.hora,
    data.servicio || '',
    data.nombre,
    data.apellido || '',
    data.telefono,
    '', // Confirmado — lo completás vos en la planilla
    new Date(),
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
