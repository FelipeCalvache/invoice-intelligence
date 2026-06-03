const SPREADSHEET_ID = '1oNuxvkJ3o_k8Db8EKvZ5Cjkw-bADcjupOrsFyejrOgk';
const SHEET_NAME = 'Transacciones';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Fecha/Hora', 'Monto', 'Moneda', 'Comercio', 'Tipo', 'App Banco', 'Descripción']);
    }

    sheet.appendRow([
      data.fechaHora,
      data.monto,
      data.moneda,
      data.comercio,
      data.tipo,
      data.appBanco,
      data.descripcion
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testInsertar() {
  const fakeRequest = {
    postData: {
      contents: JSON.stringify({
        fechaHora: '2026-06-01 10:30:00',
        monto: 50000,
        moneda: 'COP',
        comercio: 'Rappi',
        tipo: 'débito',
        appBanco: 'com.nequi.mobilebanking',
        descripcion: 'Pago de $50.000 en Rappi aprobado'
      })
    }
  };
  console.log(doPost(fakeRequest).getContent());
}

// ─── CONFIGURACIÓN ───────────────────────────────────────────
const SHEET_2026 = '2026';
const DRIVE_FOLDER_ID = '1fqJLvkrSS2kCnS1FVXp_LLTJpZR9A2aO';
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const MODEL = "gemini-3.1-flash-lite";
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1/models/${MODEL}:generateContent?key=${GEMINI_API_KEY}`;;

const COL_FECHA = 4; // D
const COL_GASTO = 3; // C
const COL_FACTURA = 7; // G
// ─────────────────────────────────────────────────────────────
const SHEET_PENDIENTES = 'Facturas Pendientes';
// ─────────────────────────────────────────────────────────────

function procesarFacturasNuevas() {
  const props = PropertiesService.getScriptProperties();
  const procesados = JSON.parse(props.getProperty('archivos_procesados') || '[]');

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const files = folder.getFiles();
  const nuevos = [];

  while (files.hasNext()) {
    const file = files.next();
    if (!procesados.includes(file.getId()) && esArchivoValido(file.getMimeType())) {
      nuevos.push(file);
    }
  }

  if (nuevos.length === 0) {
    Logger.log('Sin facturas nuevas.');
    return;
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_2026);
  const data = sheet.getDataRange().getValues();

  for (const file of nuevos) {
    try {
      Logger.log(`Procesando: ${file.getName()}`);
      const datos = analizarFactura(file);

      if (!datos?.fecha || !datos.monto) {
        Logger.log(`⚠️ Gemini no pudo extraer datos de ${file.getName()}`);
        procesados.push(file.getId());
        continue;
      }

      Logger.log(`  → Fecha: ${datos.fecha} | Monto: ${datos.monto} | Hora: ${datos.hora}`);

      const fila = buscarTransaccion(data, datos.fecha, datos.monto);

      if (fila === -1) {
        Logger.log(`  ❌ Sin match para fecha ${datos.fecha} / monto ${datos.monto}`);
        procesados.push(file.getId());
        continue;
      }

      const link = `https://drive.google.com/file/d/${file.getId()}/view`;
      sheet.getRange(fila + 1, COL_FACTURA).setValue(link);

      if (datos.hora) {
        const fechaActualizada = actualizarHora(data[fila][COL_FECHA - 1], datos.hora);
        if (fechaActualizada) {
          sheet.getRange(fila + 1, COL_FECHA).setValue(fechaActualizada);
        }
      }

      Logger.log(`  ✅ Match en fila ${fila + 1} — link guardado`);
      procesados.push(file.getId());

    } catch (e) {
      Logger.log(`❌ Error con ${file.getName()}: ${e.toString()}`);
    }
  }

  props.setProperty('archivos_procesados', JSON.stringify(procesados));
}

// ─── Analiza la imagen con Gemini Vision ───────────────────────
function analizarFactura(file) {
  const base64 = Utilities.base64Encode(file.getBlob().getBytes());
  const mimeType = file.getMimeType();

  const prompt = `Analiza este recibo o factura y extrae SOLO este JSON:
{
  "fecha": "dd/mm/yyyy",
  "hora": "HH:mm o null si no aparece en el recibo",
  "monto": número total sin símbolos ni puntos de miles,
  "comercio": "nombre del comercio"
}
Reglas:
- fecha en formato dd/mm/yyyy exacto
- monto: solo el TOTAL (número entero o decimal, ej: 150000)
- hora: formato 24h si aparece, si no pon null
- Sin texto adicional, SOLO el JSON`;

  const body = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64 } }
      ]
    }]
  };

  const response = UrlFetchApp.fetch(GEMINI_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body)
  });

  const raw = JSON.parse(response.getContentText());
  const text = raw.candidates[0].content.parts[0].text
    .replace(/```json/g, '').replace(/```/g, '').trim();

  return JSON.parse(text);
}

// ─── Busca la fila en la hoja 2026 ─────────────────────────────
function buscarTransaccion(data, fechaFactura, montoFactura) {
  const [dia, mes, anio] = fechaFactura.split('/').map(Number);

  let filaUnica = -1;
  let conteoMonto = 0;

  for (let i = 1; i < data.length; i++) {
    const celda = data[i][COL_FECHA - 1];
    if (!celda) continue;

    const montoSheet = parseFloat(String(data[i][COL_GASTO - 1]).replace(/[^0-9.]/g, ''));
    if (Math.abs(montoSheet - montoFactura) > 1) continue;

    conteoMonto++;
    filaUnica = i;

    // Intento 1: coincidencia exacta fecha + monto
    const fecha = new Date(celda);
    if (!isNaN(fecha)) {
      const mismaFecha = fecha.getDate() === dia &&
        (fecha.getMonth() + 1) === mes &&
        fecha.getFullYear() === anio;
      if (mismaFecha) return i;
    }
  }

  // Intento 2: el monto es único en toda la hoja → match seguro
  if (conteoMonto === 1) {
    Logger.log(`  ⚠️ Fecha no coincidió pero monto ${montoFactura} es único → fila ${filaUnica + 1}`);
    return filaUnica;
  }

  // Monto repetido y sin coincidencia de fecha → ambiguo, no asignar
  return -1;
}

// ─── Actualiza la hora en la fecha existente ───────────────────
function actualizarHora(celdaFecha, horaStr) {
  try {
    const fecha = new Date(celdaFecha);
    const [h, m] = horaStr.split(':').map(Number);
    fecha.setHours(h, m, 0, 0);
    return fecha;
  } catch (e) {
    return null;
  }
}

function esArchivoValido(mimeType) {
  return ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/gif', 'application/pdf'].includes(mimeType);
}

// ─── Ejecuta esto UNA SOLA VEZ para activar el trigger ─────────
function instalarTrigger() {
  // Borra triggers anteriores para no duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'procesarFacturasNuevas') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('procesarFacturasNuevas')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ Trigger instalado — revisará la carpeta cada 5 minutos');
}

function registrarPendiente(spreadsheet, file, datos) {
  let hoja = spreadsheet.getSheetByName(SHEET_PENDIENTES);

  // Crea la hoja si no existe
  if (!hoja) {
    hoja = spreadsheet.insertSheet(SHEET_PENDIENTES);
    hoja.appendRow(['Archivo', 'Link', 'Fecha extraída', 'Monto extraído',
      'Comercio', 'Procesado el', 'Estado']);
    hoja.setFrozenRows(1);
  }

  const link = `https://drive.google.com/file/d/${file.getId()}/view`;
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  hoja.appendRow([
    file.getName(),
    link,
    datos.fecha || '',
    datos.monto || '',
    datos.comercio || '',
    ahora,
    'Pendiente'
  ]);
}
