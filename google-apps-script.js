/**
 * ╔══════════════════════════════════════════════════════════════╗
 *  SISTEMA DE EVALUACIÓN LITPA — Google Apps Script
 * ╠══════════════════════════════════════════════════════════════╣
 *
 *  PASOS PARA CONFIGURAR:
 *
 *  1. Ve a https://script.google.com y crea un nuevo proyecto.
 *  2. Borra el código que viene por defecto.
 *  3. Pega TODO este archivo.
 *  4. Cambia SHEET_ID por el ID de tu Google Sheet
 *     (está en la URL: docs.google.com/spreadsheets/d/[AQUÍ]/edit)
 *  5. (Opcional) Pon tu webhook de Discord en DISCORD_WEBHOOK_URL.
 *  6. Guarda el proyecto (Ctrl+S).
 *  7. Despliega:
 *       Implementar → Nueva implementación
 *       Tipo: Aplicación web
 *       Ejecutar como: Yo (tu cuenta)
 *       Quién tiene acceso: Cualquier persona
 *       → Implementar → Autorizar → Copiar la URL
 *  8. Pega esa URL en evaluadores.html y admin.html
 *     donde dice: const APPS_SCRIPT_URL = '...'
 *
 * ╚══════════════════════════════════════════════════════════════╝
 */

// ── CONFIGURACIÓN ─────────────────────────────────────────────────
const SHEET_ID            = '16SHzGQQYqX7IWelFy_3kRF_mbNAdfUiB0UpA717-fak';
const SHEET_NAME          = 'Calificaciones';
const DISCORD_WEBHOOK_URL = '';   // Opcional — pega tu webhook aquí
// ──────────────────────────────────────────────────────────────────

// Cabeceras de la hoja
const HEADERS = [
  'Timestamp', 'Evaluador', 'Participante', 'Equipo',
  'Presentación', 'Contenido', 'Innovación', 'Impacto',
  'Total', 'Comentarios'
];

/**
 * Maneja peticiones GET — el panel admin solicita los datos
 * URL: ?action=get
 */
function doGet(e) {
  // Permitir CORS
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const action = (e && e.parameter && e.parameter.action) || '';

    if (action === 'get') {
      const data = leerCalificaciones();
      output.setContent(JSON.stringify(data));
    } else {
      output.setContent(JSON.stringify({ status: 'ok', message: 'LITPA API activa' }));
    }
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.toString() }));
  }

  return output;
}

/**
 * Maneja peticiones POST — los evaluadores envían calificaciones
 */
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    guardarCalificacion(data);

    if (DISCORD_WEBHOOK_URL) {
      notificarDiscord(data);
    }

    output.setContent(JSON.stringify({ status: 'success' }));
  } catch (err) {
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
  }

  return output;
}

/**
 * Guarda una calificación en la hoja de cálculo.
 * Si la hoja no existe, la crea con encabezados formateados.
 */
function guardarCalificacion(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);

    // Formato de encabezados
    headerRange.setBackground('#1a1a1a');
    headerRange.setFontColor('#d4af37');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // Congelar fila de encabezados
    sheet.setFrozenRows(1);

    // Ancho de columnas
    sheet.setColumnWidth(1, 180);  // Timestamp
    sheet.setColumnWidth(2, 120);  // Evaluador
    sheet.setColumnWidth(3, 200);  // Participante
    sheet.setColumnWidth(4, 180);  // Equipo
    sheet.setColumnWidth(9, 80);   // Total
    sheet.setColumnWidth(10, 250); // Comentarios
  }

  // Agregar fila de datos
  sheet.appendRow([
    data.timestamp    || new Date().toISOString(),
    data.evaluador    || '',
    data.nombre       || '',
    data.equipo       || '',
    Number(data.presentacion) || 0,
    Number(data.contenido)    || 0,
    Number(data.innovacion)   || 0,
    Number(data.impacto)      || 0,
    Number(data.total)        || 0,
    data.comentarios  || ''
  ]);

  // Colorear la fila recién agregada según el evaluador
  const lastRow   = sheet.getLastRow();
  const rowRange  = sheet.getRange(lastRow, 1, 1, HEADERS.length);
  const evalColor = {
    'Evaluador 1': '#1a1200',
    'Evaluador 2': '#0d1a00',
    'Evaluador 3': '#001a1a'
  };
  const bg = evalColor[data.evaluador] || '#111111';
  rowRange.setBackground(bg);
  rowRange.setFontColor('#d4af37');
}

/**
 * Lee todas las calificaciones y las devuelve como array de objetos.
 */
function leerCalificaciones() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();

  return rows.map(row => ({
    timestamp:    row[0] ? new Date(row[0]).toISOString() : '',
    evaluador:    row[1],
    nombre:       row[2],
    equipo:       row[3],
    presentacion: row[4],
    contenido:    row[5],
    innovacion:   row[6],
    impacto:      row[7],
    total:        row[8],
    comentarios:  row[9]
  }));
}

/**
 * Envía una notificación embed a Discord cuando se registra una calificación.
 */
function notificarDiscord(data) {
  if (!DISCORD_WEBHOOK_URL) return;

  const total     = Number(data.total) || 0;
  const pct       = Math.round((total / 40) * 100);
  const medalEmoji = pct >= 90 ? '🥇' : pct >= 75 ? '🥈' : pct >= 60 ? '🥉' : '📋';

  const payload = {
    username:   'Sistema LITPA',
    avatar_url: '',
    embeds: [{
      title:       medalEmoji + ' Nueva Calificación Registrada',
      description: 'Se ha registrado una nueva evaluación en el sistema.',
      color:       0xD4AF37,
      fields: [
        { name: '👤 Evaluador',    value: String(data.evaluador    || '—'), inline: true  },
        { name: '🎓 Participante', value: String(data.nombre       || '—'), inline: true  },
        { name: '🏆 Equipo',       value: String(data.equipo       || '—'), inline: false },
        { name: '📊 Presentación', value: String(data.presentacion || 0),   inline: true  },
        { name: '📚 Contenido',    value: String(data.contenido    || 0),   inline: true  },
        { name: '💡 Innovación',   value: String(data.innovacion   || 0),   inline: true  },
        { name: '🌟 Impacto',      value: String(data.impacto      || 0),   inline: true  },
        { name: '🏅 TOTAL',        value: '**' + total + ' / 40** (' + pct + '%)', inline: false }
      ],
      footer: {
        text: 'LITPA · Sistema de Evaluación · ' +
              new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })
      },
      timestamp: new Date().toISOString()
    }]
  };

  UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, {
    method:          'post',
    contentType:     'application/json',
    payload:         JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

/**
 * Función de prueba — ejecútala manualmente desde el editor
 * para verificar que la conexión con Sheets funciona.
 */
function testConexion() {
  const datos = {
    evaluador:    'Evaluador 1',
    nombre:       'Prueba Conexión',
    equipo:       'Equipo Test',
    presentacion: 9,
    contenido:    8.5,
    innovacion:   9,
    impacto:      8,
    total:        34.5,
    comentarios:  'Registro de prueba — puedes borrar esta fila.',
    timestamp:    new Date().toISOString()
  };

  guardarCalificacion(datos);
  Logger.log('✅ Prueba exitosa — revisa tu Google Sheet.');
}
