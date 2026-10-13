/**
 * Respuestas del formulario de seguimiento — CONAMAT Hidalgo – San Cosme
 *
 * El formulario publicado en GitHub Pages manda cada respuesta aquí y
 * se agrega como fila en la pestaña "Respuestas" de esta hoja.
 *
 * Pasos:
 *  1. Crea una hoja de Google nueva (p. ej. "Seguimiento papás 2026").
 *  2. Extensiones → Apps Script, borra lo que haya, pega este código y guarda.
 *  3. Ejecuta configurar() una vez y acepta los permisos.
 *  4. Implementar → Nueva implementación → tipo "Aplicación web"
 *     (Ejecutar como: Yo · Quién tiene acceso: Cualquier persona).
 *  5. Copia la URL que termina en /exec y pásala para conectarla en config.js.
 *  Si después cambias este código: Implementar → Gestionar implementaciones →
 *  editar → Versión: "Nueva versión"; si no, la URL sigue usando el código viejo.
 */

const HOJA = 'Respuestas';
const ENCABEZADOS = ['Fecha', 'Curso', 'Plantel', 'Grupo', 'Alumno', '¿Acepta seguimiento?',
  'Teléfono', 'Horario', 'Parentesco', 'Nombre de quien responde', 'Estatus', 'Notas'];
const CURSOS = ['UNAM', 'UAM', 'IPN', 'COLBACH'];
const PLANTELES = ['Hidalgo', 'San Cosme'];

/** Ejecuta esta función UNA vez para preparar la hoja. */
function configurar() {
  PropertiesService.getScriptProperties().setProperty('HOJA_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  hoja_();
  Logger.log('Listo. Ahora publícalo como aplicación web.');
}

function doGet() {
  return respuesta_({ ok: true, servicio: 'seguimiento-conamat' });
}

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.sitio) return respuesta_({ ok: true }); // campo trampa para bots

    const t = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, 150);
    const fila = {
      curso: t(d.curso), plantel: t(d.plantel), grupo: t(d.grupo), alumno: t(d.alumno),
      acepta: t(d.acepta), telefono: String(d.telefono || '').replace(/\D/g, ''),
      horario: t(d.horario), parentesco: t(d.parentesco), responde: t(d.responde)
    };

    if (CURSOS.indexOf(fila.curso) < 0 || PLANTELES.indexOf(fila.plantel) < 0) return error_('Curso o plantel inválido.');
    if (!fila.alumno || !fila.grupo || !fila.parentesco || !fila.responde) return error_('Faltan datos.');
    if (fila.acepta !== 'Sí' && fila.acepta !== 'No') return error_('Indique Sí o No.');
    if (fila.acepta === 'Sí' && fila.telefono.length !== 10) return error_('El teléfono debe tener 10 dígitos.');
    if (fila.acepta === 'No') { fila.telefono = ''; fila.horario = ''; }

    const bloqueo = LockService.getScriptLock();
    if (!bloqueo.tryLock(20000)) return error_('Hay muchas respuestas al mismo tiempo, intente de nuevo.');
    try {
      const h = hoja_();
      h.appendRow([new Date(), fila.curso, fila.plantel, fila.grupo, fila.alumno, fila.acepta,
        fila.telefono ? "'" + fila.telefono : '', fila.horario, fila.parentesco, fila.responde,
        fila.acepta === 'Sí' ? 'Pendiente de llamar' : 'No desea llamada', '']);
    } finally {
      bloqueo.releaseLock();
    }
    return respuesta_({ ok: true });
  } catch (err) {
    console.error(err);
    return error_('No se pudo guardar. Intente de nuevo.');
  }
}

function hoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('HOJA_ID');
  const libro = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA, 0);
    h.appendRow(ENCABEZADOS);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, ENCABEZADOS.length)
      .setFontWeight('bold').setBackground('#354F8B').setFontColor('#ffffff');
    h.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('G:G').setNumberFormat('@');
    h.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente de llamar', 'Llamada hecha', 'No contestó', 'No desea llamada'], true).build());
    h.setColumnWidths(1, ENCABEZADOS.length, 150);
    h.setColumnWidth(5, 240);
    h.setColumnWidth(10, 220);
  }
  return h;
}

function error_(msg) { return respuesta_({ ok: false, error: msg }); }

function respuesta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
