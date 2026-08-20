/**
 * Agenda Venues — guarda comentarios en la columna "Comentarios" de la tarea
 * (en la pestaña del área) y AVISA POR MAIL a la persona de esa área.
 *
 * Deploy: Manage deployments → lápiz → Version "New version" → Who has access: Anyone.
 * La primera vez pide permisos para editar la planilla y enviar mails: aceptá.
 */

var SHEET_ID = '12hFiAY2yGbfopdA-I8p3_2I8jj8hY9jGV0gQ7n6D740';

// ===== MODO PRUEBA =====
// Mientras esto tenga tu mail, TODOS los avisos te llegan a vos.
// Cuando termines de probar, dejalo vacío:  var TEST_MAIL = '';
var TEST_MAIL = '';

// Mail de la persona de cada área (a dónde va el aviso cuando NO estás en modo prueba)
var MAILS = {
  'Administración': 'diego.meggiolaro@dpvenues.com',   // Diego
  'Pagos':          'diego.meggiolaro@dpvenues.com',   // Diego
  'Comercial':      'lucas.adur@dpvenues.com',         // Turco
  'Legales':        'juan.sanjuan@dpvenues.com',       // Manuel
  'Obra':           'hernan.granato@dpvenues.com'      // Hernán
};

var APP_URL = 'https://agenda-venues.github.io/daleplay';

function doGet(e) {
  if (e && e.parameter && e.parameter.tarea) {
    return guardarComentario(e.parameter.area, e.parameter.tarea, e.parameter.comentario, e.parameter.autor);
  }
  return _json({ ok: true, msg: 'Agenda Venues: script activo' });
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    return guardarComentario(d.area, d.tarea, d.comentario, d.autor);
  } catch (err) { return _json({ ok: false, error: String(err) }); }
}

function guardarComentario(area, tarea, comentario, autor) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(area);
    if (!sh) return _json({ ok: false, error: 'No existe la pestaña: ' + area });

    var values = sh.getDataRange().getValues();
    var head = values[0];
    var colTarea = head.indexOf('Tarea');
    if (colTarea < 0) return _json({ ok: false, error: 'Falta la columna Tarea' });

    var colCom = head.indexOf('Comentarios');
    if (colCom < 0) { colCom = head.length; sh.getRange(1, colCom + 1).setValue('Comentarios'); }

    var tz = ss.getSpreadsheetTimeZone() || 'GMT-3';
    var fecha = Utilities.formatDate(new Date(), tz, 'dd/MM');
    var linea = '[' + (autor || '?') + ' ' + fecha + '] ' + comentario;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][colTarea]).trim() === String(tarea).trim()) {
        var prev = (colCom < values[i].length && values[i][colCom]) ? (values[i][colCom] + '\n') : '';
        sh.getRange(i + 1, colCom + 1).setValue(prev + linea);
        _avisar(area, tarea, comentario, autor);
        return _json({ ok: true });
      }
    }
    return _json({ ok: false, error: 'No se encontro la tarea: ' + tarea });
  } catch (err) { return _json({ ok: false, error: String(err) }); }
}

function _avisar(area, tarea, comentario, autor) {
  try {
    var mail = TEST_MAIL || MAILS[area];   // en modo prueba, siempre a TEST_MAIL
    if (!mail) return;
    var asunto = 'Nuevo comentario en tu tarea de ' + area;
    var cuerpo = (autor || 'Alguien') + ' comentó en la tarea "' + tarea + '" (' + area + '):\n\n' +
                 comentario + '\n\n— Ver la agenda: ' + APP_URL;
    MailApp.sendEmail(mail, asunto, cuerpo);
  } catch (err) { /* si falla el mail, el comentario igual se guardó */ }
}

function probarMail() {
  MailApp.sendEmail('valentina.ruzzante@dpvenues.com',
                    'Prueba de mail - Agenda Venues',
                    'Si te llega este mail, el envío está funcionando.');
}

function probarPost() {
  Logger.log(guardarComentario('Comercial', 'pago', 'comentario de prueba', 'Test').getContent());
}

function _json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
