// ═════════════════════════════════════════════════════════════
//  APPS SCRIPT — BBC Auditorías · Sistema Integral (Backend central)
//  Temple Colombia · Fase 2 — Cocina + Compromisos
//
//  Pega este código en: script.google.com → Nuevo proyecto
//  Luego: Implementar → Nueva implementación → Aplicación web
//    · Ejecutar como: Yo
//    · Quién tiene acceso: Cualquier usuario
//  Copia la URL /exec resultante y pégala en auditoria_cocina.html
//  (constante APPS_SCRIPT_URL, al inicio del <script>).
//
//  Cada vez que modifiques este código: Implementar → Gestionar
//  implementaciones → ✏️ (editar) → Versión: Nueva versión → Implementar.
//  Así la URL /exec NO cambia. (Solo cambia si creas una implementación
//  nueva desde cero, en vez de editar la existente.)
// ═════════════════════════════════════════════════════════════

const SECRET_TOKEN     = 'temple2026';
const MASTER_SHEET_ID  = '15hAgdCB9JWn5ke6X17gBbfxoUEDk7lEe-SCayaIwh58'; // BBC Auditorías - Sistema Integral

const TAB_AUDITORIAS  = 'Auditorias';
const TAB_HALLAZGOS   = 'Hallazgos';
const TAB_COMPROMISOS = 'Compromisos';

const HEADERS = {
  [TAB_AUDITORIAS]: [
    'id','tipo','pdv','zona','modalidad','tipoPdv','fecha','hora',
    'auditor','manager','coordinadorZona','receptor','camposAdicionales',
    'scoreTotal','scoreMax','scorePct','clasificacion','itemsEvaluados','creadoEn'
  ],
  [TAB_HALLAZGOS]: [
    'id','auditId','auditTipo','pdv','zona','coordinador',
    'categoriaId','categoriaNombre','descripcion','criticidad','tipo','creadoEn'
  ],
  [TAB_COMPROMISOS]: [
    'id','hallazgoId','auditId','pdv','zona','coordinador','auditTipo','categoriaId',
    'descripcionAccion','responsable','areaEjecutora','fechaCompromiso','estado',
    'evidencias','reincidenciaCount','fechaCierre','evidenciaCierre','validadoPor',
    'pdfEnviado','tipo','creadoEn',
    'responsableSeguimiento'
  ]
};

// ── Helpers de hoja ─────────────────────────────────────────

function getTab_(name) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = HEADERS[name];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground('#1b4332');
    hr.setFontColor('white');
    hr.setFontWeight('bold');
  }
  return sheet;
}

function rowsAsObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(r => r[0])
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

function appendObject_(sheet, headers, obj) {
  sheet.appendRow(headers.map(h => {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    if (typeof v === 'object') return JSON.stringify(v);
    return v;
  }));
}

function findRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // fila 1-based
  }
  return -1;
}

// ── Reincidencia ─────────────────────────────────────────────
// Cuenta compromisos previos de la misma categoría/PDV en los últimos 90 días
function contarReincidencia_(pdv, categoriaId, fechaISO) {
  const sheet = getTab_(TAB_COMPROMISOS);
  const rows = rowsAsObjects_(sheet);
  const fechaRef = new Date(fechaISO);
  const LIMITE_MS = 90 * 24 * 60 * 60 * 1000;
  return rows.filter(r =>
    r.pdv === pdv &&
    r.categoriaId === categoriaId &&
    (fechaRef - new Date(r.creadoEn)) <= LIMITE_MS &&
    (fechaRef - new Date(r.creadoEn)) >= 0
  ).length;
}

// ── Acciones de escritura (doPost) ────────────────────────────

function saveAuditoria_(payload) {
  const auditoria = payload.auditoria;
  const hallazgos = payload.hallazgos || [];
  const compromisos = payload.compromisos || [];

  const shAud = getTab_(TAB_AUDITORIAS);
  const idsAud = rowsAsObjects_(shAud).map(r => String(r.id));
  if (idsAud.includes(String(auditoria.id))) {
    return {status: 'exists', id: auditoria.id};
  }
  appendObject_(shAud, HEADERS[TAB_AUDITORIAS], auditoria);

  const shHal = getTab_(TAB_HALLAZGOS);
  hallazgos.forEach(h => appendObject_(shHal, HEADERS[TAB_HALLAZGOS], h));

  const shCom = getTab_(TAB_COMPROMISOS);
  compromisos.forEach(c => {
    c.reincidenciaCount = contarReincidencia_(c.pdv, c.categoriaId, c.creadoEn);
    appendObject_(shCom, HEADERS[TAB_COMPROMISOS], c);
  });

  return {status: 'saved', id: auditoria.id, hallazgos: hallazgos.length, compromisos: compromisos.length};
}

function updateCompromiso_(payload) {
  const sheet = getTab_(TAB_COMPROMISOS);
  const headers = HEADERS[TAB_COMPROMISOS];
  const row = findRowById_(sheet, payload.id);
  if (row === -1) return {error: 'Compromiso no encontrado: ' + payload.id};

  const campos = ['estado','fechaCierre','evidenciaCierre','validadoPor','pdfEnviado','responsableSeguimiento'];
  campos.forEach(campo => {
    if (payload[campo] !== undefined) {
      const col = headers.indexOf(campo) + 1;
      if (col > 0) sheet.getRange(row, col).setValue(payload[campo]);
    }
  });
  return {status: 'updated', id: payload.id};
}

// ── Acciones de lectura (doGet / doPost) ──────────────────────

function getCompromisos_(filters) {
  filters = filters || {};
  const sheet = getTab_(TAB_COMPROMISOS);
  let rows = rowsAsObjects_(sheet);

  // Auto-marcar VENCIDO si aplica (no persiste, solo para la vista)
  const hoy = new Date().toISOString().slice(0, 10);
  rows = rows.map(r => {
    if (r.estado !== 'CERRADO' && r.estado !== 'RECOMENDACIÓN' && r.fechaCompromiso &&
        String(r.fechaCompromiso).slice(0,10) < hoy && r.estado !== 'VENCIDO') {
      r.estado = 'VENCIDO';
    }
    return r;
  });

  if (filters.pdv) rows = rows.filter(r => r.pdv === filters.pdv);
  if (filters.zona) rows = rows.filter(r => r.zona === filters.zona);
  if (filters.coordinador) rows = rows.filter(r => r.coordinador === filters.coordinador);
  if (filters.auditTipo) rows = rows.filter(r => r.auditTipo === filters.auditTipo);
  if (filters.categoriaId) rows = rows.filter(r => r.categoriaId === filters.categoriaId);
  if (filters.estado && filters.estado !== 'TODOS') rows = rows.filter(r => r.estado === filters.estado);
  if (filters.desde) rows = rows.filter(r => String(r.fechaCompromiso) >= filters.desde);
  if (filters.hasta) rows = rows.filter(r => String(r.fechaCompromiso) <= filters.hasta);

  return {compromisos: rows, total: rows.length};
}

function getPdvTimeline_(pdv) {
  const auditorias = rowsAsObjects_(getTab_(TAB_AUDITORIAS)).filter(a => a.pdv === pdv);
  const hallazgos   = rowsAsObjects_(getTab_(TAB_HALLAZGOS)).filter(h => h.pdv === pdv);
  const compromisos = rowsAsObjects_(getTab_(TAB_COMPROMISOS)).filter(c => c.pdv === pdv);
  return {pdv, auditorias, hallazgos, compromisos};
}

function getAuditorias_(filters) {
  filters = filters || {};
  let rows = rowsAsObjects_(getTab_(TAB_AUDITORIAS));
  if (filters.tipo) rows = rows.filter(r => r.tipo === filters.tipo);
  return {auditorias: rows, total: rows.length};
}

// ── Enviar Acta de Compromisos por correo con PDF adjunto real ──
// Requiere habilitar el servicio avanzado "Drive API":
// Editor de Apps Script → Servicios (+) → buscar "Drive API" → Agregar
function enviarActaPdf_(payload) {
  const html = payload.actaHtml || '<p>(sin contenido)</p>';
  const destinatarios = payload.destinatarios;
  const asunto = payload.asunto || 'Acta de Compromisos';
  if (!destinatarios) return {error: 'Falta el destinatario'};

  const htmlBlob = Utilities.newBlob(html, 'text/html', 'acta.html');
  const recurso = {
    name: 'Acta temporal — ' + asunto,
    mimeType: 'application/vnd.google-apps.document'
  };
  const tempFile = Drive.Files.create(recurso, htmlBlob);
  let pdfBlob;
  try {
    pdfBlob = DriveApp.getFileById(tempFile.id).getAs('application/pdf');
    pdfBlob.setName(asunto + '.pdf');

    GmailApp.sendEmail(destinatarios, asunto,
      'Se adjunta el Acta de Compromisos en PDF.\n\nEste correo fue generado automáticamente por el Sistema de Auditorías BBC Colombia — Temple Colombia.',
      {attachments: [pdfBlob], name: 'BBC Auditorías'});
  } finally {
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  }
  return {status: 'enviado', destinatarios};
}

// ── Router ─────────────────────────────────────────────────

function checkToken_(payload) {
  return (payload.token || '') === SECRET_TOKEN;
}

function doPost(e) {
  try {
    console.log('doPost recibido, longitud=' + (e.postData ? e.postData.contents.length : 'SIN_POSTDATA'));
    const payload = JSON.parse(e.postData.contents);
    console.log('doPost action=' + payload.action);
    if (!checkToken_(payload)) return ok({error: 'Token inválido'});

    switch (payload.action) {
      case 'save_auditoria':    return ok(saveAuditoria_(payload));
      case 'update_compromiso': return ok(updateCompromiso_(payload));
      case 'get_compromisos':   return ok(getCompromisos_(payload.filters));
      case 'get_pdv_timeline':  return ok(getPdvTimeline_(payload.pdv));
      case 'get_auditorias':    return ok(getAuditorias_(payload.filters));
      case 'enviar_acta_pdf':   return ok(enviarActaPdf_(payload));
      default: return ok({error: 'Acción desconocida: ' + payload.action});
    }
  } catch (err) {
    console.log('doPost ERROR: ' + err.message + ' | ' + err.stack);
    return ok({error: err.message, stack: err.stack});
  }
}

// GET de conveniencia para lectura desde el dashboard (evita preflight CORS)
function doGet(e) {
  try {
    const p = e.parameter;
    if ((p.token || '') !== SECRET_TOKEN) return ok({error: 'Token inválido'});

    switch (p.action) {
      case 'get_compromisos':
        return ok(getCompromisos_({
          pdv: p.pdv, zona: p.zona, coordinador: p.coordinador, auditTipo: p.auditTipo,
          categoriaId: p.categoriaId, estado: p.estado, desde: p.desde, hasta: p.hasta
        }));
      case 'get_pdv_timeline': return ok(getPdvTimeline_(p.pdv));
      case 'get_auditorias':   return ok(getAuditorias_({tipo: p.tipo}));
      default: return ok({error: 'Acción desconocida: ' + p.action});
    }
  } catch (err) {
    return ok({error: err.message, stack: err.stack});
  }
}

// ── Ejecuta esta función UNA VEZ manualmente (botón "Ejecutar" arriba,
// con "autorizarPermisos" seleccionado en el desplegable) para que Google
// te muestre la pantalla de autorización de los permisos nuevos
// (Drive + Gmail). Es normal que después de aceptar no pase nada más:
// el objetivo es solo conceder el permiso.
function autorizarPermisos() {
  const blob = Utilities.newBlob('<p>prueba de permisos</p>', 'text/html', 'test.html');
  const archivo = Drive.Files.create({name: 'BBC - prueba de permisos (borrar)', mimeType: 'application/vnd.google-apps.document'}, blob);
  DriveApp.getFileById(archivo.id).setTrashed(true);
  Logger.log('Permisos de Drive concedidos correctamente');
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
