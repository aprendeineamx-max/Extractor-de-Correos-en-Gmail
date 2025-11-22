/************************************************************
 * Gmail → Google Sheets (Apps Script + Servicio avanzado Gmail v1)
 * Webhook (doPost) para lanzar desde fuera (PowerShell, Python, etc.)
 *
 * Hoja principal: "Correo"
 * Hoja índice:    "Index" (oculta; evita duplicados con MessageId)
 *
 * Columnas "Correo" (cabeceras en español):
 * DeNombre | DeCorreo | Para | Cc | Cco | Asunto | Mensaje | URL's | Fecha | Direccion | Carpeta |
 * Leido | Etiquetas | TextoPlano | HtmlBody | URLs | Adjuntos | ThreadId | MessageId
 ************************************************************/

/**********************
 * CONFIG (EDITABLE)
 **********************/
const SPREADSHEET_ID = ''; // Si el script está vinculado a la hoja, dejar vacío ''.

const FULLRESCAN_LATEST_ON_TOP = true; // true = más nuevo arriba (último mensaje arriba)

const SHEET_NAME    = 'Correo';
const INDEX_SHEET   = 'Index';
const BATCH_SIZE    = 500;             // Gmail API máx 500
const MAX_CELL      = 45000;           // límite por celda
const EXEC_LIMIT_MS = 5.5 * 60 * 1000; // ~6 min por ejecución

// TOKEN SECRETO del Webhook:
const RUN_TOKEN = '4b9f1e5e6f4f46b9a730b4f88a3d9e25c2a1c0a5afd349e0b6c7f9123a5c0d21';

// Índice (1-based) de la columna Date en la tabla (aquí está en la Nº 8)
const DATE_COL_INDEX = 8;

/**********************
 * SETUP + MENÚ
 **********************/
function setup() {
  ensureRunToken_();
  ensureSheets_();
  ensureHeaderAndFilter_();
  safeAddMenu_();
  createMinutelyTrigger_();
  backfillAll();
}

function onOpen() { safeAddMenu_(); }

function safeAddMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Correo')
      .addItem('Reimportar TODO (ignora historial)', 'fullRescanAll')
      .addItem('Reanudar reimportación (si quedó a medias)', 'fullRescanAll')
      .addSeparator()
      .addItem('Backfill histórico (sin duplicar)', 'backfillAll')
      .addItem('Incremental ahora (History)', 'fetchIncremental')
      .addItem('Forzar autorización', 'authorizeOnce')
      .addSeparator()
      .addItem('Limpiar historial + índice (no borra datos)', 'clearHistoryAndIndex')
      .addSeparator()
      .addItem('Ordenar por fecha (asc)', 'sortByDateAsc_')
      .addItem('Ordenar por fecha (desc)', 'sortByDateDesc_')
      .addToUi();
  } catch (_) { /* sin UI (webapp/standalone) */ }
}

/**********************
 * AUTORIZACIÓN + TRIGGER
 **********************/
function authorizeOnce() {
  const p = Gmail.Users.getProfile('me');
  Logger.log('Authorized as %s | historyId=%s', p.emailAddress, p.historyId);
}
function createMinutelyTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'fetchIncremental');
  if (!exists) ScriptApp.newTrigger('fetchIncremental').timeBased().everyMinutes(1).create();
}

/**********************
 * BACKFILL (sin duplicar, usa índice)
 **********************/
function backfillAll() {
  ensureSheets_();
  ensureHeaderAndFilter_();

  const index = loadIndex_();
  const me = Session.getActiveUser().getEmail();
  const start = Date.now();
  let pageToken = null;
  let totalNew = 0;

  do {
    const listResp = Gmail.Users.Messages.list('me', {
      q: 'in:anywhere',
      includeSpamTrash: true,
      maxResults: BATCH_SIZE,
      pageToken,
      fields: 'nextPageToken,messages/id'
    });

    const ids = (listResp.messages || []).map(m => m.id);
    if (!ids.length) break;

    const pending = ids.filter(id => !index.has(id));
    if (pending.length) {
      const rows = [];

      pending.forEach(id => {
        const m = Gmail.Users.Messages.get('me', id, { format: 'FULL' });

        const ts       = Number(m.internalDate || 0);
        const payload  = m.payload || {};
        const headers  = headersToObject_(payload.headers || []);
        const fromRaw  = headers['From']    || '';
        const to       = headers['To']      || '';
        const cc       = headers['Cc']      || '';
        const bcc      = headers['Bcc']     || '';
        const subject  = headers['Subject'] || '';
        const labels   = m.labelIds || [];
        const dateObj  = ts ? new Date(ts) : new Date();

        const fromParts = parseEmailAddress_(fromRaw);
        const direction = (fromParts.email || fromRaw).indexOf(me) !== -1 ? 'Enviado' : 'Recibido';
        const folder    = folderFromLabels_(labels);
        const unread    = labels.includes('UNREAD');
        const unreadStr = unread ? 'Sin Leer' : 'Leido';

        const bodies = getBodies_(m.id, payload);
        const htmlBody  = truncate_(bodies.html || '', MAX_CELL);
        const plainBody = truncate_(bodies.plain || '', MAX_CELL);

        const snippet = plainBody;
        const bodyUrlsArr = extractUrlsFromHtmlAndText_(htmlBody, plainBody);
        const bodyUrls    = bodyUrlsArr.join('\n\n');
        const attsMeta  = JSON.stringify(extractAttachmentsMeta_(payload));

        rows.push([
          fromParts.name, fromParts.email, to, cc, bcc, subject, snippet, bodyUrls, dateObj,
          direction, folder, unreadStr, labels.join(','),
          plainBody, htmlBody, bodyUrls, attsMeta,
          m.threadId, m.id
        ]);
      });

      if (rows.length) {
        rows.sort((a, b) => b[DATE_COL_INDEX - 1] - a[DATE_COL_INDEX - 1]); // Date desc
        writeRowsAtTop_(rows);
        addToIndex_(rows.map(r => r[r.length - 1])); // MessageId
        totalNew += rows.length;
      }
    }

    pageToken = listResp.nextPageToken || null;
    if (Date.now() - start > EXEC_LIMIT_MS) break;
  } while (pageToken);

  Logger.log('backfillAll() — filas nuevas: %s', totalNew);
}

/**********************
 * INCREMENTAL (History)
 **********************/
function fetchIncremental() {
  ensureSheets_();
  ensureHeaderAndFilter_();
  try {
    historyAndIngest_();
  } catch (err) {
    Logger.log('History fallback: %s', err);
    listAndIngestFallback_('in:anywhere newer_than:7d', /*incremental=*/true);
  }
}

function historyAndIngest_() {
  const props = PropertiesService.getScriptProperties();
  let lastHistoryId = props.getProperty('LAST_HIST_ID');
  const profile = Gmail.Users.getProfile('me');
  const nowHist = String(profile.historyId);

  if (!lastHistoryId) { props.setProperty('LAST_HIST_ID', nowHist); return; }

  let pageToken = null;
  const changedIds = new Set();

  do {
    const resp = Gmail.Users.History.list('me', {
      startHistoryId: lastHistoryId,
      historyTypes: ['messageAdded','labelAdded','messageDeleted','labelRemoved'],
      maxResults: 500,
      pageToken
    });

    (resp.history || []).forEach(h => {
      (h.messagesAdded || []).forEach(e => changedIds.add(e.message.id));
      (h.labelsAdded  || []).forEach(e => changedIds.add(e.message.id));
    });

    pageToken = resp.nextPageToken || null;
    if (resp.historyId) lastHistoryId = String(resp.historyId);
  } while (pageToken);

  if (changedIds.size) {
    const rows = fetchDetailsAndMapRows_(Array.from(changedIds));
    if (rows.length) writeRowsAtTop_(rows);
  }
  props.setProperty('LAST_HIST_ID', nowHist);
}

/**********************
 * FALLBACK por fecha si falla History
 **********************/
function listAndIngestFallback_(q, incremental) {
  const props = PropertiesService.getScriptProperties();
  const state = { lastTs: Number(props.getProperty('LAST_TS') || 0), pageToken: null };

  const start = Date.now();
  const index = loadIndex_();
  const toIndex = [];

  do {
    const listResp = Gmail.Users.Messages.list('me', {
      q, includeSpamTrash: true, maxResults: BATCH_SIZE, pageToken: state.pageToken
    });
    const ids = (listResp.messages || []).map(m => m.id);
    if (!ids.length) break;

    const rows = [];
    fetchDetailsAndMapRows_(ids, rows, index, toIndex, state, incremental);
    if (rows.length) writeRowsAtTop_(rows);

    state.pageToken = listResp.nextPageToken || null;
    if (Date.now() - start > EXEC_LIMIT_MS) break;
  } while (state.pageToken);

  props.setProperty('LAST_TS', String(state.lastTs || 0));
  if (toIndex.length) addToIndex_(toIndex);
}

/**********************
 * FULL RESCAN (ignora historial)
 **********************/
function fullRescanAll() {
  ensureSheets_();
  ensureHeaderAndFilter_();

  const props = PropertiesService.getScriptProperties();
  let pageToken = props.getProperty('FULLSCAN_PAGE') || null;

  if (!pageToken) {
    clearHistoryAndIndex(); // limpia estado + índice
    clearDataSheet_();      // limpia datos (mantiene encabezados)
  }

  const start = Date.now();
  const me = Session.getActiveUser().getEmail();

  do {
    const listResp = Gmail.Users.Messages.list('me', {
      q: 'in:anywhere', includeSpamTrash: true, maxResults: BATCH_SIZE, pageToken,
      fields: 'nextPageToken,messages/id'
    });

    const ids = (listResp.messages || []).map(m => m.id);
    if (!ids.length) { pageToken = null; break; }

    const rows = [];
    ids.forEach(id => {
      const m = Gmail.Users.Messages.get('me', id, { format: 'FULL' });

      const ts       = Number(m.internalDate || 0);
      const payload  = m.payload || {};
      const headers  = headersToObject_(payload.headers || []);
      const fromRaw  = headers['From']    || '';
      const to       = headers['To']      || '';
      const cc       = headers['Cc']      || '';
      const bcc      = headers['Bcc']     || '';
      const subject  = headers['Subject'] || '';
      const labels   = m.labelIds || [];
      const dateObj  = ts ? new Date(ts) : new Date();

      const fromParts = parseEmailAddress_(fromRaw);
      const direction = (fromParts.email || fromRaw).indexOf(me) !== -1 ? 'Enviado' : 'Recibido';
      const folder    = folderFromLabels_(labels);
      const unread    = labels.includes('UNREAD');
      const unreadStr = unread ? 'Sin Leer' : 'Leido';

      const bodies = getBodies_(m.id, payload);
      const htmlBody  = truncate_(bodies.html || '', MAX_CELL);
      const plainBody = truncate_(bodies.plain || '', MAX_CELL);
      const snippet   = plainBody;

      const bodyUrlsArr = extractUrlsFromHtmlAndText_(htmlBody, plainBody);
      const bodyUrls    = bodyUrlsArr.join('\n\n');

      const attsMeta  = JSON.stringify(extractAttachmentsMeta_(payload));

      rows.push([
        fromParts.name, fromParts.email, to, cc, bcc, subject, snippet, bodyUrls, dateObj,
        direction, folder, unreadStr, labels.join(','),
        plainBody, htmlBody, bodyUrls, attsMeta,
        m.threadId, m.id
      ]);
    });

    if (rows.length) {
      if (FULLRESCAN_LATEST_ON_TOP) {
        rows.sort((a, b) => b[DATE_COL_INDEX - 1] - a[DATE_COL_INDEX - 1]);
        writeRowsAtTop_(rows);
      } else {
        rows.sort((a, b) => a[DATE_COL_INDEX - 1] - b[DATE_COL_INDEX - 1]);
        appendRowsBottom_(rows);
      }
      addToIndex_(rows.map(r => r[r.length - 1]));
    }

    pageToken = listResp.nextPageToken || null;
    props.setProperty('FULLSCAN_PAGE', pageToken || '');
    if (Date.now() - start > EXEC_LIMIT_MS) break;
  } while (pageToken);

  if (!pageToken) {
    props.deleteProperty('FULLSCAN_PAGE');
    if (FULLRESCAN_LATEST_ON_TOP) sortByDateDesc_(); else sortByDateAsc_();
  }

  Logger.log('FullRescan — pageToken=%s', pageToken);
}

/**********************
 * BORRADO HISTORIAL + ÍNDICE
 **********************/
function clearHistoryAndIndex() {
  const props = PropertiesService.getScriptProperties();
  ['LAST_HIST_ID','LAST_TS','INC_PAGE','BF_PAGE','FULLSCAN_PAGE','SS_ID']
    .forEach(k => props.deleteProperty(k));

  const ss = getSS_();
  let idx = ss.getSheetByName(INDEX_SHEET);
  if (!idx) idx = ss.insertSheet(INDEX_SHEET);
  idx.showSheet();
  idx.clear();
  idx.getRange(1,1,1,1).setValues([['MessageId']]);
  idx.hideSheet();

  Logger.log('Historial (properties) + hoja Index borrados/recreados.');
}

/**********************
 * MAPEO DETALLE (compartido)
 **********************/
function fetchDetailsAndMapRows_(ids, rowsOut, index, toIndex, state, incremental) {
  const rows = rowsOut || [];
  const me = Session.getActiveUser().getEmail();

  ids.forEach(id => {
    const m = Gmail.Users.Messages.get('me', id, { format: 'FULL' });
    const ts = Number(m.internalDate || 0);

    if (state) {
      if (incremental && state.lastTs && ts <= state.lastTs) return;
      state.lastTs = Math.max(state.lastTs || 0, ts);
    }
    if (index && index.has(m.id)) return;
    if (toIndex) toIndex.push(m.id);

    const payload  = m.payload || {};
    const headers  = headersToObject_(payload.headers || []);
    const fromRaw  = headers['From']    || '';
    const to       = headers['To']      || '';
    const cc       = headers['Cc']      || '';
    const bcc      = headers['Bcc']     || '';
    const subject  = headers['Subject'] || '';
    const labels   = m.labelIds || [];
    const dateObj  = ts ? new Date(ts) : new Date();

    const fromParts = parseEmailAddress_(fromRaw);
    const direction = (fromParts.email || fromRaw).indexOf(me) !== -1 ? 'Enviado' : 'Recibido';
    const folder    = folderFromLabels_(labels);
    const unread    = labels.includes('UNREAD');
    const unreadStr = unread ? 'Sin Leer' : 'Leido';

    const bodies = getBodies_(m.id, payload);
    const htmlBody  = truncate_(bodies.html || '', MAX_CELL);
    const plainBody = truncate_(bodies.plain || '', MAX_CELL);
    const snippet   = plainBody;

    const bodyUrlsArr = extractUrlsFromHtmlAndText_(htmlBody, plainBody);
    const bodyUrls    = bodyUrlsArr.join('\n\n');

    const attsMeta  = JSON.stringify(extractAttachmentsMeta_(payload));

    rows.push([
      fromParts.name, fromParts.email, to, cc, bcc, subject, snippet, bodyUrls, dateObj,
      direction, folder, unreadStr, labels.join(','),
      plainBody, htmlBody, bodyUrls, attsMeta,
      m.threadId, m.id
    ]);
  });

  return rows;
}

/**********************
 * CUERPOS / ADJUNTOS / URLs / HELPERS
 **********************/
function getBodies_(messageId, payload) {
  let html = getTextPart_('text/html',  payload, messageId) || '';
  let text = getTextPart_('text/plain', payload, messageId) || '';

  if (!html || !text) {
    try {
      const msg = GmailApp.getMessageById(messageId);
      if (!html) html = msg.getBody();
      if (!text) text = msg.getPlainBody();
    } catch (_) {}
  }

  if (html && (!text || text.length < 120)) {
    text = htmlToText_(html);
  }

  return { html, plain: text };
}

function getTextPart_(wantedMime, part, messageId) {
  if (!part) return '';
  if (part.mimeType === wantedMime) {
    const b = part.body || {};
    if (typeof b.data === 'string' && b.data) return decodeB64Url_(b.data);
    if (b.attachmentId) return fetchAttachmentText_(messageId, b.attachmentId);
  }
  if (part.parts && part.parts.length) {
    for (const p of part.parts) {
      const v = getTextPart_(wantedMime, p, messageId);
      if (v) return v;
    }
  }
  return '';
}

function htmlToText_(html) {
  if (!html) return '';
  let s = String(html);

  s = s.replace(/<script[\s\S]*?<\/script>/gi, '')
       .replace(/<style[\s\S]*?<\/style>/gi, '');

  s = s.replace(/<(br|BR)\s*\/?>/g, '\n');
  s = s.replace(/<\/(p|div|h\d|li|tr)\s*>/gi, '\n');
  s = s.replace(/<li[^>]*>/gi, '\n• ');
  s = s.replace(/<[^>]+>/g, '');

  s = decodeHtmlEntities_(s);

  s = s.replace(/\r/g, '')
       .replace(/\t/g, ' ')
       .replace(/\u00A0/g, ' ')
       .replace(/[ \t]+\n/g, '\n')
       .replace(/\n{3,}/g, '\n\n')
       .trim();

  return s;
}

function fetchAttachmentText_(messageId, attachmentId) {
  const att = Gmail.Users.Messages.Attachments.get('me', messageId, attachmentId);
  return att && att.data ? decodeB64Url_(att.data) : '';
}

function extractAttachmentsMeta_(payload) {
  const out = [];
  (function walk(p){
    if (!p) return;
    const body = p.body || {};
    if (p.filename && body.attachmentId) {
      out.push({ filename: p.filename, mimeType: p.mimeType || '', size: body.size || null, attachmentId: body.attachmentId });
    }
    if (p.parts && p.parts.length) p.parts.forEach(walk);
  })(payload);
  return out;
}

// Devuelve {name, email} a partir de un string tipo 'Nombre <correo@dominio>'
function parseEmailAddress_(raw) {
  if (!raw) return { name: '', email: '' };
  const first = String(raw).split(',')[0].trim(); // toma solo el primer remitente si hay varios
  const m = first.match(/^(.*)<([^>]+)>$/);
  if (m) {
    const name = m[1].trim().replace(/^"|"$/g, '');
    const email = m[2].trim();
    return { name: name || email, email };
  }
  // si solo viene el correo o texto plano sin <>
  const emailLike = first.includes('@') ? first : '';
  return { name: first, email: emailLike || first };
}

function decodeB64Url_(data) {
  if (typeof data !== 'string' || !data) return '';
  let s = data.replace(/-/g, '+').replace(/_/g, '/');
  const pad = s.length % 4; if (pad) s += '='.repeat(4 - pad);
  return Utilities
    .newBlob(Utilities.base64Decode(s))
    .getDataAsString('UTF-8');
}

function decodeHtmlEntities_(text) {
  if (!text) return '';
  const map = {
    '&nbsp;': ' ', '&amp;': '&', '&lt;': '<', '&gt;': '>',
    '&quot;': '"', '&#39;': '\'', '&apos;': '\''
  };
  let s = text.replace(/&(nbsp|amp|lt|gt|quot|#39|apos);/g, m => map[m] || m);
  s = s.replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)));
  s = s.replace(/&#x([0-9A-Fa-f]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  return s;
}

function normalizeRedirect_(url) {
  try {
    const u = new URL(url);
    if ((u.hostname.endsWith('google.com') || u.hostname.endsWith('googleusercontent.com')) && u.search) {
      const q = u.searchParams.get('q') || u.searchParams.get('url');
      if (q) return decodeURIComponent(q);
    }
  } catch (_) {}
  return url;
}

function extractUrlsFromHtmlAndText_(html, plain) {
  const out = [];
  const seen = new Set();

  const push = (u) => {
    if (!u) return;
    let url = decodeHtmlEntities_(String(u)).trim();
    if (url.startsWith('//')) url = 'https:' + url;
    if (!/^https?:\/\//i.test(url)) return;
    url = normalizeRedirect_(url);
    if (!seen.has(url)) { seen.add(url); out.push(url); }
  };

  const h = html || '';
  const p = plain || '';

  const reHrefDq = /<a\b[^>]*?\bhref\s*=\s*"([^"]+)"/gi;
  const reHrefSq = /<a\b[^>]*?\bhref\s*=\s*'([^']+)'/gi;
  const reSafe   = /data-saferedirecturl\s*=\s*"([^"]+)"/gi;
  let m;
  while ((m = reHrefDq.exec(h)) !== null) push(m[1]);
  while ((m = reHrefSq.exec(h)) !== null) push(m[1]);
  while ((m = reSafe.exec(h))   !== null) push(m[1]);

  const both = `${h}\n${p}`;
  const reUrl = /\bhttps?:\/\/[^\s"'<>\)\]]+/gi;
  let mt;
  while ((mt = reUrl.exec(both)) !== null) push(mt[0]);

  return out;
}

function headersToObject_(arr) {
  const o = {};
  const canon = { from: 'From', to: 'To', cc: 'Cc', bcc: 'Bcc', subject: 'Subject' };
  (arr || []).forEach(h => {
    if (!h || !h.name) return;
    const name = String(h.name).trim();
    const val  = (h.value || '').toString();
    o[name] = val;
    const lc = name.toLowerCase();
    if (canon[lc]) o[canon[lc]] = val;
  });
  return o;
}

function folderFromLabels_(labelIds) {
  if (!labelIds || !labelIds.length) return 'Other';

  // Prioridad: Spam/Trash/Sent/Draft > Categorías de Gmail > Inbox > Other
  if (labelIds.includes('SPAM'))  return 'Spam';
  if (labelIds.includes('TRASH')) return 'Trash';
  if (labelIds.includes('SENT'))  return 'Sent';
  if (labelIds.includes('DRAFT')) return 'Draft';

  if (labelIds.includes('CATEGORY_PROMOTIONS')) return 'Promotions';
  if (labelIds.includes('CATEGORY_SOCIAL'))      return 'Social';
  if (labelIds.includes('CATEGORY_FORUMS'))      return 'Forums';

  // Tratar Updates como parte de Inbox para no mostrar “Updates” separado
  if (labelIds.includes('INBOX') || labelIds.includes('CATEGORY_UPDATES')) return 'Inbox';
  return 'Other';
}

function truncate_(s, limit) {
  if (!s) return '';
  return s.length <= limit ? s : s.substring(0, limit);
}

/**********************
 * HOJA + ÍNDICE + ESCRITURA
 **********************/
function getSS_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SS_ID');

  if (!id && SPREADSHEET_ID) {
    props.setProperty('SS_ID', SPREADSHEET_ID);
    id = SPREADSHEET_ID;
  }

  if (!id) {
    throw new Error('No hay Spreadsheet activo ni SPREADSHEET_ID definido. Abre el script desde la hoja (vinculado) o rellena SPREADSHEET_ID.');
  }
  return SpreadsheetApp.openById(id);
}

function ensureSheets_() {
  const ss = getSS_();
  if (!ss.getSheetByName(SHEET_NAME)) ss.insertSheet(SHEET_NAME);
  if (!ss.getSheetByName(INDEX_SHEET)) {
    const sh = ss.insertSheet(INDEX_SHEET);
    sh.getRange(1,1,1,1).setValues([['MessageId']]);
    sh.hideSheet();
  }
}

function ensureHeaderAndFilter_() {
  const sh = getSS_().getSheetByName(SHEET_NAME);
  const headers = [
    'DeNombre','DeCorreo','Para','Cc','Cco','Asunto','Mensaje',"URL's",'Fecha',
    'Direccion','Carpeta','Leido','Etiquetas',
    'TextoPlano','HtmlBody','URLs','Adjuntos','ThreadId','MessageId'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  if (!sh.getFilter()) sh.getRange(1,1,sh.getMaxRows(),headers.length).createFilter();
}

function writeRowsAtTop_(rows) {
  if (!rows || !rows.length) return;
  const sh = getSS_().getSheetByName(SHEET_NAME);
  sh.insertRowsAfter(1, rows.length);
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function appendRowsBottom_(rows) {
  if (!rows || !rows.length) return;
  const sh = getSS_().getSheetByName(SHEET_NAME);
  const startRow = sh.getLastRow() + 1;
  sh.insertRowsAfter(sh.getLastRow(), rows.length);
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function loadIndex_() {
  const sh = getSS_().getSheetByName(INDEX_SHEET);
  const lr = sh.getLastRow();
  const ids = lr > 1 ? sh.getRange(2,1,lr-1,1).getValues().flat() : [];
  return new Set(ids.filter(Boolean));
}
function addToIndex_(ids) {
  if (!ids || !ids.length) return;
  const sh = getSS_().getSheetByName(INDEX_SHEET);
  sh.insertRowsAfter(1, ids.length);
  sh.getRange(2,1,ids.length,1).setValues(ids.map(id=>[id]));
}

/**********************
 * ORDENACIONES
 **********************/
function sortByDateAsc_() {
  const sh = getSS_().getSheetByName(SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow > 2) sh.getRange(2,1,lastRow-1, sh.getLastColumn()).sort({column: DATE_COL_INDEX, ascending:true});
}
function sortByDateDesc_() {
  const sh = getSS_().getSheetByName(SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow > 2) sh.getRange(2,1,lastRow-1, sh.getLastColumn()).sort({column: DATE_COL_INDEX, ascending:false});
}

/**********************
 * LIMPIEZAS
 **********************/
function clearDataSheet_() {
  const sh = getSS_().getSheetByName(SHEET_NAME);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2,1,lastRow-1, sh.getLastColumn()).clearContent();
}

/**********************
 * TOKEN helpers + WEBHOOK (Web App) — doPost
 **********************/
function ensureRunToken_() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('RUN_TOKEN');
  if (existing !== RUN_TOKEN) props.setProperty('RUN_TOKEN', RUN_TOKEN);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) props.setProperty('SS_ID', ss.getId());
}
function setRunToken() {
  PropertiesService.getScriptProperties().setProperty('RUN_TOKEN', RUN_TOKEN);
  Logger.log('RUN_TOKEN establecido.');
}
function ping() { return json_({ ok: true, message: 'alive' }); }

function doPost(e) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const secret = String(props.getProperty('RUN_TOKEN') || RUN_TOKEN);
    const token  = String(getTokenFromRequest_(e) || '');

    if (!secret || token !== secret) {
      return json_({ ok:false, error:'unauthorized' });
    }

    const action = (e && e.parameter && e.parameter.action ? String(e.parameter.action) : 'fullrescan').toLowerCase();

    if (action === 'fullrescan' || action === 'fullrescanall') {
      fullRescanAll();
      return json_({ ok:true, started:'fullRescanAll' });
    } else if (action === 'backfill') {
      backfillAll();
      return json_({ ok:true, started:'backfillAll' });
    } else if (action === 'incremental') {
      fetchIncremental();
      return json_({ ok:true, started:'fetchIncremental' });
    } else if (action === 'ping') {
      return json_({ ok:true, started:'ping' });
    }

    return json_({ ok:false, error:'unknown action' });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function getTokenFromRequest_(e) {
  const headers = (e && e.headers) ? e.headers : {};
  for (const k in headers) if (k && k.toString().toLowerCase() === 'x-run-token') return headers[k];
  if (e && e.parameter && e.parameter.token) return e.parameter.token;

  if (e && e.postData && e.postData.contents) {
    try {
      if (e.postData.type && e.postData.type.indexOf('json') >= 0) {
        const obj = JSON.parse(e.postData.contents);
        if (obj && obj.token) return obj.token;
      } else {
        const parts = e.postData.contents.split('&').map(s => s.split('='));
        for (const [k, v] of parts) if ((k || '').toLowerCase() === 'token') return decodeURIComponent(v || '');
      }
    } catch (_) {}
  }
  return '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
