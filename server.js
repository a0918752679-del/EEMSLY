require('dotenv').config();

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const express = require('express');
const cors = require('cors');
const morgan = require('morgan');
const multer = require('multer');
const XLSX = require('xlsx');
const { google } = require('googleapis');

const app = express();
const PORT = Number(process.env.PORT || 8080);
const ROOT = __dirname;
const SCHEMA = require('./columns.json');

const DATA_DIR = process.env.DATA_DIR || path.join(ROOT, 'data');
const DATA_FILE = path.join(DATA_DIR, 'cases.json');
const TEMPLATE_FILE = path.join(ROOT, 'templates', 'EEMS案件審核匯入匯出範本.xlsx');
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');

const upload = multer({ dest: UPLOAD_DIR });

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, '[]', 'utf8');

const corsOptions = process.env.ALLOWED_ORIGIN
  ? { origin: process.env.ALLOWED_ORIGIN }
  : {};
app.use(cors(corsOptions));
app.use(express.json({ limit: '8mb' }));
app.use(morgan('tiny'));
app.use(express.static(path.join(ROOT, 'public')));

function nowTaipei() {
  return new Date().toISOString();
}

function todayTaipeiDate() {
  const dtf = new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Asia/Taipei',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  return dtf.format(new Date());
}

function readCases() {
  try {
    const raw = fs.readFileSync(DATA_FILE, 'utf8');
    return JSON.parse(raw || '[]');
  } catch (err) {
    console.error('readCases failed:', err);
    return [];
  }
}

function writeCases(cases) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(refreshRepeatFlags(cases), null, 2), 'utf8');
}

function addLog(item, action, by, note = '') {
  item.logs = Array.isArray(item.logs) ? item.logs : [];
  item.logs.unshift({ at: nowTaipei(), action, by: by || 'system', note });
  item.updatedAt = nowTaipei();
}

function normalizeText(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).trim();
}

function parseBool(value) {
  if (typeof value === 'boolean') return value;
  const v = normalizeText(value).toLowerCase();
  return ['true', '1', 'yes', 'y', '是', '核可', '已核可', 'v', '✓', '勾選'].includes(v);
}

function isRepeatSource(source = {}) {
  return parseBool(source['是否累犯']) || parseBool(source['1年內再次違反']) || normalizeText(source['前次告發']);
}

function caseNeedsEscalation(item) {
  return Boolean(item?.flags?.repeatOverLimit || item?.flags?.escalated || isRepeatSource(item?.source));
}

function refreshRepeatFlags(cases) {
  const plateCounts = new Map();
  cases.forEach((item) => {
    const plate = normalizeText(item?.source?.['車號']).toUpperCase();
    if (plate) plateCounts.set(plate, (plateCounts.get(plate) || 0) + 1);
  });

  cases.forEach((item) => {
    const plate = normalizeText(item?.source?.['車號']).toUpperCase();
    const count = plate ? (plateCounts.get(plate) || 0) : 0;
    const sourceRepeat = isRepeatSource(item?.source);
    item.flags = item.flags || {};
    item.flags.plateRepeatCount = count;
    item.flags.repeatOverLimit = Boolean(plate && count > 1) || sourceRepeat;
    item.flags.escalated = item.flags.repeatOverLimit;
    item.flags.repeatNote = item.flags.repeatOverLimit
      ? `車號重複超標${count > 1 ? count + '件' : ''}，註明加重`
      : '';
  });
  return cases;
}

function findCaseIndexById(cases, id) {
  return cases.findIndex((c) => c.id === id);
}

function requireDeleteAdmin(req, res, next) {
  const required = process.env.ADMIN_DELETE_PASSCODE || '69677323';
  const provided = req.header('x-admin-delete-passcode') || req.body?.adminPassword || req.query.adminPassword;
  if (provided === required) return next();
  return res.status(403).json({ error: '刪除案件需要管理權限密碼。' });
}

function cleanObjectByHeaders(row, headers) {
  const out = {};
  headers.forEach((h) => {
    out[h] = normalizeText(row[h]);
  });
  return out;
}

function buildCaseId(type, row, source) {
  const prefix = type === '告發' ? 'A' : 'I';
  const explicit = normalizeText(row['案件編號']);
  if (explicit) return explicit;
  const preferred = type === '告發'
    ? normalizeText(row['稽查編號'] || row['告發單號'])
    : '';
  if (preferred) return `${prefix}-${preferred}`;
  const seed = [
    type,
    source['車號'],
    source['量測日期'],
    source['稽查時間'],
    source['量測位置地點']
  ].filter(Boolean).join('|');
  const hash = crypto.createHash('sha1').update(seed || crypto.randomUUID()).digest('hex').slice(0, 10).toUpperCase();
  return `${prefix}-${hash}`;
}

function statusFromRow(row, type) {
  const explicit = normalizeText(row['案件狀態']);
  if (explicit) return explicit;
  if (parseBool(row['長官核可'])) return '行政處理中';
  if (normalizeText(row['退回原因'])) return '退回修正';
  return '待長官審核';
}

function caseFromFlatRow(type, row, options = {}) {
  if (!SCHEMA[type]) throw new Error(`不支援的案件類型：${type}`);
  const source = cleanObjectByHeaders(row, SCHEMA[type].sourceHeaders);
  const explicitId = normalizeText(row['案件編號']);
  if (options.requireExplicitId && !explicitId) {
    const err = new Error('Google Sheet 同步略過：未填案件編號');
    err.code = 'MISSING_CASE_ID';
    throw err;
  }
  const id = buildCaseId(type, row, source);
  const chiefApproved = parseBool(row['長官核可']);
  const status = statusFromRow(row, type);
  const item = {
    id,
    type,
    status,
    source,
    review: {
      consultantChecked: parseBool(row['顧問初核']) || parseBool(row['承辦複核']) || parseBool(row['備註']),
      consultantReviewer: normalizeText(row['承辦人員']) || '聯韻初核',
      submittedAt: normalizeText(row['送審日期']) || todayTaipeiDate(),
      chiefApproved,
      chiefReviewer: normalizeText(row['長官姓名']),
      chiefReviewedAt: normalizeText(row['核可日期']),
      rejectedReason: normalizeText(row['退回原因'])
    },
    admin: {
      stage: normalizeText(row['行政流程階段']) || (chiefApproved ? '待製作公文' : '尚未啟動'),
      officialDocumentNo: normalizeText(row['公文文號']) || normalizeText(row['公文案號']),
      officialDocumentDate: normalizeText(row['發文日期']) || normalizeText(row['公文日期']),
      dispositionNo: normalizeText(row['裁處字號']) || normalizeText(row['裁處書號']),
      penaltyAmount: normalizeText(row['裁罰金額']) || normalizeText(row['金額']),
      deliveryDate: normalizeText(row['送達日期']) || normalizeText(row['送達狀態']),
      paymentStatus: normalizeText(row['繳款狀態']) || '尚未繳款',
      inspectionDueDate: normalizeText(row['通檢期限']),
      inspectionStatus: normalizeText(row['到檢狀態']) || (type === '通檢' ? '通知待發文' : '未啟動'),
      closeDate: normalizeText(row['結案日期']),
      notes: normalizeText(row['行政備註'])
    },
    flags: {
      repeatOverLimit: isRepeatSource(source),
      escalated: isRepeatSource(source),
      repeatNote: isRepeatSource(source) ? '車號重複超標，註明加重' : '',
      plateRepeatCount: 0
    },
    logs: [],
    createdAt: nowTaipei(),
    updatedAt: nowTaipei()
  };
  addLog(item, '匯入案件', 'system', `${type}固定格式匯入`);
  return item;
}

function toFlatRow(item) {
  const sourceHeaders = SCHEMA[item.type].sourceHeaders;
  const row = {};
  sourceHeaders.forEach((h) => { row[h] = normalizeText(item.source && item.source[h]); });
  row['案件編號'] = item.id;
  row['案件類型'] = item.type;
  row['案件狀態'] = item.status;
  row['顧問初核'] = item.review?.consultantChecked ? 'TRUE' : '';
  row['承辦人員'] = item.review?.consultantReviewer || '';
  row['送審日期'] = item.review?.submittedAt || '';
  row['長官核可'] = item.review?.chiefApproved ? 'TRUE' : '';
  row['長官姓名'] = item.review?.chiefReviewer || '';
  row['核可日期'] = item.review?.chiefReviewedAt || '';
  row['退回原因'] = item.review?.rejectedReason || '';
  row['行政流程階段'] = item.admin?.stage || '';
  row['公文文號'] = item.admin?.officialDocumentNo || '';
  row['發文日期'] = item.admin?.officialDocumentDate || '';
  row['裁處字號'] = formatDispositionNo(item.admin?.dispositionNo || '');
  row['裁罰金額'] = item.admin?.penaltyAmount || '';
  row['送達日期'] = item.admin?.deliveryDate || item.admin?.deliveryStatus || '';
  row['繳款狀態'] = item.admin?.paymentStatus || '';
  row['通檢期限'] = item.admin?.inspectionDueDate || '';
  row['到檢狀態'] = item.admin?.inspectionStatus || '';
  row['結案日期'] = item.admin?.closeDate || '';
  row['行政備註'] = item.admin?.notes || '';
  row['重複超標註記'] = caseNeedsEscalation(item) ? (item.flags?.repeatNote || '車號重複超標，註明加重') : '';
  row['更新時間'] = item.updatedAt || '';
  return row;
}

function exportHeaders(type) {
  return [...SCHEMA[type].sourceHeaders, ...SCHEMA.workflowHeaders];
}

function rowsToSheetValues(type, cases) {
  refreshRepeatFlags(cases);
  const headers = exportHeaders(type);
  const dataRows = cases
    .filter((c) => c.type === type)
    .map((c) => {
      const row = toFlatRow(c);
      return headers.map((h) => row[h] ?? '');
    });
  return [headers, ...dataRows];
}

function makeWorkbookForCases(type, cases) {
  refreshRepeatFlags(cases);
  const wb = XLSX.utils.book_new();
  if (type === '全部') {
    ['告發', '通檢'].forEach((t) => {
      const ws = XLSX.utils.aoa_to_sheet(rowsToSheetValues(t, cases));
      XLSX.utils.book_append_sheet(wb, ws, t);
    });
  } else {
    const ws = XLSX.utils.aoa_to_sheet(rowsToSheetValues(type, cases));
    XLSX.utils.book_append_sheet(wb, ws, type);
  }
  const logRows = [['案件編號', '時間', '動作', '執行者', '備註']];
  cases.forEach((c) => {
    (c.logs || []).forEach((log) => logRows.push([c.id, log.at, log.action, log.by, log.note]));
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(logRows), '案件追蹤紀錄');
  return wb;
}

function readUploadedWorkbook(filePath, preferredType) {
  const wb = XLSX.readFile(filePath, { cellDates: true, raw: false, defval: '' });
  const sheetName = wb.SheetNames.includes(preferredType) ? preferredType : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
  return rows;
}

function upsertCases(existing, incoming) {
  const map = new Map(existing.map((item) => [item.id, item]));
  let inserted = 0;
  let updated = 0;
  incoming.forEach((item) => {
    if (map.has(item.id)) {
      const old = map.get(item.id);
      const merged = {
        ...old,
        ...item,
        logs: [...(item.logs || []), ...(old.logs || [])],
        createdAt: old.createdAt || item.createdAt,
        updatedAt: nowTaipei()
      };
      map.set(item.id, merged);
      updated += 1;
    } else {
      map.set(item.id, item);
      inserted += 1;
    }
  });
  return { cases: refreshRepeatFlags(Array.from(map.values())), inserted, updated };
}

function filterCases(cases, query) {
  let out = cases;
  if (query.type && query.type !== '全部') out = out.filter((c) => c.type === query.type);
  if (query.status && query.status !== '全部') out = out.filter((c) => c.status === query.status);
  if (query.q) {
    const q = String(query.q).toLowerCase();
    out = out.filter((c) => JSON.stringify(c).toLowerCase().includes(q));
  }
  return out;
}

function stats(cases) {
  const byStatus = {};
  const byType = {};
  SCHEMA.statusOptions.forEach((s) => { byStatus[s] = 0; });
  ['告發', '通檢'].forEach((t) => { byType[t] = 0; });
  cases.forEach((c) => {
    byStatus[c.status] = (byStatus[c.status] || 0) + 1;
    byType[c.type] = (byType[c.type] || 0) + 1;
  });
  return {
    total: cases.length,
    byStatus,
    byType,
    pendingReview: cases.filter((c) => c.status === '待長官審核').length,
    approved: cases.filter((c) => c.review?.chiefApproved).length,
    adminProcessing: cases.filter((c) => c.status === '行政處理中').length,
    closed: cases.filter((c) => c.status === '已結案').length
  };
}

function authGuard(req, res, next) {
  const required = process.env.APP_PASSCODE;
  if (!required) return next();
  const provided = req.header('x-app-passcode') || req.query.passcode;
  if (provided === required) return next();
  return res.status(401).json({ error: '未授權：請輸入內部驗證碼。' });
}

app.use('/api', authGuard);

async function getSheetsClient() {
  const sheetId = process.env.GOOGLE_SHEET_ID;
  if (!sheetId) throw new Error('尚未設定 GOOGLE_SHEET_ID');

  let credentials = null;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64) {
    const raw = Buffer.from(process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64, 'base64').toString('utf8');
    credentials = JSON.parse(raw);
  } else if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } else {
    throw new Error('尚未設定 GOOGLE_SERVICE_ACCOUNT_JSON 或 GOOGLE_SERVICE_ACCOUNT_JSON_BASE64');
  }

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  const sheets = google.sheets({ version: 'v4', auth });
  return { sheets, sheetId };
}

async function ensureGoogleSheetsTabs(sheets, sheetId) {
  const desired = ['告發', '通檢', '案件追蹤紀錄'];
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const existing = new Set((meta.data.sheets || []).map((s) => s.properties.title));
  const requests = desired
    .filter((title) => !existing.has(title))
    .map((title) => ({ addSheet: { properties: { title } } }));
  if (requests.length) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: sheetId, requestBody: { requests } });
  }
}

app.get('/health', async (req, res) => {
  res.json({ ok: true, service: 'ntpc-eems-case-review-system', status: 'running', time: nowTaipei() });
});

app.get('/api/health', async (req, res) => {
  const payload = {
    ok: true,
    service: 'EEMS聲音照相案件審核系統',
    organization: '聯韻聲學環保顧問股份有限公司',
    time: nowTaipei(),
    googleSheets: { configured: Boolean(process.env.GOOGLE_SHEET_ID), connected: false, message: '' }
  };
  try {
    const { sheets, sheetId } = await getSheetsClient();
    await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    payload.googleSheets.connected = true;
    payload.googleSheets.message = 'Google Sheet 連線成功';
  } catch (err) {
    payload.googleSheets.message = err.message;
  }
  res.json(payload);
});

app.get('/api/schema', (req, res) => {
  res.json(SCHEMA);
});

app.get('/api/cases', (req, res) => {
  const all = refreshRepeatFlags(readCases());
  const list = filterCases(all, req.query);
  res.json({ items: list, stats: stats(all) });
});

app.get('/api/cases/:id', (req, res) => {
  const item = refreshRepeatFlags(readCases()).find((c) => c.id === req.params.id);
  if (!item) return res.status(404).json({ error: '找不到案件' });
  res.json(item);
});

app.post('/api/import/:type', upload.single('file'), (req, res) => {
  const type = req.params.type;
  if (!SCHEMA[type]) return res.status(400).json({ error: '案件類型僅支援告發或通檢' });
  if (!req.file) return res.status(400).json({ error: '請上傳 Excel 檔案' });

  try {
    const rows = readUploadedWorkbook(req.file.path, type);
    const incoming = rows
      .filter((row) => normalizeText(row['車號']) || normalizeText(row['稽查編號']) || normalizeText(row['告發單號']))
      .map((row) => caseFromFlatRow(type, row));

    const result = upsertCases(readCases(), incoming);
    writeCases(result.cases);
    res.json({ ok: true, type, totalRows: rows.length, imported: incoming.length, inserted: result.inserted, updated: result.updated });
  } catch (err) {
    res.status(500).json({ error: err.message });
  } finally {
    try { fs.unlinkSync(req.file.path); } catch (_) {}
  }
});

app.get('/api/export/:type', (req, res) => {
  const type = req.params.type;
  if (!['告發', '通檢', '全部'].includes(type)) return res.status(400).json({ error: '匯出類型僅支援告發、通檢或全部' });
  const wb = makeWorkbookForCases(type, readCases());
  const fileName = encodeURIComponent(`EEMS案件審核匯出_${type}_${todayTaipeiDate()}.xlsx`);
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fileName}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.end(buffer);
});

app.get('/api/template', (req, res) => {
  if (fs.existsSync(TEMPLATE_FILE)) return res.download(TEMPLATE_FILE);
  const wb = makeWorkbookForCases('全部', []);
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent('EEMS案件審核匯入匯出範本.xlsx')}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.end(buffer);
});

app.patch('/api/cases/:id/update', (req, res) => {
  const cases = readCases();
  const index = findCaseIndexById(cases, req.params.id);
  if (index < 0) return res.status(404).json({ error: '找不到案件' });
  const item = cases[index];
  const originalId = item.id;
  const newId = normalizeText(req.body.id || req.body.newId || item.id);
  if (!newId) return res.status(400).json({ error: '案件編號不可空白' });
  if (newId !== originalId && cases.some((c) => c.id === newId)) return res.status(409).json({ error: '案件編號已存在，無法修改' });

  const nextType = normalizeText(req.body.type || item.type);
  if (!SCHEMA[nextType]) return res.status(400).json({ error: '案件類型僅支援告發或通檢' });
  item.id = newId;
  item.type = nextType;
  if (Object.prototype.hasOwnProperty.call(req.body, 'status')) item.status = normalizeText(req.body.status) || item.status;
  item.source = item.source || {};
  const sourcePatch = req.body.source && typeof req.body.source === 'object' ? req.body.source : {};
  Object.entries(sourcePatch).forEach(([key, value]) => {
    if (SCHEMA[nextType].sourceHeaders.includes(key)) item.source[key] = normalizeText(value);
  });
  item.flags = item.flags || {};
  if (Object.prototype.hasOwnProperty.call(req.body, 'repeatOverLimit')) item.flags.repeatOverLimit = parseBool(req.body.repeatOverLimit);
  if (Object.prototype.hasOwnProperty.call(req.body, 'escalated')) item.flags.escalated = parseBool(req.body.escalated);
  if (Object.prototype.hasOwnProperty.call(req.body, 'repeatNote')) item.flags.repeatNote = normalizeText(req.body.repeatNote);
  if (item.flags.repeatOverLimit || item.flags.escalated) item.flags.repeatNote = item.flags.repeatNote || '車號重複超標，註明加重';
  addLog(item, '修改案件', normalizeText(req.body.operator) || '承辦人員', newId !== originalId ? `案件編號 ${originalId} 改為 ${newId}` : '更新案件資料');
  writeCases(cases);
  res.json({ ok: true, item });
});

app.delete('/api/cases/:id', requireDeleteAdmin, (req, res) => {
  const cases = readCases();
  const index = findCaseIndexById(cases, req.params.id);
  if (index < 0) return res.status(404).json({ error: '找不到案件' });
  const [removed] = cases.splice(index, 1);
  writeCases(cases);
  res.json({ ok: true, deleted: removed.id });
});

app.patch('/api/cases/:id/review', (req, res) => {
  const cases = readCases();
  const item = cases.find((c) => c.id === req.params.id);
  if (!item) return res.status(404).json({ error: '找不到案件' });

  const approved = parseBool(req.body.approved);
  const reviewer = normalizeText(req.body.reviewer) || '局內長官';
  const rejectedReason = normalizeText(req.body.rejectedReason || req.body.reason);

  item.review = item.review || {};
  item.review.chiefApproved = approved;
  item.review.chiefReviewer = reviewer;
  item.review.chiefReviewedAt = todayTaipeiDate();
  item.review.rejectedReason = approved ? '' : rejectedReason;

  item.admin = item.admin || {};
  if (approved) {
    item.status = '行政處理中';
    item.admin.stage = item.admin.stage && item.admin.stage !== '尚未啟動' ? item.admin.stage : '待製作公文';
    if (item.type === '通檢' && (!item.admin.inspectionStatus || item.admin.inspectionStatus === '未啟動')) {
      item.admin.inspectionStatus = '通知待發文';
    }
    addLog(item, '長官核可', reviewer, '案件進入後續行政流程');
  } else {
    item.status = '退回修正';
    item.admin.stage = item.admin.stage || '尚未啟動';
    addLog(item, '退回修正', reviewer, rejectedReason);
  }
  writeCases(cases);
  res.json({ ok: true, item });
});

app.post('/api/cases/bulk-review', (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids : [];
  const approved = parseBool(req.body.approved);
  const reviewer = normalizeText(req.body.reviewer) || '局內長官';
  const rejectedReason = normalizeText(req.body.rejectedReason || req.body.reason);
  const cases = readCases();
  let changed = 0;

  cases.forEach((item) => {
    if (!ids.includes(item.id)) return;
    item.review = item.review || {};
    item.admin = item.admin || {};
    item.review.chiefApproved = approved;
    item.review.chiefReviewer = reviewer;
    item.review.chiefReviewedAt = todayTaipeiDate();
    item.review.rejectedReason = approved ? '' : rejectedReason;

    if (approved) {
      item.status = '行政處理中';
      item.admin.stage = item.admin.stage && item.admin.stage !== '尚未啟動' ? item.admin.stage : '待製作公文';
      if (item.type === '通檢' && (!item.admin.inspectionStatus || item.admin.inspectionStatus === '未啟動')) item.admin.inspectionStatus = '通知待發文';
      addLog(item, '批次長官核可', reviewer, `批次核可 ${ids.length} 件`);
    } else {
      item.status = '退回修正';
      addLog(item, '批次退回修正', reviewer, rejectedReason);
    }
    changed += 1;
  });

  writeCases(cases);
  res.json({ ok: true, changed });
});

app.patch('/api/cases/:id/admin', (req, res) => {
  const cases = readCases();
  const item = cases.find((c) => c.id === req.params.id);
  if (!item) return res.status(404).json({ error: '找不到案件' });

  if (Object.prototype.hasOwnProperty.call(req.body, 'officialDocumentNo')) {
    const d = digitsOnly(req.body.officialDocumentNo);
    if (d && d.length !== 10) return res.status(400).json({ error: '公文文號需為10碼數字' });
    req.body.officialDocumentNo = d;
  }
  if (Object.prototype.hasOwnProperty.call(req.body, 'dispositionNo')) {
    const d = digitsOnly(req.body.dispositionNo);
    if (d && d.length !== 11) return res.status(400).json({ error: '裁處字號需為11碼數字' });
    req.body.dispositionNo = d;
  }

  const allowed = [
    'stage', 'officialDocumentNo', 'officialDocumentDate', 'dispositionNo',
    'penaltyAmount', 'deliveryDate', 'paymentStatus',
    'inspectionDueDate', 'inspectionStatus', 'closeDate', 'notes'
  ];

  item.admin = item.admin || {};
  allowed.forEach((key) => {
    if (Object.prototype.hasOwnProperty.call(req.body, key)) item.admin[key] = normalizeText(req.body[key]);
  });

  if (item.admin.closeDate || item.admin.paymentStatus === '繳款結案' || item.admin.inspectionStatus === '結案') {
    item.status = '已結案';
    item.admin.stage = item.admin.stage || '已結案';
  } else if (item.review?.chiefApproved) {
    item.status = '行政處理中';
  }

  addLog(item, '更新行政流程', '局內長官', item.admin.notes || '');
  writeCases(cases);
  res.json({ ok: true, item });
});

app.post('/api/sync/pull', async (req, res) => {
  try {
    const { sheets, sheetId } = await getSheetsClient();
    await ensureGoogleSheetsTabs(sheets, sheetId);
    const incoming = [];
    let skippedMissingCaseId = 0;

    for (const type of ['告發', '通檢']) {
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `${type}!A:AZ`
      });
      const values = result.data.values || [];
      if (!values.length) continue;
      const headers = values[0];
      values.slice(1).forEach((arr) => {
        const row = {};
        headers.forEach((h, idx) => { row[h] = arr[idx] || ''; });
        if (normalizeText(row['車號']) || normalizeText(row['稽查編號']) || normalizeText(row['告發單號'])) {
          if (!normalizeText(row['案件編號'])) {
            skippedMissingCaseId += 1;
            return;
          }
          incoming.push(caseFromFlatRow(type, row, { requireExplicitId: true }));
        }
      });
    }

    const result = upsertCases(readCases(), incoming);
    writeCases(result.cases);
    res.json({ ok: true, pulled: incoming.length, inserted: result.inserted, updated: result.updated, skippedMissingCaseId });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/sync/push', async (req, res) => {
  try {
    const { sheets, sheetId } = await getSheetsClient();
    await ensureGoogleSheetsTabs(sheets, sheetId);
    const cases = readCases();

    for (const type of ['告發', '通檢']) {
      await sheets.spreadsheets.values.clear({ spreadsheetId: sheetId, range: `${type}!A:AZ` });
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `${type}!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: rowsToSheetValues(type, cases) }
      });
    }

    const logRows = [['案件編號', '案件類型', '時間', '動作', '執行者', '備註']];
    cases.forEach((c) => (c.logs || []).forEach((log) => logRows.push([c.id, c.type, log.at, log.action, log.by, log.note])));
    await sheets.spreadsheets.values.clear({ spreadsheetId: sheetId, range: `案件追蹤紀錄!A:Z` });
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `案件追蹤紀錄!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: logRows }
    });

    res.json({ ok: true, pushed: cases.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(ROOT, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`EEMS case review system listening on ${PORT}`);
});
