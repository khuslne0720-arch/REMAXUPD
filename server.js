require('dotenv').config();
const express = require('express');
const compression = require('compression');
const multer = require('multer');
const sharp = require('sharp');
const fs = require('fs');
const { generateDocx } = require('./docxGenerator');
const { getTemplate } = require('./templates');
const path = require('path');
const xlsx = require('xlsx');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const app = express();
const CONTRACTS_FILE = path.join(__dirname, 'contracts.json');

// ── Аюулгүй байдал ──
app.use(helmet({ contentSecurityPolicy: false }));

// Analyze: 1 минутад 10 удаа
const analyzeLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  message: { error: 'Хэт олон хүсэлт. 1 минут хүлээнэ үү.' }
});

// Generate: 1 минутад 20 удаа
const generateLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 20,
  message: { error: 'Хэт олон хүсэлт. 1 минут хүлээнэ үү.' }
});

// Admin: 1 минутад 30 удаа
const adminLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 30,
  message: { error: 'Хэт олон хүсэлт.' }
});

// ── Middleware ──
app.use(compression());
app.use(express.json());
app.use(express.static('public'));

// ── Admin Session ──
const crypto = require('crypto');
const sessions = new Map();
const loginAttempts = new Map();

function generateToken() { return crypto.randomBytes(32).toString('hex'); }

function isLockedOut(ip) {
  const a = loginAttempts.get(ip);
  if (!a) return false;
  if (a.lockedUntil && Date.now() < a.lockedUntil) return true;
  if (a.lockedUntil && Date.now() >= a.lockedUntil) { loginAttempts.delete(ip); return false; }
  return false;
}

function recordFailedLogin(ip) {
  const a = loginAttempts.get(ip) || { count: 0, lockedUntil: null };
  a.count++;
  if (a.count >= 5) { a.lockedUntil = Date.now() + 15 * 60 * 1000; a.count = 0; console.log(`[SECURITY] IP хаагдлаа: ${ip}`); }
  loginAttempts.set(ip, a);
}

function requireAdmin(req, res, next) {
  const token = req.headers['x-admin-token'];
  if (!token || !sessions.has(token)) return res.status(401).json({ error: 'Нэвтрээгүй байна.' });
  const session = sessions.get(token);
  if (Date.now() - session.createdAt > 8 * 60 * 60 * 1000) { sessions.delete(token); return res.status(401).json({ error: 'Session дууссан.' }); }
  next();
}

app.post('/admin/login', adminLimiter, (req, res) => {
  const { key } = req.body || {};
  const ip = req.ip;
  if (isLockedOut(ip)) return res.status(429).json({ error: '5 удаа буруу оруулсан. 15 минут хүлээнэ үү.' });
  if (key !== process.env.ADMIN_KEY) {
    recordFailedLogin(ip);
    const a = loginAttempts.get(ip);
    const remaining = 5 - (a?.count || 0);
    console.log(`[SECURITY] Буруу нууц үг: ${ip}`);
    return res.status(403).json({ error: `Түлхүүр буруу байна. ${remaining} оролдлого үлдсэн.` });
  }
  loginAttempts.delete(ip);
  const token = generateToken();
  sessions.set(token, { ip, createdAt: Date.now() });
  console.log(`[SECURITY] Admin нэвтэрлээ: ${ip}`);
  res.json({ token });
});

app.post('/admin/logout', (req, res) => {
  const token = req.headers['x-admin-token'];
  if (token) sessions.delete(token);
  res.json({ ok: true });
});

function readContracts() {
  if (!fs.existsSync(CONTRACTS_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(CONTRACTS_FILE, 'utf-8')); }
  catch { return []; }
}

function saveContract(contractData) {
  const contracts = readContracts();
  contracts.push(contractData);
  fs.writeFileSync(CONTRACTS_FILE, JSON.stringify(contracts, null, 2), 'utf-8');
}

function writeContracts(data) {
  fs.writeFileSync(CONTRACTS_FILE, JSON.stringify(data, null, 2), 'utf-8');
}
const upload = multer({ dest: 'uploads/', limits: { fileSize: 20 * 1024 * 1024 } });

// ── Google Vision OCR (GOOGLE_VISION_KEY байвал ашиглана) ──
async function googleVisionOCR(base64Image) {
  const response = await fetch(
    `https://vision.googleapis.com/v1/images:annotate?key=${process.env.GOOGLE_VISION_KEY}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requests: [{
          image: { content: base64Image },
          features: [{ type: 'DOCUMENT_TEXT_DETECTION', maxResults: 1 }],
          imageContext: { languageHints: ['mn', 'ru'] }  // Монгол + Кирилл
        }]
      })
    }
  );
  const data = await response.json();
  if (data.error) throw new Error('Google Vision: ' + data.error.message);
  return data.responses?.[0]?.fullTextAnnotation?.text || '';
}

// ── Claude OCR (fallback) ──
async function claudeOCR(base64Image, mimeType) {
  const ocrPrompt = [
    'You are a precise OCR engine for official Mongolian government property certificates.',
    'These certificates use a special decorative calligraphic font. Read EVERY character with maximum care.',
    'The image has been preprocessed: grayscale, contrast enhanced, sharpened, binarized (black text on white).',
    '',
    'CRITICAL font confusion pairs:',
    '  Ц vs У — Ц has a small descender/tail at bottom-right. У does not.',
    '  Т vs П — Т has ONE vertical stroke. П has TWO.',
    '  Ү vs У — Ү has TWO dots above. У has none.',
    '  Ю vs И — Ю has a vertical bar on LEFT connecting two curves. И has no left bar. УЮ not УИ.',
    '  Ц vs Ч — Ц has a descender at bottom-right. Ч does not. Write ИЦ not ИЧ.',
    '  х vs т — х has two crossing diagonal strokes. т has a horizontal top bar. "хишиг" not "түнии".',
    '  я vs л — я curves right at top. л is straight diagonal.',
    '  лм cluster — never skip л before м (Тэлмэн not Тэмэн).',
    '  ц vs з, э vs о, н vs и, ү vs у, ө vs о',
    '',
    'TASK: Transcribe the ENTIRE certificate text verbatim, character by character.',
    'Output plain text only — no JSON, no markdown, no commentary.',
  ].join('\n');

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'interleaved-thinking-2025-05-14'
    },
    body: JSON.stringify({
      model: 'claude-opus-4-5',
      max_tokens: 8000,
      thinking: { type: 'enabled', budget_tokens: 5000 },
      messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Image } },
        { type: 'text', text: ocrPrompt }
      ]}]
    })
  });
  const data = await response.json();
  if (data.error) throw new Error(data.error.message);
  return (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('') || '';
}

function buildParsePrompt(rawText) {
  return [
    'You are an expert parser for Mongolian property certificate text.',
    'Below is the raw OCR text from a certificate. Extract the fields and return ONLY this JSON:',
    '{"name":"","register":"","name2":"","register2":"","ownerCount":"","address":"","area":"","cert":"","purpose":""}',
    '',
    '--- RAW TEXT ---',
    rawText,
    '--- END TEXT ---',
    '',
    'EXTRACTION RULES:',
    '',
    'ownerCount: Find "/нэг иргэний өмч/" → "1", "/хоёр иргэний өмч/" → "2", "/гурав иргэний өмч/" → "3"',
    '',
    'OWNER NAME PATTERNS:',
    'PATTERN A — contains "овогтой":',
    '  "[CLAN] овгийн [SURNAME] овогтой [FIRSTNAME] [REGISTER]"',
    '  → name = SURNAME + " " + FIRSTNAME',
    '  Example: "Иижэн овгийн Дэмбэрэлсамбуу овогтой Алтанцэцэг ЦБ56050863"',
    '  → name="Дэмбэрэлсамбуу Алтанцэцэг", register="ЦБ56050863"',
    '  CRITICAL NAME ENDING: "цэцэг" (flower) is a very common Mongolian name ending.',
    '  In this font "цэц" is often misread as "зуу" or "зуус". If you see "Алтанзуус", "Мөнхзуус", correct to "Алтанцэцэг", "Мөнхцэцэг".',
    '  The ending -зуус does NOT exist in Mongolian names. Always correct to -цэцэг.',
    '',
    'PATTERN B — "овгийн" WITHOUT "овогтой":',
    '  "[CLAN] овгийн [SURNAME] [FIRSTNAME] [REGISTER]"',
    '  → name = SURNAME + " " + FIRSTNAME  (IGNORE the CLAN word before овгийн)',
    '  Example: "Хээр Данан овгийн Баяр Тэлмэн УТ06231710"',
    '  → name="Баяр Тэлмэн", register="УТ06231710"',
    '',
    'COMMON MONGOLIAN NAMES (ҮСХ-ийн жагсаалт) — хэрэв уншсан нэр доорхтой төстэй боловч арай өөр байвал зөв нэрийг ашигла:',
    '  Эмэгтэй: Алтанцэцэг, Оюунчимэг, Болормаа, Энхтуяа, Наранцэцэг, Оюунцэцэг,',
    '    Нарантуяа, Энхжаргал, Баярмаа, Ариунаа, Солонго, Номин, Хулан, Дулмаа,',
    '    Цэцэгмаа, Мөнхзул, Оюун, Нандинцэцэг, Энхцэцэг, Баясгалан, Мөнхцэцэг,',
    '    Батцэцэг, Ундрах, Номинчимэг, Оюунбилэг, Тэгшжаргал, Зулаа, Анужин,',
    '    Энхтөгөлдөр, Буянжаргал, Гэрэлмаа, Туяа, Дөлгөөн, Мишээл, Билгүүн',
    '  Эрэгтэй: Бат-Эрдэнэ, Отгонбаяр, Батбаяр, Лхагвасүрэн, Мөнх-Эрдэнэ,',
    '    Гантулга, Ганболд, Ганбаатар, Баярсайхан, Ганзориг, Батжаргал, Батсайхан,',
    '    Тэмүүлэн, Энхболд, Мөнхбаяр, Батмөнх, Дорж, Пүрэв, Сүрэн, Цэрэн,',
    '    Жаргал, Баатар, Болд, Энхтайван, Батхүү, Мөнхбат, Ганхуяг, Батхишиг,',
    '    Эрдэнэхишиг, Пагваа, Гансүх, Түвшинтөгс, Золзаяа, Амарзаяа',
    '  ДАГАВАР ДҮРЭМ: -зуус гэж дуусах нэр байдаггүй → -цэцэг болгон засах',
    '    Жишээ: Алтанзуус→Алтанцэцэг, Наранзуус→Наранцэцэг, Оюунзуус→Оюунцэцэг',
    '',
    'MULTIPLE OWNERS RULES:',
    '  FIRST owner → name + register',
    '  SECOND owner only → name2 + register2',
    '  3rd owner and beyond → IGNORED (no field for them)',
    '  CRITICAL: name2 must contain EXACTLY ONE person\'s name (2 words max), never combine multiple owners into one field.',
    '  Example (3 owners): "АмарЗаяа овгийн Пагваа Эрдэнэхүү ХП75092868, Монгол овгийн Эрдэнэхүү Золзаяа ТЗ81082305, Гөрөөлийн овгийн Гансүх Түвшинтөгс ГЮ80122512 /гурван иргэний өмч/"',
    '  → name="Пагваа Эрдэнэхүү", register="ХП75092868", name2="Эрдэнэхүү Золзаяа", register2="ТЗ81082305", ownerCount="3"',
    '',
    'register: exactly 2 Cyrillic uppercase letters + 8 digits. Example: ЦБ56050863, УТ06231710',
    'address: full address including дүүрэг, хороо, байр, тоот',
    'area: digits + м.кв before "талбайтай". Example: "43 м.кв"',
    'cert: alphanumeric code near "гэрчилгээ олгов". Starts with Ү-, Э-, Г-, Y-, V-. Example: Ү-2204001484, V-2204155889',
    'purpose: text before "зориулалттай" or "зориулалтаар"',
    '',
    'Return ONLY the JSON object, no explanation.',
  ].join('\n');
}
app.post('/analyze', analyzeLimiter, upload.single('image'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No image uploaded' });
  try {
    let imageBuffer = fs.readFileSync(req.file.path);
    let sharpImg = sharp(imageBuffer);
    const meta = await sharpImg.metadata();

    // Зургийг OCR-д зориулж сайжруулах
    // 5MB-аас том бол л багасгана, жижиг зургийг хэвээр үлдээнэ
    if (imageBuffer.length > 5 * 1024 * 1024) {
      sharpImg = sharpImg.resize({ width: 3500, height: 3500, fit: 'inside', withoutEnlargement: true });
    }
    const cropMeta = await sharpImg.clone().metadata();
    const cw = cropMeta.width  || meta.width;
    const ch = cropMeta.height || meta.height;
    sharpImg = sharpImg.extract({
      left:   Math.round(cw * 0.06),
      top:    Math.round(ch * 0.12),
      width:  Math.round(cw * 0.88),
      height: Math.round(ch * 0.78),
    });
    // Crop хийсний дараах хэмжээ
    const cropW = Math.round(cw * 0.88);
    // Жижиг зураг (<1500px) бол 2x upscale, том зураг бол 2000px-д багасгана
    const targetW = cropW < 1500 ? cropW * 2 : Math.min(cropW, 2000);
    imageBuffer = await sharpImg
      .grayscale()
      .normalise()
      .resize({ width: targetW, kernel: sharp.kernel.cubic })
      .sharpen({ sigma: 1.5, m1: 0.5, m2: 3 })
      .jpeg({ quality: 85 })
      .toBuffer();

    // Claude API-д 5MB-аас хэтэрсэн бол чанараа бууруулж дахин compress
    if (imageBuffer.length > 5 * 1024 * 1024) {
      imageBuffer = await sharp(imageBuffer).jpeg({ quality: 70 }).toBuffer();
    }
    const base64Image = imageBuffer.toString('base64');
    const mimeType = 'image/jpeg';

    // ── Google Vision байвал ашиглах, үгүй бол Claude OCR ──
    let rawText = '';
    if (process.env.GOOGLE_VISION_KEY) {
      rawText = await googleVisionOCR(base64Image);
    } else {
      rawText = await claudeOCR(base64Image, mimeType);
    }
    // ── АЛХАМ 2: Raw текстээс JSON задлах ──
    const parsePrompt = buildParsePrompt(rawText);

    const parseResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-opus-4-5',
        max_tokens: 1024,
        messages: [{ role: 'user', content: parsePrompt }]
      })
    });
    const parseData = await parseResponse.json();
    fs.unlinkSync(req.file.path);
    if (parseData.error) throw new Error(parseData.error.message);
    const text = parseData.content?.[0]?.text || '{}';
    const extracted = JSON.parse(text.replace(/```json|```/g, '').trim());

    // ── Серверт мэдэгдэж буй алдааг засах + regex validation ──
    // Cert: У- → Ү-
    if (extracted.cert) {
      extracted.cert = extracted.cert.replace(/^У-/i, 'Ү-');
    }
    // Register regex шалгах: 2 Кирилл том үсэг + 8 цифр
    const fixRegister = (r) => {
      if (!r) return r;
      // Мэдэгдэж буй confusion pair-үүд — угтвар үсгийн засвар
      r = r.replace(/^УП/, 'УТ');   // П→Т
      r = r.replace(/^ИЧ/, 'ИЦ');   // Ч→Ц
      r = r.replace(/^УИ/, 'УЮ');   // И→Ю
      r = r.replace(/^ШЗ/, 'ТЗ');   // Ш→Т (энэ фонтод Ш, Т төстэй)
      r = r.replace(/^ШИ/, 'ТИ');
      r = r.replace(/^ШО/, 'ТО');
      r = r.replace(/^ШБ/, 'ТБ');
      // Regex: 2 том Кирилл + 8 цифр
      const m = r.match(/([А-ЯӨҮЁа-яөүё]{2})(\d{8})/u);
      return m ? (m[1].toUpperCase() + m[2]) : r;
    };
    extracted.register  = fixRegister(extracted.register);
    extracted.register2 = fixRegister(extracted.register2);

    // Area: "43 м.кв" хэлбэр шалгах
    if (extracted.area) {
      const areaM = extracted.area.match(/(\d+[\.,]?\d*)\s*м/i);
      if (areaM) extracted.area = areaM[1] + ' м.кв';
    }
    // Cert format: V-/Ү- + яг 10 цифр
    if (extracted.cert) {
      extracted.cert = extracted.cert.replace(/^У-/i, 'Ү-');
      const certM = extracted.cert.match(/([VҮЭГYvүэг])-(\d{10})/i);
      if (certM) extracted.cert = certM[1].toUpperCase() + '-' + certM[2];
    }

    // Нэрний алдаа засах — -зуус гэдэг Монгол нэрэнд байдаггүй, -цэцэг байх ёстой
    const fixName = (n) => n ? n.replace(/зуус$/i, 'цэцэг').replace(/Зуус$/i, 'цэцэг') : n;
    extracted.name  = fixName(extracted.name);
    extracted.name2 = fixName(extracted.name2);

    res.json({ success: true, data: extracted, rawText });
  } catch (err) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch (_) {}
    res.status(500).json({ error: 'Failed: ' + err.message });
  }
});
app.post('/parse-text', async (req, res) => {
  const { rawText } = req.body;
  if (!rawText) return res.status(400).json({ error: 'rawText missing' });
  try {
    const parsePrompt = buildParsePrompt(rawText);
    const parseResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-opus-4-5', max_tokens: 1024, messages: [{ role: 'user', content: parsePrompt }] })
    });
    const parseData = await parseResponse.json();
    if (parseData.error) throw new Error(parseData.error.message);
    const text = parseData.content?.[0]?.text || '{}';
    const extracted = JSON.parse(text.replace(/```json|```/g, '').trim());
    if (extracted.cert) extracted.cert = extracted.cert.replace(/^У-/i, 'Ү-');
    const fixReg = r => r ? r.replace(/^УП/, 'УТ') : r;
    extracted.register  = fixReg(extracted.register);
    extracted.register2 = fixReg(extracted.register2);
    res.json({ success: true, data: extracted });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/generate', generateLimiter, async (req, res) => {
  const { type, subtype, data } = req.body;
  if (!type || !subtype || !data) return res.status(400).json({ error: 'Missing fields' });
  try {
    const template = getTemplate(type, subtype);
    if (!template) return res.status(400).json({ error: 'Template not found' });
    const docBuffer = await generateDocx(template, data, type, subtype);
    saveContract({
      id:             Date.now().toString(),
      contractNumber: data.contractNumber ?? '',
      duration:       data.duration       ?? '',
      agent:          data.agent          ?? '',
      startDate:      data.startDate      ?? '',
      endDate:        data.endDate        ?? '',
      propertyId:     data.propertyId     ?? data.cert ?? '',
      listingType:    `${type} / ${subtype}`,
      area:           data.area           ?? '',
      address:        data.address        ?? '',
      owner:          data.name           ?? '',
      register:       data.register       ?? '',
      createdAt:      new Date().toISOString(),
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="contract_${type}_${subtype}.docx"`);
    res.send(docBuffer);
  } catch (err) {
    res.status(500).json({ error: 'Failed: ' + err.message });
  }
});

app.post('/preview', (req, res) => {
  const { type, subtype, data } = req.body;
  if (!type || !subtype || !data) return res.status(400).json({ error: 'Missing fields' });
  const template = getTemplate(type, subtype);
  if (!template) return res.status(400).json({ error: 'Template not found' });
  const { fillTemplate } = require('./templates');
  const preview = fillTemplate(template, data);
  res.json({ preview });
});
app.get('/admin/contracts', adminLimiter, requireAdmin, (req, res) => {
  res.json(readContracts());
});
// Admin: гэрээ устгах
app.delete('/admin/contracts/:id', adminLimiter, requireAdmin, (req, res) => {
  const contracts = readContracts().filter(c => c.id !== req.params.id);
  writeContracts(contracts);
  res.json({ ok: true });
});
app.get('/admin/export-excel', adminLimiter, requireAdmin, (req, res) => {
  const contracts = readContracts();
  if (contracts.length === 0)
    return res.status(400).json({ error: 'Мэдээлэл хоосон байна.' });
  const rows = contracts.map((c) => ({
    'Гэрээний дугаар':           c.contractNumber ?? '',
    'Агент':                     c.agent          ?? '',
    'Эхлэх':                     c.startDate      ?? '',
    'Дуусах':                    c.endDate        ?? '',
    'ҮХХ дугаар':                c.propertyId     ?? '',
    'Листингийн төрөл':          c.listingType    ?? '',
    'Талбайн хэмжээ':            c.area           ?? '',
    'Листингийн байршил':        c.address        ?? '',
    'ҮХХ эзэмшигчийн мэдээлэл': c.owner          ?? '',
    'РД':                        c.register       ?? '',
  }));
  const workbook  = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 14 },
    { wch: 18 }, { wch: 16 }, { wch: 30 }, { wch: 28 }, { wch: 12 },
  ];
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Гэрээнүүд');
  const buffer   = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const filename = `contracts_${new Date().toISOString().slice(0, 10)}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});
app.get('/next-contract-number', (req, res) => {
  const { type, subtype } = req.query;
  
  const config = {
    'sell_exclusive': { prefix: 'ОХ', startAt: 200 },
    'sell_standard':  { prefix: 'ЭХ', startAt: 400 },
    'rent_exclusive': { prefix: 'ОТ', startAt: 300 },
    'rent_standard':  { prefix: 'ЭТ', startAt: 200 },
  };

  const key = `${type}_${subtype}`;
  const { prefix, startAt } = config[key] || { prefix: 'ГЭ', startAt: 400 };
  const year = new Date().getFullYear().toString().slice(-2);
  const contracts = readContracts().filter(c => c.listingType === `${type} / ${subtype}`);
  const next = String(contracts.length + startAt).padStart(4, '0');
  
  res.json({ contractNumber: `${prefix}${year}/${next}` });
});
app.listen(process.env.PORT || 3000, () => console.log('Server running at http://localhost:3000'));
