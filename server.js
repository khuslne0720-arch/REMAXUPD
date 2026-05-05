require('dotenv').config();
const express = require('express');
const compression = require('compression');
const multer = require('multer');
const sharp = require('sharp');
const fs = require('fs');
const crypto = require('crypto');
const { generateDocx } = require('./docxGenerator');
const { getTemplate } = require('./templates');
const path = require('path');
const xlsx = require('xlsx');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const app = express();

// ── Env validation ──
const REQUIRED_ENV = ['ADMIN_KEY', 'ANTHROPIC_API_KEY', 'SITE_KEY'];
const missingEnv = REQUIRED_ENV.filter(k => !process.env[k]);
if (missingEnv.length) {
  console.error('[FATAL] Дутуу environment variables:', missingEnv.join(', '));
  process.exit(1);
}
if (process.env.ADMIN_KEY.length < 12) {
  console.error('[FATAL] ADMIN_KEY хамгийн багадаа 12 тэмдэгт байх ёстой!');
  process.exit(1);
}

// ── Security log ──
const SECURITY_LOG = process.env.SECURITY_LOG || path.join(__dirname, 'security.log');
function securityLog(event, details) {
  const line = `${new Date().toISOString()} [${event}] ${JSON.stringify(details)}\n`;
  fs.appendFile(SECURITY_LOG, line, () => {});
  console.log(`[SECURITY] ${event}`, details);
}

// ── Site-key middleware ──
function requireSiteKey(req, res, next) {
  if (req.headers['x-site-key'] !== process.env.SITE_KEY) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch (_) {}
    securityLog('SITE_KEY_FAIL', { ip: req.ip, path: req.path });
    return res.status(403).json({ error: 'Forbidden' });
  }
  next();
}

const CONTRACTS_FILE = process.env.CONTRACTS_FILE || path.join(__dirname, 'contracts.json');
const CONTRACTS_DIR = path.dirname(CONTRACTS_FILE);
if (!fs.existsSync(CONTRACTS_DIR)) fs.mkdirSync(CONTRACTS_DIR, { recursive: true });

// ── AES-256-CBC шифрлэлт ──
const ENC_KEY     = crypto.createHash('sha256').update(process.env.ENCRYPT_KEY || process.env.ADMIN_KEY || 'default').digest();
const OLD_ENC_KEY = crypto.createHash('sha256').update(process.env.ADMIN_KEY   || 'default').digest();
const IV_LEN = 16;

function encrypt(text) {
  const iv = crypto.randomBytes(IV_LEN);
  const cipher = crypto.createCipheriv('aes-256-cbc', ENC_KEY, iv);
  const enc = Buffer.concat([cipher.update(text, 'utf8'), cipher.final()]);
  return iv.toString('hex') + ':' + enc.toString('hex');
}

function tryDecryptWith(key, text) {
  const [ivHex, encHex] = text.split(':');
  const decipher = crypto.createDecipheriv('aes-256-cbc', key, Buffer.from(ivHex, 'hex'));
  return Buffer.concat([decipher.update(Buffer.from(encHex, 'hex')), decipher.final()]).toString('utf8');
}

function decrypt(text) {
  try { return tryDecryptWith(ENC_KEY, text); } catch (_) {}
  try {
    const plain = tryDecryptWith(OLD_ENC_KEY, text);
    fs.writeFileSync(CONTRACTS_FILE, encrypt(plain), 'utf-8');
    return plain;
  } catch (_) {}
  throw new Error('Decrypt failed');
}

// ── Аюулгүй байдал ──
app.set('trust proxy', 1);
app.use(helmet({
  contentSecurityPolicy: false,
  hsts: { maxAge: 31536000, includeSubDomains: true },
  frameguard: { action: 'deny' },
  noSniff: true,
}));

// Analyze: 1 минутад 10 удаа
const analyzeLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  validate: { xForwardedForHeader: false },
  message: { error: 'Хэт олон хүсэлт. 1 минут хүлээнэ үү.' }
});

// Generate: 1 минутад 20 удаа
const generateLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 20,
  validate: { xForwardedForHeader: false },
  message: { error: 'Хэт олон хүсэлт. 1 минут хүлээнэ үү.' }
});

// Admin: 1 минутад 30 удаа
const adminLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 30,
  validate: { xForwardedForHeader: false },
  message: { error: 'Хэт олон хүсэлт.' }
});

// ── Middleware ──
app.use(compression());
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true, limit: '2mb' }));
app.use(express.static('public'));

// ── Admin Session ──
const SESSION_TTL_MS = 8 * 60 * 60 * 1000;
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
  if (Date.now() - session.createdAt > SESSION_TTL_MS) { sessions.delete(token); return res.status(401).json({ error: 'Session дууссан.' }); }
  next();
}

app.post('/admin/login', adminLimiter, (req, res) => {
  const { key } = req.body || {};
  const ip = req.ip;
  if (isLockedOut(ip)) {
    securityLog('LOGIN_LOCKED', { ip });
    return res.status(429).json({ error: '5 удаа буруу оруулсан. 15 минут хүлээнэ үү.' });
  }
  if (key !== process.env.ADMIN_KEY) {
    recordFailedLogin(ip);
    const a = loginAttempts.get(ip);
    const remaining = 5 - (a?.count || 0);
    securityLog('LOGIN_FAIL', { ip, attemptsLeft: remaining });
    return res.status(403).json({ error: `Түлхүүр буруу байна. ${remaining} оролдлого үлдсэн.` });
  }
  loginAttempts.delete(ip);
  const token = generateToken();
  sessions.set(token, { ip, createdAt: Date.now() });
  securityLog('LOGIN_SUCCESS', { ip });
  res.json({ token });
});

app.post('/admin/logout', (req, res) => {
  const token = req.headers['x-admin-token'];
  if (token) sessions.delete(token);
  res.json({ ok: true });
});

function readContracts() {
  if (!fs.existsSync(CONTRACTS_FILE)) return [];
  try {
    const raw = fs.readFileSync(CONTRACTS_FILE, 'utf-8').trim();
    if (!raw) return [];
    // Хуучин шифрлэгдээгүй JSON → автомат migrate
    if (raw.startsWith('[') || raw.startsWith('{')) {
      const data = JSON.parse(raw);
      try { fs.writeFileSync(CONTRACTS_FILE, encrypt(JSON.stringify(data)), 'utf-8'); } catch (_) {}
      return data;
    }
    return JSON.parse(decrypt(raw));
  } catch { return []; }
}

function saveContract(contractData) {
  const contracts = readContracts();
  contracts.push(contractData);
  fs.writeFileSync(CONTRACTS_FILE, encrypt(JSON.stringify(contracts)), 'utf-8');
}

function writeContracts(data) {
  fs.writeFileSync(CONTRACTS_FILE, encrypt(JSON.stringify(data)), 'utf-8');
}
const ALLOWED_IMAGE_MIMES = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/heic', 'image/heif'];
const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (!ALLOWED_IMAGE_MIMES.includes(file.mimetype)) {
      securityLog('UPLOAD_REJECT_MIME', { ip: req.ip, mime: file.mimetype });
      return cb(new Error('Зөвхөн зураг хүлээн авна (jpg, png, webp, heic)'));
    }
    cb(null, true);
  }
});

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
    'You are transcribing an official Mongolian property certificate (гэрчилгээ).',
    'Read the ENTIRE text verbatim, character by character. Output plain text only.',
    '',
    'Key font confusion pairs to watch:',
    '  Ц vs У (Ц has bottom-right tail), Т vs П (Т=1 stroke, П=2), Ү vs У (Ү has dots)',
    '  Ш vs Т, Ю vs И, Ц vs Ч, 2 vs 9, х vs т',
    '  П vs Н — П has TWO vertical strokes with top bar. Н has crossbar in MIDDLE. "Пагваа" not "Нагваа".',
    '  э vs ө — э opens to the RIGHT. ө is closed circle with dots. "Тэлмэн" not "Төхөм".',
    '  л vs х — л has a diagonal stroke going down-right. х has two CROSSING strokes.',
  ].join('\n');

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 2048,
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
    '{"name":"","register":"","name2":"","register2":"","ownerCount":"","address":"","area":"","rooms":"","cert":"","purpose":""}',
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
    '  CRITICAL: Copy the register number EXACTLY as it appears in the text. Do NOT reorder or change any digits.',
    'address: full address including дүүрэг, хороо, байр, тоот',
    'area: digits + м.кв before "талбайтай". Example: "43 м.кв"',
    'rooms: number of rooms before "өрөө". нэг=1, хоёр=2, гурав=3, дөрөв=4. Example: "2"',
    'cert: alphanumeric code near "гэрчилгээ олгов". Starts with Ү-, Э-, Г-, Y-, V-. Example: Ү-2204001484, V-2204155889',
    'purpose: text before "зориулалттай" or "зориулалтаар"',
    '',
    'Return ONLY the JSON object, no explanation.',
  ].join('\n');
}
app.post('/analyze', analyzeLimiter, requireSiteKey, upload.single('image'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No image uploaded' });
  try {
    let imageBuffer = fs.readFileSync(req.file.path);
    let sharpImg = sharp(imageBuffer);
    const meta = await sharpImg.metadata();

    // Preprocessing: зөвхөн 5MB шалгах + grayscale + sharpen
    if (imageBuffer.length > 5 * 1024 * 1024) {
      sharpImg = sharpImg.resize({ width: 3000, height: 3000, fit: 'inside', withoutEnlargement: true });
    }
    imageBuffer = await sharpImg
      .grayscale()
      .normalise()
      .sharpen({ sigma: 1.5, m1: 0.5, m2: 3 })
      .jpeg({ quality: 90 })
      .toBuffer();

    // Compress хэрэв хэтэрсэн бол
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
        model: 'claude-sonnet-4-6',
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

    // Нэрний алдаа засах
    const fixName = (n) => {
      if (!n) return n;
      n = n.replace(/зуус$/i, 'цэцэг').replace(/Зуус$/i, 'цэцэг');
      n = n.replace(/^Нагваа/, 'Пагваа');
      n = n.replace(/^Нагва /, 'Пагваа ');
      // Тэлмэн нэр давтан буруу уншигдаж байна
      n = n.replace(/Төхөм$/i, 'Тэлмэн');
      n = n.replace(/Тохом$/i, 'Тэлмэн');
      return n;
    };
    extracted.name  = fixName(extracted.name);
    extracted.name2 = fixName(extracted.name2);

    res.json({ success: true, data: extracted, rawText });
  } catch (err) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch (_) {}
    res.status(500).json({ error: 'Failed: ' + err.message });
  }
});
app.post('/parse-text', analyzeLimiter, requireSiteKey, async (req, res) => {
  const { rawText } = req.body;
  if (!rawText) return res.status(400).json({ error: 'rawText missing' });
  try {
    const parsePrompt = buildParsePrompt(rawText);
    const parseResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 1024, messages: [{ role: 'user', content: parsePrompt }] })
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

app.post('/generate', generateLimiter, requireSiteKey, async (req, res) => {
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
      rooms:          data.rooms          ?? '',
      purpose:        data.purpose        ?? '',
      address:        data.address        ?? '',
      owner:          data.name           ?? '',
      phone:          data.phone          ?? '',
      register:       data.register       ?? '',
      createdAt:      new Date().toISOString(),
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="contract_${type}_${subtype}.docx"`);
    res.send(docBuffer);
  } catch (err) {
    // Docxtemplater multi-error дэлгэрэнгүй харуулах
    if (err.properties && err.properties.errors) {
      const details = err.properties.errors.map(e => e.properties?.explanation || e.message).join('; ');
      return res.status(500).json({ error: 'Template error: ' + details });
    }
    res.status(500).json({ error: 'Failed: ' + err.message });
  }
});

app.post('/preview', generateLimiter, requireSiteKey, (req, res) => {
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

// Admin: гэрээ засах
app.put('/admin/contracts/:id', adminLimiter, requireAdmin, (req, res) => {
  const contracts = readContracts();
  const idx = contracts.findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Олдсонгүй' });
  const allowed = ['contractNumber','agent','startDate','endDate','owner','register','phone','address','area','rooms','purpose','listingType'];
  allowed.forEach(k => { if (req.body[k] !== undefined) contracts[idx][k] = req.body[k]; });
  writeContracts(contracts);
  res.json({ ok: true });
});

// Admin: гэрээ татах (.docx)
app.get('/admin/contracts/:id/download', adminLimiter, requireAdmin, async (req, res) => {
  try {
    const c = readContracts().find(c => c.id === req.params.id);
    if (!c) return res.status(404).json({ error: 'Олдсонгүй' });
    const [type, subtype] = (c.listingType || 'sell / standard').split(' / ').map(s => s.trim());
    const { getTemplate } = require('./templates');
    const template = getTemplate(type, subtype);
    if (!template) return res.status(404).json({ error: 'Template олдсонгүй' });
    const { generateDocx } = require('./docxGenerator');
    const docBuffer = await generateDocx(template, c, type, subtype);
    const filename = `contract_${c.contractNumber || c.id}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(docBuffer);
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
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
    'Өрөөний тоо':               c.rooms          ?? '',
    'Зориулалт':                 c.purpose        ?? '',
    'Листингийн байршил':        c.address        ?? '',
    'ҮХХ эзэмшигчийн мэдээлэл': c.owner          ?? '',
    'Утасны дугаар':             c.phone          ?? '',
    'РД':                        c.register       ?? '',
  }));
  const workbook  = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 14 },
    { wch: 18 }, { wch: 14 }, { wch: 10 }, { wch: 22 }, { wch: 30 }, { wch: 28 }, { wch: 14 }, { wch: 12 },
  ];
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Гэрээнүүд');
  const buffer   = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const filename = `contracts_${new Date().toISOString().slice(0, 10)}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});
app.get('/next-contract-number', generateLimiter, requireSiteKey, (req, res) => {
  const { type, subtype } = req.query;

  const config = {
    'sell_exclusive': { prefix: 'ОХ', startAt: 35 },
    'sell_standard':  { prefix: 'ЭХ', startAt: 40 },
    'rent_exclusive': { prefix: 'ОТ', startAt: 1  },
    'rent_standard':  { prefix: 'ЭТ', startAt: 15 },
  };

  const key = `${type}_${subtype}`;
  const { prefix, startAt } = config[key] || { prefix: 'ГЭ', startAt: 1 };
  const year = new Date().getFullYear().toString().slice(-2);
  const pad = 3;
  const contracts = readContracts().filter(c => c.listingType === `${type} / ${subtype}`);

  // Хэрэглэгдсэн дугааруудыг цуглуулах
  const usedNums = new Set();
  let maxNum = startAt - 1;
  contracts.forEach(c => {
    if (c.contractNumber) {
      const m = c.contractNumber.match(/\/(\d+)$/);
      if (m) {
        const n = parseInt(m[1], 10);
        usedNums.add(n);
        if (n > maxNum) maxNum = n;
      }
    }
  });

  // Устгасан (gap) дугааруудаас хамгийн бага нэгийг хайх
  let next = null;
  for (let i = startAt; i <= maxNum; i++) {
    if (!usedNums.has(i)) { next = i; break; }
  }
  // Gap байхгүй бол max + 1
  if (next === null) next = maxNum + 1;

  res.json({ contractNumber: `${prefix}${year}/${String(next).padStart(pad, '0')}` });
});
app.listen(process.env.PORT || 3000, () => {
  console.log(`Server running on port ${process.env.PORT || 3000}`);
  console.log('[STARTUP] Site-key, env validation OK');

  // ── Background cleanup ──
  // Session cleanup
  setInterval(() => {
    const now = Date.now();
    for (const [token, session] of sessions.entries()) {
      if (now - session.createdAt > SESSION_TTL_MS) sessions.delete(token);
    }
  }, 10 * 60 * 1000);

  // Login attempts cleanup
  setInterval(() => {
    const now = Date.now();
    for (const [ip, a] of loginAttempts.entries()) {
      if (a.lockedUntil && now >= a.lockedUntil) loginAttempts.delete(ip);
    }
  }, 5 * 60 * 1000);

  // Uploads cleanup (1 цагийн дараа)
  function cleanupUploads() {
    const dir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(dir)) return;
    const cutoff = Date.now() - 60 * 60 * 1000;
    try {
      fs.readdirSync(dir).forEach(f => {
        try {
          const fp = path.join(dir, f);
          if (fs.statSync(fp).mtimeMs < cutoff) fs.unlinkSync(fp);
        } catch (_) {}
      });
    } catch (_) {}
  }
  setInterval(cleanupUploads, 30 * 60 * 1000);
  cleanupUploads();
});

// ── Multer error handler ──
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    return res.status(err.code === 'LIMIT_FILE_SIZE' ? 413 : 400).json({ error: err.message });
  }
  if (err?.message?.includes('Зөвхөн зураг')) return res.status(400).json({ error: err.message });
  next(err);
});
 
