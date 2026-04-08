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
app.post('/analyze', analyzeLimiter, upload.single('image'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No image uploaded' });
  try {
    let imageBuffer = fs.readFileSync(req.file.path);
    // Зургийг OCR-д зориулж сайжруулах: resize + grayscale + contrast + sharpen
    let sharpImg = sharp(imageBuffer);
    const meta = await sharpImg.metadata();
    if (meta.width > 3500 || meta.height > 3500) {
      sharpImg = sharpImg.resize({ width: 3500, height: 3500, fit: 'inside', withoutEnlargement: true });
    }
    imageBuffer = await sharpImg
      .grayscale()
      .normalise()
      .sharpen({ sigma: 1.5, m1: 0.5, m2: 3 })
      .jpeg({ quality: 95 })
      .toBuffer();
    const base64Image = imageBuffer.toString('base64');
    const mimeType = 'image/jpeg';
    const prompt = [
      'You are an expert OCR specialist for official Mongolian government property certificates (үл хөдлөх хөрөнгийн гэрчилгээ).',
      'These certificates use a SPECIAL DECORATIVE CALLIGRAPHIC FONT that is unique to official Mongolian documents.',
      'This font has specific visual characteristics that commonly cause OCR errors. Apply the corrections below.',
      '',
      'Return ONLY this JSON, no explanation, no markdown:',
      '{"name":"","register":"","name2":"","register2":"","ownerCount":"","address":"","area":"","cert":"","purpose":""}',
      '',
      '--- FONT DISAMBIGUATION RULES (apply to ALL text reading) ---',
      '',
      'This decorative font commonly causes these specific confusions. Always check:',
      '  "ц" vs "у" — CRITICAL: Ц has a small tail/descender at bottom-right, У does not. Register prefixes like ЦБ are often misread as УБ.',
      '  "т" vs "п" — CRITICAL: Т has a single vertical stroke, П has two vertical strokes. УТ is often misread as УП.',
      '  "я" vs "л" — CRITICAL: я curves to the right at top, л has a straight diagonal. "Баяр" is often misread as "Балр".',
      '  "лм" cluster — the letter л before м can be missed entirely. "Тэлмэн" may be misread as "Тэмэн".',
      '  "ц" vs "з" — цэцэг has TWO ц letters (not з)',
      '  "э" vs "о" — look at the opening direction of the curve',
      '  "н" vs "и" — н has a crossbar in the middle, и has it at the top',
      '  "г" vs "т" — г curves down-right, т has a horizontal top',
      '  "ү" vs "у" — ү has an umlaut (two dots), у does not',
      '  "ө" vs "о" — ө has an umlaut (two dots), о does not',
      '  "л" vs "д" — very similar in this font, check the base stroke',
      '  Tall decorative strokes on letters like "б","д","р" can be misread as other letters',
      '',
      'REGISTER PREFIX SPECIAL RULES:',
      '  The first 2 letters of a register are ALWAYS Cyrillic uppercase.',
      '  Most common real prefixes: УБ, ЦБ, УЛ, УО, ИЦ, АА, ОУ, БА, ДА, УТ, НУ, ЖН',
      '  NEVER write УП — "П" almost never appears as a register prefix. If you read УП, it is most likely УТ.',
      '  NEVER confuse Ц with У — look carefully at the bottom of the letter for the Ц tail.',
      '',
      'COMMON MONGOLIAN FIRST NAMES — if you read something that does not match a real Mongolian name,',
      'reconsider your reading using these frequent name components:',
      '  Endings: -цэцэг, -нүүр, -сүрэн, -баяр, -болд, -мөнх, -гэрэл, -наран, -өлзий, -сайхан,',
      '           -хүү, -баатар, -дорж, -жаргал, -мягмар, -энхэ, -зул, -уянга, -дулам, -номин',
      '  Prefixes: Алтан-, Номун-, Энх-, Мөнх-, Баян-, Түмэн-, Дэлгэр-, Эрдэнэ-, Ган-, Түвшин-',
      '',
      '--- FIELD RULES ---',
      '',
      'ownerCount: Find the text in format "/X иргэний өмч/" or "/хоёр иргэний өмч/".',
      '  Examples: "/нэг иргэний өмч/" → "1", "/хоёр иргэний өмч/" → "2", "/гурав иргэний өмч/" → "3"',
      '  This text is usually at the end of the owner section.',
      '',
      'OWNER NAMES - Two patterns exist:',
      '',
      'PATTERN A - "овогтой" (single surname word):',
      '  "...овгийн [SURNAME] овогтой [FIRSTNAME] [REGISTER]"',
      '  Take ONLY the word immediately before "овогтой" as surname.',
      '  Example: "Иижэн овгийн Дэмбэрэлсамбуу овогтой Алтанцэцэг УБ56050863"',
      '  → name="Дэмбэрэлсамбуу Алтанцэцэг", register="УБ56050863"',
      '  NOTE: "Алтанцэцэг" ends in -цэцэг (flower), NOT -нүүрс or -нүүр.',
      '  After reading a name, ask yourself: is this a real Mongolian name? If unsure, re-read the letters.',
      '',
      'PATTERN B - "овгийн" WITHOUT "овогтой":',
      '  "[CLAN] овгийн [SURNAME] [FIRSTNAME] [REGISTER]"',
      '  The CLAN word before "овгийн" is NOT part of the name. IGNORE it completely.',
      '  Take ONLY the TWO words immediately after "овгийн" as Surname+Firstname.',
      '  Example: "Бургууд овгийн Дашдэлэг Байгалмаа ПО99072721"',
      '  → CLAN="Бургууд" (ignore), name="Дашдэлэг Байгалмаа", register="ПО99072721"',
      '  Example: "Дайдал овгийн Болдбаатар Солонго УЛ087061526"',
      '  → CLAN="Дайдал" (ignore), name="Болдбаатар Солонго", register="УЛ087061526"',
      '  CRITICAL: NEVER include the clan/tribe word (word before "овгийн") in the name field.',
      '',
      'FOR MULTIPLE OWNERS (when ownerCount >= 2):',
      '  The certificate lists owners separated by comma or newline.',
      '  FIRST owner → name + register',
      '  SECOND owner → name2 + register2',
      '  Example: "Дайдал овгийн Болдбаатар Солонго УЛ087061526, Шаравд овгийн Доржсүрэн Батзаяа ИЦ88022819 /хоёр иргэний өмч/"',
      '  → name="Болдбаатар Солонго", register="УЛ087061526", name2="Доржсүрэн Батзаяа", register2="ИЦ88022819", ownerCount="2"',
      '',
      'register format: exactly 2 Cyrillic UPPERCASE letters + 8 digits. Examples: УБ56050863, УЛ08706152, ИЦ88022819',
      '  Common register prefixes: УБ, УЛ, УО, ИЦ, АА, ББ, ОУ, БА, ДА — always 2 letters.',
      'address: Full address with дүүрэг/хороо/байр/тоот.',
      'area: Number + м.кв before "талбайтай". Example: "51 м.кв", "43 м.кв", "45.71 м.кв"',
      'cert: Near "бүртгэж гэрчилгээ олгов". Starts with Ү-,Э-,Г-,Y-,V-. Example: V-2204001484, Y-2204099086',
      '',
      'purpose: Find the property purpose/type in the certificate.',
      '  TYPE A (үл хөдлөх): Look for text before "зориулалттай".',
      '  Examples: "Орон сууцны зориулалттай" → purpose="Орон сууцны зориулалттай"',
      '            "Оффисын зориулалттай" → purpose="Оффисын зориулалттай"',
      '  TYPE B (газар): Look for text before "зориулалтаар".',
      '  Examples: "Гэр, орон сууцны хашааны газар зориулалтаар" → purpose="Гэр, орон сууцны хашааны газар"',
      '  IMPORTANT: Do NOT confuse "хашааны газар" (land) with "орон сууцны" (apartment).',
      '',
      'Return ONLY the JSON object.'
    ].join('\n');
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-opus-4-5', max_tokens: 2048, messages: [{ role: 'user', content: [{ type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Image } }, { type: 'text', text: prompt }] }] })
    });
    const aiData = await response.json();
    fs.unlinkSync(req.file.path);
    if (aiData.error) return res.status(500).json({ error: aiData.error.message });
    const text = aiData.content?.[0]?.text || '{}';
    const extracted = JSON.parse(text.replace(/```json|```/g, '').trim());
    res.json({ success: true, data: extracted });
  } catch (err) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch (_) {}
    res.status(500).json({ error: 'Failed: ' + err.message });
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
