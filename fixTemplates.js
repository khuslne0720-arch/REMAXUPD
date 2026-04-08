// fixTemplates.js - contractNumber, startDate, endDate, agent placeholder нэмэх
const PizZip = require('pizzip');
const fs = require('fs');
const path = require('path');

function fixTemplate(filePath) {
  const content = fs.readFileSync(filePath, 'binary');
  const zip = new PizZip(content);
  let xml = zip.files['word/document.xml'].asText();

  // 1. Гэрээний дугаар: №ОТ26/ эсвэл №ЭХ26/ гэх мэт → {{contractNumber}}
  xml = xml.replace(/№[А-ЯӨҮЁ]{2}26\/[^<]*/g, '{{contractNumber}}');
  xml = xml.replace(/№[А-ЯӨҮЁ]{2}26\//g, '{{contractNumber}}');

  // 2. Эхлэх огноо: "....-р сарын ....-ний/ны өдөр эхэлж" → {{startDate}} оруулах
  // Гэрээний хугацааны заалт дахь огнооны dots
  xml = xml.replace(
    /(хугацаатай байх бөгөөд 2026 оны\s*(?:<\/w:t>[\s\S]*?<w:t[^>]*>)?)[.\u2026]{3,}(-р сарын\s*(?:<\/w:t>[\s\S]*?<w:t[^>]*>)?)[.\u2026]{3,}(ний\/ны өдөр эхэлж)/g,
    '$1{{startDate}}$2__$3'
  );
  
  // 3. Дуусах огноо pattern
  xml = xml.replace(
    /(эхэлж,\s*(?:<\/w:t>[\s\S]*?<w:t[^>]*>)?)[.\u2026]{3,}(оны\s*(?:<\/w:t>[\s\S]*?<w:t[^>]*>)?)[.\u2026]{3,}(-р сарын\s*(?:<\/w:t>[\s\S]*?<w:t[^>]*>)?)[.\u2026]{3,}(ний\/ны өдөр дуусгавар)/g,
    '$1{{endDate}}$2__$3__$4'
  );

  // 4. Агент: standalone dots after "АГЕНТ:" section  
  // Already handled in previous scripts

  zip.file('word/document.xml', xml);
  const buffer = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(filePath, buffer);
}

const files = ['sell_standard','sell_exclusive','rent_standard','rent_exclusive'];
files.forEach(f => {
  const p = path.join('templates', f + '.docx');
  try {
    fixTemplate(p);
    console.log('✅ ' + f);
  } catch(e) {
    console.log('❌ ' + f + ': ' + e.message);
  }
});
console.log('Done!');
