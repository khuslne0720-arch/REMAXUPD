const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');

function formatMN(dateStr) {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    const year  = parseInt(parts[0]);
    const month = parseInt(parts[1]);
    const day   = parseInt(parts[2]);
    return `${year} оны ${month}-р сарын ${day}`;
  }
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  return `${d.getFullYear()} оны ${d.getMonth()+1}-р сарын ${d.getDate()}`;
}

// Үнийг форматлах: 150000000 → 150,000,000
function formatPrice(val) {
  if (!val && val !== 0) return '';
  const num = parseFloat(String(val).replace(/[,\s₮]/g, ''));
  if (isNaN(num)) return String(val);
  return num.toLocaleString('en-US');
}

async function generateDocx(template, data, type, subtype) {
  const templatePath = path.join(__dirname, 'templates', `${type}_${subtype}.docx`);
  if (!fs.existsSync(templatePath)) throw new Error(`Template файл олдсонгүй: ${type}_${subtype}.docx`);

  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
    nullGetter: () => '',
  });

  doc.render({
    contractNumber: data.contractNumber || '',
    duration:       data.duration       || '',
    startDate:      formatMN(data.startDate),
    endDate:        formatMN(data.endDate),
    name:           data.name           || '',
    register:       data.register       || '',
    name2:          data.name2          || '',
    register2:      data.register2      || '',
    address:        data.address        || '',
    area:           data.area           || '',
    cert:           data.cert           || '',
    phone:          data.phone          || '',
    email:          data.email          || '',
    price:          formatPrice(data.price),
    agent:              data.agent              || '',
    residentialAddress: data.residentialAddress || '',
    commissionRate:     data.commissionRate     || '',
    commissionAmount:   formatPrice((data.commissionAmount || '').replace(/[₮\s]/g, '').replace(/,/g, '')),
    purpose:            data.purpose            || '',
    rooms:              data.rooms              || '',
  });

  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generateDocx };
