// addPlaceholders.js - Нэг удаа ажиллуулна!
const PizZip = require('pizzip');
const fs = require('fs');
const path = require('path');

function processDocx(inputPath, outputPath) {
  const content = fs.readFileSync(inputPath, 'binary');
  const zip = new PizZip(content);
  let xml = zip.files['word/document.xml'].asText();

  // XML run-уудын хооронд текст хуваагддаг тул
  // эхлээд бүх runs-ийг нэгтгэсэн дараа replace хийнэ
  
  // Run-уудыг нэгтгэх: <w:t>text1</w:t>...<w:t>text2</w:t> -> <w:t>text1text2</w:t>
  // Гэхдээ энэ нь форматыг алдагдуулна тул өөр арга ашиглана:
  // XML-д ..... pattern-ийг шууд replace хийнэ
  
  // Хамгийн олон цэгтэй pattern-уудыг эхлээд орлуулна (урттаас богиш руу)
  const dotReplacements = [
    // Хаяг - address (хамгийн урт dots, иргэн болон хаягт хооронд)
    [/(\u0438\u0440\u0433\u044d\u043d\s*)\.{20,}(\s*\u0445\u0430\u044f\u0433\u0442)/g, '$1{{address}}$2'],
    // Нэр - name (оршин суух болон овогтой хооронд)  
    [/(\u043e\u0440\u0448\u0438\u043d\s*\u0441\u0443\u0443\u0445\s*)\.{10,}(\s*\u043e\u0432\u043e\u0433\u0442\u043e\u0439)/g, '$1{{name}}$2'],
    // Регистр - register
    [/(\u0420\u0435\u0433\u0438\u0441\u0442\u0440\u0438\u0439\u043d\s*\u0434\u0443\u0433\u0430\u0430\u0440:\s*)\.{5,}(\s*\))/g, '$1{{register}}$2'],
    // Агент
    [/(\u0410\u0433\u0435\u043d\u0442\s*\u0430\u0436\u0438\u043b\u0442\u0430\u0439\s*)\.{10,}(\s*\()/g, '$1{{agent}}$2'],
    // Signature: Овог нэр
    [/(\u041e\u0432\u043e\u0433\s*\u043d\u044d\u0440:\s*)\.{5,}/g, '$1{{name}} /                    /'],
    // Signature: Хаяг
    [/(\u0425\u0430\u044f\u0433:\s*)\.{5,}/g, '$1{{address}}'],
    // Signature: Утас
    [/(\u0423\u0442\u0430\u0441:\s*)\.{5,}/g, '$1{{phone}}'],
    // Signature: И-мэйл
    [/(\u0418-\u043c\u044d\u0439\u043b\s*\u0445\u0430\u044f\u0433:\s*)\.{5,}/g, '$1{{email}}'],
    // Agent signature line (standalone dots)
    [/^\.{15,}$/gm, '{{agent}} /                    /'],
  ];

  for (const [pattern, replacement] of dotReplacements) {
    xml = xml.replace(pattern, replacement);
  }

  zip.file('word/document.xml', xml);
  const buffer = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(outputPath, buffer);
}

const files = ['rent_exclusive', 'rent_standard', 'sell_exclusive', 'sell_standard'];

console.log('Placeholder нэмж байна...\n');
let success = 0;
for (const name of files) {
  const filePath = path.join('templates', `${name}.docx`);
  try {
    processDocx(filePath, filePath);
    console.log(`✅ ${name}.docx`);
    success++;
  } catch (e) {
    console.log(`❌ ${name}.docx - ${e.message}`);
  }
}
console.log(`\n${success}/${files.length} амжилттай!`);
