const fs = require('fs');
const path = require('path');

import('../../node_modules/xlsx/xlsx.mjs').then(async (XLSX) => {
  const rootDir = path.join(__dirname, '..', '..');
  const files = fs.readdirSync(rootDir);
  const xlsxFiles = files.filter(f => f.endsWith('.xlsx'));

  console.log('Scanning xlsx files in:', rootDir);

  xlsxFiles.forEach(file => {
    const filePath = path.join(rootDir, file);
    try {
      const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
      const marchSheets = workbook.SheetNames.filter(s => s.toLowerCase().indexOf('март') !== -1);
      if (marchSheets.length > 0) {
        console.log(`FOUND March sheets in "${file}":`, marchSheets);
      } else {
        console.log(`File "${file}" sheet names:`, workbook.SheetNames);
      }
    } catch (e) {
      console.error(`Error reading ${file}:`, e.message);
    }
  });
});
