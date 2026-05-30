const fs = require('fs');
const path = require('path');

import('../../node_modules/xlsx/xlsx.mjs').then(async (XLSX) => {
  const rootDir = path.join(__dirname, '..', '..');
  const filePath = path.join(rootDir, 'ULTIMATE_KILOVATT_DASHBOARD.xlsx');
  
  if (!fs.existsSync(filePath)) {
    console.error('File ULTIMATE_KILOVATT_DASHBOARD.xlsx not found');
    return;
  }

  const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
  console.log('Sheet names:', workbook.SheetNames);
  
  const sheet = workbook.Sheets['Dashboard'];
  if (!sheet) {
    console.error('No Dashboard sheet found');
    return;
  }

  const ref = sheet['!ref'];
  console.log('Range:', ref);
  const range = XLSX.utils.decode_range(ref);

  // Scan cells to look for date or March string
  for (let r = range.s.r; r <= Math.min(range.e.r, 50); r++) {
    for (let c = range.s.c; c <= Math.min(range.e.c, 20); c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v) {
        const valStr = String(cell.v);
        if (valStr.toLowerCase().includes('март') || valStr.includes('03.2026') || valStr.includes('03.03') || valStr.includes('01.03')) {
          console.log(`Found March reference at cell ${XLSX.utils.encode_col(c)}${r + 1}:`, valStr);
        }
      }
    }
  }
});
