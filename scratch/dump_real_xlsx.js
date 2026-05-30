const fs = require('fs');
const path = require('path');

import('../../node_modules/xlsx/xlsx.mjs').then(async (XLSX) => {
  const filePath = path.join(__dirname, '..', '..', 'ULTIMATE_KILOVATT_DASHBOARD.xlsx');
  
  if (!fs.existsSync(filePath)) {
    console.error('File not found:', filePath);
    return;
  }

  const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
  const sheet = workbook.Sheets['Dashboard'];
  const ref = sheet['!ref'];
  console.log('ULTIMATE_KILOVATT_DASHBOARD.xlsx Dashboard Range:', ref);
  const range = XLSX.utils.decode_range(ref);

  for (let r = range.s.r; r <= range.e.r; r++) {
    const rowCells = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      rowCells.push(cell ? String(cell.v) : '');
    }
    if (rowCells.some(val => val !== '')) {
      console.log(`Row ${r + 1}:`, rowCells.map(val => val.padEnd(15)).join(' | '));
    }
  }
});
