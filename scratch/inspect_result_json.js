const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, '..', '..', 'result.json');
if (!fs.existsSync(filePath)) {
  console.log('result.json not found');
} else {
  const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  console.log('Keys in result.json:', Object.keys(data).slice(0, 10));
  if (Array.isArray(data)) {
    console.log('Length of data:', data.length);
    console.log('Sample item:', data[0]);
  } else {
    console.log('Type of data:', typeof data);
    console.log('Sample:', JSON.stringify(data).slice(0, 500));
  }
}
