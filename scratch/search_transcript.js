const fs = require('fs');
const path = require('path');

const logPath = 'C:\\Users\\Honor\\.gemini\\antigravity-ide\\brain\\38515e2c-74ad-4699-b0de-a80f373fd64e\\.system_generated\\logs\\transcript.jsonl';

if (!fs.existsSync(logPath)) {
  console.log('transcript.jsonl not found at:', logPath);
} else {
  const content = fs.readFileSync(logPath, 'utf8');
  const lines = content.split('\n');
  console.log(`Found ${lines.length} lines in transcript.`);

  lines.forEach((line, idx) => {
    if (line.toLowerCase().includes('март') || line.toLowerCase().includes('march')) {
      console.log(`Line ${idx + 1}:`, line.slice(0, 1000));
    }
  });
}
