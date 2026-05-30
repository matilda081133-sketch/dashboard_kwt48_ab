const fs = require('fs');
const path = require('path');

const logPath = 'C:\\Users\\Honor\\.gemini\\antigravity-ide\\brain\\38515e2c-74ad-4699-b0de-a80f373fd64e\\.system_generated\\logs\\transcript.jsonl';

if (!fs.existsSync(logPath)) {
  console.log('transcript.jsonl not found at:', logPath);
} else {
  const content = fs.readFileSync(logPath, 'utf8');
  const lines = content.split('\n');
  lines.forEach((line, idx) => {
    if (!line) return;
    try {
      const obj = JSON.parse(line);
      if (obj.source === 'USER_EXPLICIT' || obj.type === 'USER_INPUT') {
        console.log(`Step ${obj.step_index}:`, obj.content);
      }
    } catch(e) {}
  });
}
