const fs = require('fs');
const path = require('path');

const logPath = 'C:\\Users\\Honor\\.gemini\\antigravity-ide\\brain\\38515e2c-74ad-4699-b0de-a80f373fd64e\\.system_generated\\logs\\transcript.jsonl';

if (!fs.existsSync(logPath)) {
  console.log('transcript.jsonl not found');
} else {
  const content = fs.readFileSync(logPath, 'utf8');
  const lines = content.split('\n');
  lines.forEach((line, idx) => {
    if (!line) return;
    const lower = line.toLowerCase();
    if (lower.includes('%') || lower.includes('выполн') || lower.includes('процент') || lower.includes('цвет')) {
      try {
        const obj = JSON.parse(line);
        if (obj.source === 'USER_EXPLICIT' || obj.type === 'USER_INPUT') {
          console.log(`Step ${obj.step_index} USER:`, obj.content);
        } else if (obj.source === 'MODEL' && obj.type === 'PLANNER_RESPONSE') {
          console.log(`Step ${obj.step_index} MODEL:`, obj.content ? obj.content.slice(0, 300) : '');
        }
      } catch(e) {}
    }
  });
}
