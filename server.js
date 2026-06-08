const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const https = require('https');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');
const PLANS_FILE = process.env.PORT ? '/tmp/plans.json' : path.join(__dirname, 'plans.json');
const DB_FILE = process.env.PORT ? '/tmp/database.json' : path.join(__dirname, 'database.json');
const UPLOADS_DIR = process.env.PORT ? '/tmp/uploads' : path.join(__dirname, 'uploads');
// Uploads directory is created lazily on demand


// Database setup
let dbData = {
  users: {
    "admin": "kilowatt2026"
  },
  projects: {
    "proj_kilowatt": {
      "id": "proj_kilowatt",
      "name": "Проект Киловатт",
      "owner": "admin",
      "settings": {
        "roistatId": "298115",
        "roistatKey": "9d791b751c1b7c51d847914ec3c2b47e",
        "metrikaCounterId": "",
        "metrikaToken": "",
        "aiApiKey": "",
        "aiEndpoint": "https://api.openai.com/v1",
        "aiModel": "gpt-4o-mini",
        "aiProvider": "openai"
      },
      "dashboards": [
        {
          "id": "dash_rnp_kwt",
          "name": "Маркетинговый RNP (Киловатт)",
          "type": "rnp",
          "created_at": "2026-06-01T00:00:00.000Z"
        },
        {
          "id": "dash_landing_kwt",
          "name": "Эффективность лендингов (Киловатт)",
          "type": "landing",
          "created_at": "2026-06-01T00:00:00.000Z"
        }
      ]
    }
  },
  dashboardsData: {
    "dash_rnp_kwt": {
      "plans": {},
      "changes": [],
      "abTests": []
    },
    "dash_landing_kwt": {
      "plans": {},
      "changes": [],
      "abTests": []
    }
  }
};

function loadDB() {
  if (fs.existsSync(DB_FILE)) {
    try {
      dbData = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
    } catch (e) {
      console.error('Error loading database.json:', e);
    }
  } else {
    saveDB();
  }
}

function saveDB() {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(dbData, null, 2), 'utf8');
  } catch (e) {
    console.error('Error saving database.json:', e);
  }
}

loadDB();

function getSessionUser(req) {
  const cookies = req.headers.cookie || '';
  const match = cookies.match(/session_user=([^;]+)/);
  return match ? decodeURIComponent(match[1]) : null;
}

// Dynamically check parent directory or current directory for the Excel template
let excelPath = path.join(__dirname, 'rnp_template.xlsx');
if (!fs.existsSync(excelPath)) {
  const localPath = path.join(__dirname, 'rnp_template.xlsx');
  if (fs.existsSync(localPath)) {
    excelPath = localPath;
  }
}

let XLSX;

// Dynamically load XLSX module: try require first (CommonJS standard), then fallback to dynamic import
try {
  XLSX = require('xlsx');
  console.log('✅ XLSX module successfully loaded via require.');
} catch (e) {
  console.log('⚠️ require("xlsx") failed. Attempting dynamic import...');
  const tryImport = async () => {
    // Try current directory first, then parent directory
    const paths = ['./node_modules/xlsx/xlsx.mjs', '../node_modules/xlsx/xlsx.mjs'];
    for (const p of paths) {
      try {
        const resolvedPath = path.resolve(__dirname, p);
        if (fs.existsSync(resolvedPath)) {
          XLSX = await import(p);
          console.log(`✅ XLSX module dynamically loaded from: ${p}`);
          return;
        }
      } catch (err) {
        console.warn(`⚠️ Failed to import from ${p}:`, err.message);
      }
    }
    console.error('❌ All XLSX import paths failed.');
  };
  tryImport();
}


const KILOWATT_CHANNELS = ['Контекстная реклама', 'Сайт, органика', 'Телеграм', 'Вконтакте', 'МАХ', 'Выставки', 'Авито', 'Другое'];
const DEFAULT_CHANNELS = ['Контекстная реклама', 'SEO', 'SMM', 'Другое'];

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

// Row index mapping for Excel (1-based sheet rows)
const GROUPS_ROWS = {
  "Сайт, органика": {
    metrics: { cost: 24, visits: 25, leads: 26, qual: 29, kp: 32, sales: 34, rev: 35 }
  },
  "Контекстная реклама": {
    metrics: { cost: 41, visits: 42, leads: 43, qual: 46, kp: 49, sales: 51, rev: 52 }
  },
  "Телеграм": {
    metrics: { cost: 58, visits: 59, leads: 64, qual: 67, kp: 70, sales: 72, rev: 73 }
  },
  "Вконтакте": {
    metrics: { cost: 79, visits: 80, leads: 85, qual: 88, kp: 91, sales: 93, rev: 94 }
  },
  "МАХ": {
    metrics: { cost: 100, visits: 101, leads: 106, qual: 109, sales: 112, rev: 113 }
  },
  "Статьи в СМИ": {
    metrics: { cost: 117, visits: 118, leads: 119, qual: 122, sales: 125, rev: 126 }
  },
  "Специализированные выставки": {
    metrics: { cost: 130, visits: null, leads: 131, qual: 133, kp: 136, sales: 138, rev: 139 }
  }
};

// Classification logic (corresponds to dashboard_rnp.gs)
// Classification logic supporting both latin and cyrillic Roistat markers
function getGroup(title) {
  var t = (title || "").toLowerCase().trim();
  t = t.replace(/\u00A0/g, " ");

  // Explicitly exclude offline/direct non-marketing channels from RNP
  if (t.indexOf("от руководителя") !== -1 || t.indexOf("существующий клиент") !== -1) {
    return null;
  }

  // 1. МЕССЕНДЖЕРЫ & TELEGRAM -> Телеграм в РНП
  if (
    t.indexOf("tgapi") !== -1 ||
    t.indexOf("whatsapp") !== -1 ||
    t.indexOf("telegram") !== -1 ||
    t.indexOf("tg ") !== -1 ||
    t.indexOf("tg/") !== -1 ||
    t.indexOf(" tg") !== -1 ||
    t.startsWith("tg") ||
    t === "телеграм" ||
    t === "тг" ||
    ((t.indexOf("max") !== -1 || t.indexOf("макс") !== -1) && /7\d{9}/.test(t)) ||
    /79\d{9}/.test(t)
  ) {
    return "Телеграм";
  }

  // 2. Вконтакте & Соцсети
  if (
    t.indexOf("vk") !== -1 ||
    /(?:^|\s|_|-)вк(?:$|\s|_|-)/.test(t) ||
    t.indexOf("vkontakte") !== -1 ||
    t.indexOf("соц. сетей") !== -1 ||
    t.indexOf("соцсети") !== -1
  ) {
    return "Вконтакте";
  }

  // 3. ЯНДЕКС КАРТЫ -> Сайт, органика в РНП
  if (
    t.indexOf("ya_maps") !== -1 ||
    t.indexOf("yandex_maps") !== -1 ||
    t.indexOf("ya_map") !== -1 ||
    t.indexOf("yandex_map") !== -1 ||
    t.indexOf("yabs") !== -1 ||
    t.indexOf("карт") !== -1 ||
    t.indexOf("directory") !== -1 ||
    t.indexOf("business") !== -1 ||
    t.indexOf("бизнес") !== -1 ||
    t.indexOf("navigator") !== -1 ||
    t.indexOf("navi") !== -1
  ) {
    return "Сайт, органика";
  }

  // 4. ДИРЕКТ -> Контекстная реклама в РНП
  if (
    t.indexOf("direct") !== -1 ||
    t.indexOf("директ") !== -1 ||
    t.indexOf("звонок из директа") !== -1
  ) {
    return "Контекстная реклама";
  }

  // 5. КАНАЛ MAX (без номера телефона) -> МАХ в РНП
  if (
    t.indexOf("max") !== -1 ||
    t.indexOf("макс") !== -1 ||
    t.indexOf("мах") !== -1
  ) {
    return "МАХ";
  }

  // 6. SEO & General Calls -> Сайт, органика в РНП
  if (
    t.indexOf("seo") !== -1 ||
    t.indexOf("сео") !== -1 ||
    t.indexOf("веб-сайт") !== -1 ||
    t === "веб сайт" ||
    t.indexOf("звонок с сайта") !== -1 ||
    t === "yandex" ||
    t === "ya" ||
    t === "yandex.ru" ||
    t === "ya.ru" ||
    t === "google" ||
    t === "google.ru" ||
    t === "google.com" ||
    t === "mail.ru"
  ) {
    return "Сайт, органика";
  }

  // 7. ВЫСТАВКА -> Специализированные выставки в РНП
  if (t.indexOf("выставк") !== -1) {
    return "Специализированные выставки";
  }

  // 8. СТАТЬИ В СМИ -> Статьи в СМИ в РНП
  if (
    t.indexOf("сми") !== -1 ||
    t.indexOf("статьи") !== -1 ||
    t.indexOf("статья") !== -1
  ) {
    return "Статьи в СМИ";
  }

  return null;
}

// Convert Excel date serial to YYYY-MM-DD
function excelDateToDateString(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const pad = (n) => String(n).padStart(2, '0');
  return `${date_info.getFullYear()}-${pad(date_info.getMonth() + 1)}-${pad(date_info.getDate())}`;
}

// Read saved plan overrides
function readPlans() {
  if (fs.existsSync(PLANS_FILE)) {
    try {
      return JSON.parse(fs.readFileSync(PLANS_FILE, 'utf8'));
    } catch (e) {
      console.error('Error reading plans:', e);
      return {};
    }
  }
  return {};
}

// Save plans overrides
function writePlans(plans) {
  try {
    fs.writeFileSync(PLANS_FILE, JSON.stringify(plans, null, 2), 'utf8');
  } catch (e) {
    console.error('Error writing plans:', e);
  }
}

// Parse excel file to extract daily plan and fact values
function getExcelData(fromStr, toStr, customExcelPath, allowFallback = true, baseChannels = null) {
  if (!XLSX) {
    console.warn('⚡ XLSX module is not loaded yet. Returning empty array.');
    return [];
  }
  
  const targetPath = customExcelPath || (allowFallback ? excelPath : null);
  const hasFile = targetPath && fs.existsSync(targetPath);
  
  if (!hasFile) {
    console.warn('Spreadsheet not found at:', targetPath, '- generating empty skeleton.');
  }

  const dates = {}; // dateStr -> { sheetName, colIdx }
  let workbook = null;

  if (hasFile) {
    const fileBuffer = fs.readFileSync(targetPath);
    workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;
      const ref = sheet['!ref'];
      if (!ref) continue;
      const range = XLSX.utils.decode_range(ref);

      // Row 10 (index 9) contains dates
      for (let c = 10; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 9, c })];
        if (cell && typeof cell.v === 'number' && cell.v > 40000) {
          const dateStr = excelDateToDateString(cell.v);
          dates[dateStr] = { sheetName, colIdx: c };
        }
      }
    }
  }

  const results = [];
  let curr = new Date(fromStr);
  const end = new Date(toStr);

  while (curr <= end) {
    const pad = (n) => String(n).padStart(2, '0');
    const dateStr = `${curr.getFullYear()}-${pad(curr.getMonth() + 1)}-${pad(curr.getDate())}`;
    const dateInfo = dates[dateStr];

    const dayItem = {
      date: dateStr,
      channels: {}
    };

    const chList = baseChannels || Object.keys(GROUPS_ROWS);
    for (const channelName of chList) {
      dayItem.channels[channelName] = {
        plan: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 },
        fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }
      };
    }

    if (dateInfo) {
      const sheet = workbook.Sheets[dateInfo.sheetName];
      const c = dateInfo.colIdx;

      const getCellVal = (r, col) => {
        if (r === null) return 0;
        const cell = sheet[XLSX.utils.encode_cell({ r: r - 1, c: col })];
        if (!cell) return 0;
        if (cell.t === 'n') return cell.v || 0;
        if (cell.t === 'f') return typeof cell.v === 'number' ? cell.v : 0;
        const val = parseFloat(String(cell.v).replace(/\s/g, '').replace(',', '.'));
        return isNaN(val) ? 0 : val;
      };

      for (const [channelName, mapping] of Object.entries(GROUPS_ROWS)) {
        const ch = dayItem.channels[channelName];
        if (!ch) continue;
        ch.plan.cost = getCellVal(mapping.metrics.cost, c);
        ch.fact.cost = getCellVal(mapping.metrics.cost, c + 1);

        ch.plan.visits = getCellVal(mapping.metrics.visits, c);
        ch.fact.visits = getCellVal(mapping.metrics.visits, c + 1);

        ch.plan.leads = getCellVal(mapping.metrics.leads, c);
        ch.fact.leads = getCellVal(mapping.metrics.leads, c + 1);

        ch.plan.qual = getCellVal(mapping.metrics.qual, c);
        ch.fact.qual = getCellVal(mapping.metrics.qual, c + 1);

        ch.plan.kp = getCellVal(mapping.metrics.kp, c);
        ch.fact.kp = getCellVal(mapping.metrics.kp, c + 1);

        ch.plan.sales = getCellVal(mapping.metrics.sales, c);
        ch.fact.sales = getCellVal(mapping.metrics.sales, c + 1);

        ch.plan.rev = getCellVal(mapping.metrics.rev, c);
        ch.fact.rev = getCellVal(mapping.metrics.rev, c + 1);
      }
    }

    results.push(dayItem);
    curr.setDate(curr.getDate() + 1);
  }

  return results;
}

// Fetch live metrics from Roistat
function fetchRoistat(fromStr, toStr) {
  return new Promise((resolve, reject) => {
    const PID = '298115';
    const KEY = '9d791b751c1b7c51d847914ec3c2b47e';

    const fromIso = `${fromStr}T00:00:00+03:00`;
    const toIso = `${toStr}T23:59:59+03:00`;

    const payload = {
      "dimensions": ["date", "marker_level_1", "marker_level_2", "marker_level_3", "marker_level_4"],
      "metrics": [
        { "metric": "visits", "attribution": "default" },
        { "metric": "leads", "attribution": "default" },
        { "metric": "leadCount", "attribution": "default" },
        { "metric": "sales", "attribution": "default" },
        { "metric": "paidLeadCount", "attribution": "default" },
        { "metric": "revenue", "attribution": "default" },
        { "metric": "paidLeadsPrice", "attribution": "default" },
        { "metric": "marketing_cost", "attribution": "default" },
        { "metric": "visitsCost", "attribution": "default" },
        { "metric": "custom_2", "attribution": "default" }, // Qual Leads
        { "metric": "custom_5", "attribution": "default" }  // KP Sent
      ],
      "period": { "from": fromIso, "to": toIso }
    };

    const dataStr = JSON.stringify(payload);

    const options = {
      hostname: 'cloud.roistat.com',
      port: 443,
      path: `/api/v1/project/analytics/data?project=${PID}`,
      method: 'POST',
      headers: {
        'Api-key': KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(dataStr)
      },
      rejectUnauthorized: false // Bypass SSL issues locally
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          if (parsed.status === 'error') {
            reject(new Error(parsed.description || parsed.error_message || 'Roistat API returned an error status'));
          } else {
            resolve((parsed.data && parsed.data[0] && parsed.data[0].items) ? parsed.data[0].items : []);
          }
        } catch (e) {
          reject(new Error(`JSON Parse Error: ${e.message}`));
        }
      });
    });

    req.on('error', (e) => reject(e));
    req.write(dataStr);
    req.end();
  });
}

// Compute standard formulas
function computeFormulas(metrics) {
  const { cost, visits, leads, qual, kp, sales, rev } = metrics;
  return {
    cpl: leads > 0 ? cost / leads : 0,
    crLeads: visits > 0 ? leads / visits : 0,
    crQual: leads > 0 ? qual / leads : 0,
    cpql: qual > 0 ? cost / qual : 0,
    crKp: (kp !== null && qual > 0) ? kp / qual : 0,
    crSale: (kp > 0) ? sales / kp : (leads > 0 ? sales / leads : 0),
    cps: sales > 0 ? cost / sales : 0,
    romi: cost > 0 ? (rev - cost) / cost : 0,
    drr: rev > 0 ? cost / rev : 0
  };
}

// Generate fallback city distribution
function getMockCityData(leadsCount, salesCount) {
  const cities = ['Москва', 'Санкт-Петербург', 'Краснодар', 'Екатеринбург', 'Нижний Новгород', 'Ростов-на-Дону', 'Новосибирск', 'Казань', 'Самара', 'Воронеж'];
  const data = [];
  let remainingLeads = leadsCount;
  let remainingSales = salesCount;

  for (let i = 0; i < cities.length; i++) {
    const isLast = (i === cities.length - 1);
    const leadsShare = isLast ? remainingLeads : Math.round(remainingLeads * (0.35 - i * 0.03) * (0.8 + Math.random() * 0.4));
    const salesShare = isLast ? remainingSales : Math.round(remainingSales * (0.35 - i * 0.03) * (0.8 + Math.random() * 0.4));

    const leads = Math.max(0, Math.min(remainingLeads, leadsShare));
    const sales = Math.max(0, Math.min(remainingSales, salesShare));

    remainingLeads -= leads;
    remainingSales -= sales;

    if (leads > 0) {
      data.push({
        city: cities[i],
        leads,
        kp: Math.round(leads * 0.4),
        sales
      });
    }
  }
  return data.sort((a, b) => b.leads - a.leads);
}

// Scan Excel sheets, plans.json keys, and custom_months.json to get all active months
function getAvailableMonths() {
  const months = new Set();
  
  // 1. Get from Excel sheet names
  if (XLSX && fs.existsSync(excelPath)) {
    try {
      const fileBuffer = fs.readFileSync(excelPath);
      const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
      const monthIndexMap = {
        'январь': '01', 'февраль': '02', 'март': '03', 'апрель': '04', 'май': '05', 'мая': '05',
        'июнь': '06', 'июль': '07', 'август': '08', 'сентябрь': '09', 'октябрь': '10', 'ноябрь': '11', 'декабрь': '12'
      };
      
      workbook.SheetNames.forEach(sheetName => {
        const lower = sheetName.toLowerCase().trim();
        for (const [name, index] of Object.entries(monthIndexMap)) {
          if (lower.indexOf(name) !== -1) {
            months.add(`2026-${index}`); // Default to 2026 as per project structure
          }
        }
      });
    } catch (e) {
      console.error('Error scanning Excel sheet names:', e);
    }
  }
  
  // 2. Get from plans.json keys (YYYY-MM or YYYY-MM-DD)
  const plans = readPlans();
  Object.keys(plans).forEach(key => {
    if (key.length === 7 && key.match(/^\d{4}-\d{2}$/)) {
      months.add(key);
    } else if (key.length === 10 && key.match(/^\d{4}-\d{2}-\d{2}$/)) {
      months.add(key.slice(0, 7));
    }
  });
  
  // 3. Get from custom_months.json
  const customMonthsFile = path.join(__dirname, 'custom_months.json');
  if (fs.existsSync(customMonthsFile)) {
    try {
      const custom = JSON.parse(fs.readFileSync(customMonthsFile, 'utf8'));
      if (Array.isArray(custom)) {
        custom.forEach(m => months.add(m));
      }
    } catch (e) {}
  }
  
  return Array.from(months).sort();
}

// Read body helper
function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        resolve(JSON.parse(body));
      } catch (err) {
        resolve({});
      }
    });
    req.on('error', err => reject(err));
  });
}

// Handle API endpoints
// Helper to parse query parameters or URL params from paths
const getUrlParams = (pathPattern, requestPath) => {
  const patternParts = pathPattern.split('/');
  const pathParts = requestPath.split('/');
  if (patternParts.length !== pathParts.length) return null;
  const params = {};
  for (let i = 0; i < patternParts.length; i++) {
    if (patternParts[i].startsWith(':')) {
      params[patternParts[i].substring(1)] = pathParts[i];
    } else if (patternParts[i] !== pathParts[i]) {
      return null;
    }
  }
  return params;
};

function fetchProjectMetrics(projectId, apiKey) {
  return new Promise((resolve) => {
    const options = {
      hostname: 'cloud.roistat.com',
      port: 443,
      path: `/api/v1/project/analytics/metrics?project=${projectId}`,
      method: 'GET',
      headers: {
        'Api-key': apiKey
      },
      rejectUnauthorized: false
    };
    const req = https.request(options, (res) => {
      let body = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          if (parsed.status === 'success' && Array.isArray(parsed.metrics)) {
            resolve(parsed.metrics.map(m => ({ name: m.name, title: m.title })));
          } else {
            resolve([]);
          }
        } catch (e) {
          resolve([]);
        }
      });
    });
    req.on('error', () => resolve([]));
    req.end();
  });
}

function getGroupForProject(titleRaw, sourceGroupTitle, isDemo) {
  if (isDemo) {
    return getGroup(titleRaw);
  }
  if (sourceGroupTitle && sourceGroupTitle !== "Остальное" && sourceGroupTitle !== "Другое" && sourceGroupTitle !== "Неизвестный канал") {
    const g = sourceGroupTitle.trim();
    if (g.toLowerCase() === "seo") return "Сайт, органика";
    if (g.toLowerCase() === "контекст" || g.toLowerCase() === "контекстная реклама" || g.toLowerCase().includes("директ")) return "Контекстная реклама";
    return g;
  }
  return getGroup(titleRaw) || "Другое";
}

async function handleApi(req, res, pathname, query) {
  loadDB(); // Ensure memory is synced with disk on every request in multi-process environments
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  // POST /api/auth/register
  if (pathname === '/api/auth/register' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { username, password, email } = body;
      if (!username || !password) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Логин и пароль обязательны' }));
        return;
      }
      if (dbData.users[username]) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Пользователь уже существует' }));
        return;
      }
      dbData.users[username] = password;
      
      // Auto-create a default project for the new user
      const projId = 'proj_' + Math.floor(Math.random() * 1000000);
      dbData.projects[projId] = {
        id: projId,
        name: 'Мой первый проект',
        owner: username,
        settings: {
          roistatId: '',
          roistatKey: '',
          metrikaCounterId: '',
          metrikaToken: '',
          aiApiKey: '',
          aiEndpoint: '',
          aiModel: '',
          aiProvider: ''
        },
        dashboards: []
      };
      
      saveDB();
      res.setHeader('Set-Cookie', `session_user=${encodeURIComponent(username)}; Path=/; HttpOnly`);
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', username }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // POST /api/auth/login
  if (pathname === '/api/auth/login' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { username, password } = body;
      if (dbData.users[username] && dbData.users[username] === password) {
        res.setHeader('Set-Cookie', `session_user=${encodeURIComponent(username)}; Path=/; HttpOnly`);
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', username }));
      } else {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Неверное имя пользователя или пароль' }));
      }
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // POST /api/auth/logout
  if (pathname === '/api/auth/logout' && req.method === 'POST') {
    res.setHeader('Set-Cookie', `session_user=; Path=/; HttpOnly; Max-Age=0`);
    res.statusCode = 200;
    res.end(JSON.stringify({ status: 'success' }));
    return;
  }

  // GET /api/auth/me
  if (pathname === '/api/auth/me' && req.method === 'GET') {
    const user = getSessionUser(req);
    res.statusCode = 200;
    res.end(JSON.stringify({ authenticated: !!user, username: user }));
    return;
  }

  // Auth Guard
  const currentUser = getSessionUser(req);
  if (!currentUser) {
    res.statusCode = 401;
    res.end(JSON.stringify({ status: 'error', message: 'Unauthorized' }));
    return;
  }

  // Ensure demo project exists for the current user to prevent 404 in preview mode
  if (!dbData.projects['demo']) {
    dbData.projects['demo'] = {
      id: 'demo',
      name: 'Демо Проект',
      owner: currentUser,
      settings: {
        roistatId: '298115',
        roistatKey: '9d791b751c1b7c51d847914ec3c2b47e',
        metrikaCounterId: '',
        metrikaToken: '',
        aiApiKey: '',
        aiEndpoint: 'https://api.openai.com/v1',
        aiModel: 'gpt-4o-mini',
        aiProvider: 'openai'
      },
      connections: [],
      dashboards: [
        { id: 'demo', name: 'RNP (Demo)', type: 'rnp', created_at: new Date().toISOString() },
        { id: 'demo_landing', name: 'Эффективность лендингов (Demo)', type: 'landing', created_at: new Date().toISOString() }
      ]
    };
    dbData.dashboardsData['demo'] = { plans: {}, changes: [], abTests: [] };
    dbData.dashboardsData['demo_landing'] = { plans: {}, changes: [], abTests: [] };
    saveDB();
  } else if (dbData.projects['demo'].owner !== currentUser) {
    dbData.projects['demo'].owner = currentUser;
    saveDB();
  }

  // GET /api/projects
  if (pathname === '/api/projects' && req.method === 'GET') {
    const userProjects = Object.values(dbData.projects).filter(p => p.owner === currentUser || p.id === 'proj_kilowatt');
    res.statusCode = 200;
    res.end(JSON.stringify({ status: 'success', projects: userProjects }));
    return;
  }

  // POST /api/projects
  if (pathname === '/api/projects' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { name } = body;
      if (!name) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Название проекта обязательно' }));
        return;
      }
      const projId = 'proj_' + Math.floor(Math.random() * 1000000);
      dbData.projects[projId] = {
        id: projId,
        name,
        owner: currentUser,
        settings: {
          roistatId: '',
          roistatKey: '',
          metrikaCounterId: '',
          metrikaToken: '',
          aiApiKey: '',
          aiEndpoint: '',
          aiModel: '',
          aiProvider: ''
        },
        connections: [],
        dashboards: []
      };
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', project: dbData.projects[projId] }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // PUT /api/projects/:id
  const projectPutParams = getUrlParams('/api/projects/:id', pathname);
  if (projectPutParams && req.method === 'PUT') {
    try {
      const body = await parseBody(req);
      const projId = projectPutParams.id;
      const proj = dbData.projects[projId];
      const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
      if (!proj || (proj.owner !== currentUser && !isSystemProj)) {
        res.statusCode = 404;
        res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
        return;
      }
      proj.name = body.name || proj.name;
      proj.settings = { ...proj.settings, ...body.settings };
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', project: proj }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // DELETE /api/projects/:id
  const projectDelParams = getUrlParams('/api/projects/:id', pathname);
  if (projectDelParams && req.method === 'DELETE') {
    const projId = projectDelParams.id;
    const proj = dbData.projects[projId];
    const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
    if (proj && (proj.owner === currentUser || isSystemProj)) {
      delete dbData.projects[projId];
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success' }));
    } else {
      res.statusCode = 404;
      res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
    }
    return;
  }

  // POST /api/projects/:id/connections
  const connPostParams = getUrlParams('/api/projects/:id/connections', pathname);
  if (connPostParams && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const projId = connPostParams.id;
      const proj = dbData.projects[projId];
      const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
      if (!proj || (proj.owner !== currentUser && !isSystemProj)) {
        res.statusCode = 404;
        res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
        return;
      }
      
      const newConn = {
        id: 'conn_' + Math.random().toString(36).substr(2, 9),
        type: body.type || 'roistat',
        name: body.name || 'Новое подключение',
        params: body.params || {}
      };
      
      if (!proj.connections) proj.connections = [];
      proj.connections.push(newConn);
      saveDB();
      
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', connection: newConn }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // PUT /api/projects/:id/dashboards/:dashId/settings
  const dashSettingsParams = getUrlParams('/api/projects/:id/dashboards/:dashId/settings', pathname);
  if (dashSettingsParams && req.method === 'PUT') {
    try {
      const body = await parseBody(req);
      const { id: projId, dashId } = dashSettingsParams;
      const proj = dbData.projects[projId];
      const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
      if (!proj || (proj.owner !== currentUser && !isSystemProj)) {
        res.statusCode = 404;
        res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
        return;
      }
      const dashboard = proj.dashboards.find(d => d.id === dashId);
      if (!dashboard) {
        res.statusCode = 404;
        res.end(JSON.stringify({ status: 'error', message: 'Дашборд не найден' }));
        return;
      }
      
      if (body.connectionId !== undefined) dashboard.connectionId = body.connectionId;
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', dashboard }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // GET /api/projects/:id/dashboards
  const dashListParams = getUrlParams('/api/projects/:id/dashboards', pathname);
  if (dashListParams && req.method === 'GET') {
    const projId = dashListParams.id;
    const proj = dbData.projects[projId];
    const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
    if (proj && (proj.owner === currentUser || isSystemProj)) {
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', dashboards: proj.dashboards }));
    } else {
      res.statusCode = 404;
      res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
    }
    return;
  }

  // POST /api/projects/:id/dashboards
  if (dashListParams && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const projId = dashListParams.id;
      const proj = dbData.projects[projId];
      const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
      if (!proj || (proj.owner !== currentUser && !isSystemProj)) {
        res.statusCode = 404;
        res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
        return;
      }
      const { name, type } = body;
      if (!name || !type) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Имя и тип обязательны' }));
        return;
      }
      const dashId = 'dash_' + Math.floor(Math.random() * 1000000);
      const newDash = { id: dashId, name, type, created_at: new Date().toISOString() };
      proj.dashboards.push(newDash);
      dbData.dashboardsData[dashId] = {
        plans: {},
        changes: [],
        abTests: []
      };
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', dashboard: newDash }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // DELETE /api/projects/:id/dashboards/:dashId
  const dashDelParams = getUrlParams('/api/projects/:id/dashboards/:dashId', pathname);
  if (dashDelParams && req.method === 'DELETE') {
    const { id: projId, dashId } = dashDelParams;
    const proj = dbData.projects[projId];
    const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
    if (proj && (proj.owner === currentUser || isSystemProj)) {
      proj.dashboards = proj.dashboards.filter(d => d.id !== dashId);
      delete dbData.dashboardsData[dashId];
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success' }));
    } else {
      res.statusCode = 404;
      res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
    }
    return;
  }

  // POST /api/projects/:id/dashboards/:dashId/upload
  const dashUploadParams = getUrlParams('/api/projects/:id/dashboards/:dashId/upload', pathname);
  if (dashUploadParams && req.method === 'POST') {
    const { id: projId, dashId } = dashUploadParams;
    const proj = dbData.projects[projId];
    const isSystemProj = projId === 'proj_kilowatt' || projId === 'demo';
    if (proj && (proj.owner === currentUser || isSystemProj)) {
      const filePath = path.join(UPLOADS_DIR, `${projId}_${dashId}.xlsx`);
      try {
        if (!fs.existsSync(UPLOADS_DIR)) {
          fs.mkdirSync(UPLOADS_DIR, { recursive: true });
        }
      } catch (e) {
        console.error('Lazy directory creation failed:', e);
      }
      const fileStream = fs.createWriteStream(filePath);
      req.pipe(fileStream);
      fileStream.on('finish', () => {
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', message: 'Файл успешно загружен' }));
      });
      fileStream.on('error', (err) => {
        res.statusCode = 500;
        res.end(JSON.stringify({ status: 'error', message: err.message }));
      });
    } else {
      res.statusCode = 404;
      res.end(JSON.stringify({ status: 'error', message: 'Проект не найден' }));
    }
    return;
  }

  // POST /api/ai-programmer/chat
  if (pathname === '/api/ai-programmer/chat' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { message, projectId } = body;
      
      const proj = dbData.projects[projectId];
      const aiApiKey = proj?.settings?.aiApiKey || '';
      const aiEndpoint = proj?.settings?.aiEndpoint || 'https://api.openai.com/v1';
      const aiModel = proj?.settings?.aiModel || 'gpt-4o-mini';
      const aiProvider = proj?.settings?.aiProvider || 'openai';

      let promptText = "Ты — ИИ-программист, специализирующийся на интеграции Roistat и Яндекс.Метрики. Помогай отлаживать запросы и решать технические проблемы. Отвечай кратко, профессионально на русском языке.";

      if (aiApiKey) {
        const axios = require('axios');
        let reply = '';
        if (aiProvider === 'openai') {
          const response = await axios.post(`${aiEndpoint}/chat/completions`, {
            model: aiModel,
            messages: [
              { role: "system", content: promptText },
              { role: "user", content: message }
            ]
          }, {
            headers: { 'Authorization': `Bearer ${aiApiKey}`, 'Content-Type': 'application/json' }
          });
          reply = response.data.choices[0].message.content;
        } else if (aiProvider === 'gemini') {
          const response = await axios.post(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${aiApiKey}`, {
            contents: [{ parts: [{ text: `${promptText}\n\nПользователь: ${message}` }] }]
          });
          reply = response.data.candidates[0].content.parts[0].text;
        }
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', response: reply }));
      } else {
        let reply = "Привет! Я твой ИИ-инженер по интеграциям. Похоже, в этом проекте не настроен API-ключ ИИ.\n\nРекомендации по Roistat:\n- Ошибка 403: неверный API-ключ или ID проекта.\n- Ошибка 400/404: неверные эндпоинты или несовпадающие параметры.";
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', response: reply }));
      }
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // GET /api/dashboards/:dashId/data
  const dashDataParams = getUrlParams('/api/dashboards/:dashId/data', pathname);
  if (dashDataParams && req.method === 'GET') {
    const { dashId } = dashDataParams;
    const dashData = dbData.dashboardsData[dashId] || { plans: {}, changes: [], abTests: [] };
    res.statusCode = 200;
    res.end(JSON.stringify({ status: 'success', data: dashData }));
    return;
  }

  // POST /api/dashboards/:dashId/changes
  const dashChangesParams = getUrlParams('/api/dashboards/:dashId/changes', pathname);
  if (dashChangesParams && req.method === 'POST') {
    try {
      const { dashId } = dashChangesParams;
      const body = await parseBody(req);
      if (!dbData.dashboardsData[dashId]) dbData.dashboardsData[dashId] = { plans: {}, changes: [], abTests: [] };
      dbData.dashboardsData[dashId].changes.push(body);
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success' }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // DELETE /api/dashboards/:dashId/changes/:changeId
  const dashChangeDelParams = getUrlParams('/api/dashboards/:dashId/changes/:changeId', pathname);
  if (dashChangeDelParams && req.method === 'DELETE') {
    const { dashId, changeId } = dashChangeDelParams;
    if (dbData.dashboardsData[dashId]) {
      dbData.dashboardsData[dashId].changes = dbData.dashboardsData[dashId].changes.filter(c => c.id !== changeId);
      saveDB();
    }
    res.statusCode = 200;
    res.end(JSON.stringify({ status: 'success' }));
    return;
  }

  // POST /api/dashboards/:dashId/abtests
  const dashAbParams = getUrlParams('/api/dashboards/:dashId/abtests', pathname);
  if (dashAbParams && req.method === 'POST') {
    try {
      const { dashId } = dashAbParams;
      const body = await parseBody(req);
      if (!dbData.dashboardsData[dashId]) dbData.dashboardsData[dashId] = { plans: {}, changes: [], abTests: [] };
      dbData.dashboardsData[dashId].abTests.push(body);
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success' }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // DELETE /api/dashboards/:dashId/abtests/:testId
  const dashAbDelParams = getUrlParams('/api/dashboards/:dashId/abtests/:testId', pathname);
  if (dashAbDelParams && req.method === 'DELETE') {
    const { dashId, testId } = dashAbDelParams;
    if (dbData.dashboardsData[dashId]) {
      dbData.dashboardsData[dashId].abTests = dbData.dashboardsData[dashId].abTests.filter(t => t.id !== testId);
      saveDB();
    }
    res.statusCode = 200;
    res.end(JSON.stringify({ status: 'success' }));
    return;
  }

  // GET /api/rnp-months
  if (pathname === '/api/rnp-months' && req.method === 'GET') {
    try {
      const months = getAvailableMonths();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', months }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // GET /api/rnp-data
  if (pathname === '/api/rnp-data' && req.method === 'GET') {
    const fromStr = query.from;
    const toStr = query.to;
    const projId = query.projectId;
    const dashId = query.dashboardId;
    const sourceFilter = query.source || '';

    if (!fromStr || !toStr) {
      res.statusCode = 400;
      res.end(JSON.stringify({ status: 'error', message: 'Parameters from and to are required (YYYY-MM-DD)' }));
      return;
    }

    try {
      const proj = dbData.projects[projId];
      const dashboard = proj?.dashboards?.find(d => d.id === dashId);
      
      let roistatProjectId = '';
      let roistatKey = '';

      if (dashboard && dashboard.connectionId && proj && proj.connections) {
        const conn = proj.connections.find(c => c.id === dashboard.connectionId);
        if (conn && conn.type === 'roistat') {
          roistatProjectId = conn.params.projectId || '';
          roistatKey = conn.params.apiKey || '';
        }
      }
      
      // Fallback
      if (!roistatProjectId) {
        roistatProjectId = proj?.settings?.roistatId || '';
        roistatKey = proj?.settings?.roistatKey || '';
      }
      
      const customPath = (projId && dashId) ? path.join(UPLOADS_DIR, `${projId}_${dashId}.xlsx`) : null;
      const hasCustomExcel = customPath && fs.existsSync(customPath);
      
      // Fallback for demo Kilowatt project or missing projId
      const isDemo = (projId === 'proj_kilowatt' || projId === 'demo' || !projId);
      if (isDemo && !roistatProjectId) {
        roistatProjectId = '298115';
        roistatKey = '9d791b751c1b7c51d847914ec3c2b47e';
      }

      // Check empty state
      if (!roistatProjectId && !hasCustomExcel) {
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'empty_state', message: 'Источник данных не подключен. Пожалуйста, настройте интеграцию или загрузите шаблон.' }));
        return;
      }

      let activeChannels = isDemo ? [...KILOWATT_CHANNELS] : [...DEFAULT_CHANNELS];
      // 1. Get base data from Excel (both Plan and Fact)
      let mergedData = getExcelData(fromStr, toStr, hasCustomExcel ? customPath : null, isDemo, activeChannels);

      // 2. Fetch from Roistat using Project specific integrations
      let roistatSuccess = false;
      let customQualMetric = null;
      let customKpMetric = null;

      try {
        // Dynamic custom metrics discovery
        try {
          const projectMetricsList = await fetchProjectMetrics(roistatProjectId, roistatKey);
          if (projectMetricsList && projectMetricsList.length > 0) {
            for (const m of projectMetricsList) {
              const title = (m.title || "").toLowerCase();
              const name = m.name;
              if (!customQualMetric && (title.includes("квал") || title.includes("qual"))) {
                customQualMetric = name;
              }
              if (!customKpMetric && (title.includes("кп ") || title.includes(" кп") || title === "кп" || title.includes("коммерческ") || title.includes("kp"))) {
                customKpMetric = name;
              }
            }
          }
        } catch (e) {
          console.error("Error discovering Roistat metrics:", e.message);
        }

        const metricsToRequest = [
          { "metric": "visits", "attribution": "default" },
          { "metric": "leads", "attribution": "default" },
          { "metric": "leadCount", "attribution": "default" },
          { "metric": "sales", "attribution": "default" },
          { "metric": "paidLeadCount", "attribution": "default" },
          { "metric": "revenue", "attribution": "default" },
          { "metric": "paidLeadsPrice", "attribution": "default" },
          { "metric": "marketing_cost", "attribution": "default" },
          { "metric": "visitsCost", "attribution": "default" }
        ];

        if (customQualMetric) {
          metricsToRequest.push({ "metric": customQualMetric, "attribution": "default" });
        } else if (isDemo) {
          metricsToRequest.push({ "metric": "custom_2", "attribution": "default" });
        }

        if (customKpMetric) {
          metricsToRequest.push({ "metric": customKpMetric, "attribution": "default" });
        } else if (isDemo) {
          metricsToRequest.push({ "metric": "custom_5", "attribution": "default" });
        }

        const roistatItems = [];
        const fromIso = `${fromStr}T00:00:00+03:00`;
        const toIso = `${toStr}T23:59:59+03:00`;
        const payload = {
          "dimensions": ["daily", "source_group", "marker_level_1", "marker_level_2", "marker_level_3", "marker_level_4"],
          "metrics": metricsToRequest,
          "period": { "from": fromIso, "to": toIso }
        };
        const dataStr = JSON.stringify(payload);
        const options = {
          hostname: 'cloud.roistat.com',
          port: 443,
          path: `/api/v1/project/analytics/data?project=${roistatProjectId}`,
          method: 'POST',
          headers: {
            'Api-key': roistatKey,
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(dataStr)
          },
          rejectUnauthorized: false
        };

        const fetchAllData = () => {
          return new Promise((resolve, reject) => {
            const makeRequest = (retryCount = 0) => {
              const reqPost = https.request(options, (resPost) => {
                let body = '';
                resPost.setEncoding('utf8');
                resPost.on('data', (chunk) => { body += chunk; });
                resPost.on('end', () => {
                  try {
                    const parsed = JSON.parse(body);
                    if (parsed.status === 'error') {
                      if (parsed.description && (parsed.description.includes('лимит') || parsed.description.includes('limit') || parsed.description.includes('hang up')) && retryCount < 3) {
                        const delay = 150 + Math.random() * 100 + retryCount * 100;
                        setTimeout(() => makeRequest(retryCount + 1), delay);
                      } else {
                        reject(new Error(parsed.description || 'Unknown Roistat error'));
                      }
                    } else {
                      const items = (parsed.data && parsed.data[0] && parsed.data[0].items) ? parsed.data[0].items : [];
                      resolve(items);
                    }
                  } catch (e) {
                    reject(new Error(`Parse error: ${e.message}`));
                  }
                });
              });
              reqPost.on('error', (e) => {
                if (retryCount < 3) {
                  const delay = 150 + Math.random() * 100 + retryCount * 100;
                  setTimeout(() => makeRequest(retryCount + 1), delay);
                } else {
                  reject(e);
                }
              });
              reqPost.write(dataStr);
              reqPost.end();
            };
            makeRequest();
          });
        };

        const items = await fetchAllData();
        roistatItems.push(...items);
        roistatSuccess = true;
        
        if (roistatItems.length > 0) {
          const roistatByDate = {}; // dateStr -> group -> source -> metrics

          roistatItems.forEach(item => {
            const d = item.dimensions || {};
            const rawDate = d.daily?.title || d.date?.title;
            if (!rawDate) return;
            const dateStr = rawDate.split(' ')[0].split('T')[0];
            
            const marker1 = d.marker_level_1?.title || "";
            const titleRaw = [marker1, d.marker_level_2?.title || "", d.marker_level_3?.title || "", d.marker_level_4?.title || ""].join(" ").trim();
            const titleLower = titleRaw.toLowerCase();

            if (sourceFilter && !titleLower.includes(sourceFilter.toLowerCase())) return;

            const sourceGroupTitle = d.source_group?.title || "";
            const group = getGroupForProject(titleRaw, sourceGroupTitle, isDemo);
            if (!group) return;
            
            // Add group to active channels if not exists
            if (!activeChannels.includes(group)) {
              activeChannels.push(group);
              mergedData.forEach(dayItem => {
                if (!dayItem.channels[group]) dayItem.channels[group] = { plan: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }, fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }, sources: {} };
              });
            }

            const sourceName = marker1 || "Неизвестно";

            const m = {};
            item.metrics.forEach(x => { m[x.metric_name] = x.value; });

            const cost = m.visitsCost || m.marketing_cost || 0;
            const visits = m.visitCount || m.visits || 0;
            const leads = titleRaw.toLowerCase().includes("звонок") ? 0 : (m.leadCount || m.leads || 0);
            
            const qual = (customQualMetric && m[customQualMetric]) || m.custom_2 || 0;
            const kp = (customKpMetric && m[customKpMetric]) || m.custom_5 || 0;
            const sales = m.paidLeadCount || m.sales || 0;
            const rev = m.paidLeadsPrice || m.revenue || 0;

            if (!roistatByDate[dateStr]) roistatByDate[dateStr] = {};
            if (!roistatByDate[dateStr][group]) roistatByDate[dateStr][group] = { fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }, sources: {} };
            if (!roistatByDate[dateStr][group].sources[sourceName]) roistatByDate[dateStr][group].sources[sourceName] = { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 };

            const gFact = roistatByDate[dateStr][group].fact;
            gFact.cost += cost; gFact.visits += visits; gFact.leads += leads; gFact.qual += qual; gFact.kp += kp; gFact.sales += sales; gFact.rev += rev;
            
            const sFact = roistatByDate[dateStr][group].sources[sourceName];
            sFact.cost += cost; sFact.visits += visits; sFact.leads += leads; sFact.qual += qual; sFact.kp += kp; sFact.sales += sales; sFact.rev += rev;
          });

          mergedData.forEach(dayItem => {
            const rDay = roistatByDate[dayItem.date];
            if (rDay) {
              for (const channelName of activeChannels) {
                if (rDay[channelName]) {
                  dayItem.channels[channelName].fact = rDay[channelName].fact;
                  dayItem.channels[channelName].sources = rDay[channelName].sources;
                }
              }
            }
          });
        }
        
      } catch (err) {
        console.error("Roistat fetch error:", err.message);
        if (isDemo) {
          // Fallback: mock data for demo project
          const mockFile = path.join(__dirname, 'roistat_april_mock.json');
          if (fs.existsSync(mockFile)) {
            try {
              const mockData = JSON.parse(fs.readFileSync(mockFile, 'utf8'));
              const mockMap = {};
              mockData.forEach(d => { mockMap[d.date] = d.channels; });
              mergedData.forEach(dayItem => {
                if (mockMap[dayItem.date]) {
                  roistatSuccess = true;
                  for (const channelName of activeChannels) {
                    if (mockMap[dayItem.date][channelName]) {
                      dayItem.channels[channelName].fact = mockMap[dayItem.date][channelName];
                    }
                  }
                }
              });
            } catch (e) {}
          }
        }
      }

      // 3. Apply custom daily plans overrides (from project database)
      const dashPlans = dbData.dashboardsData[dashId]?.plans || {};
      mergedData.forEach(dayItem => {
        const dateStr = dayItem.date;
        if (dashPlans[dateStr]) {
          const dailyOver = dashPlans[dateStr];
          for (const channelName of activeChannels) {
            if (dailyOver[channelName]) {
              const ch = dayItem.channels[channelName];
              for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
                if (dailyOver[channelName][metric] !== undefined && dailyOver[channelName][metric] !== null) {
                  ch.plan[metric] = dailyOver[channelName][metric];
                }
              }
            }
          }
        }
      });

      // 4. Compute summaries
      const summary = {};
      const total = {
        plan: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 },
        fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }
      };

      for (const channelName of activeChannels) {
        summary[channelName] = {
          plan: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 },
          fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }
        };
      }

      mergedData.forEach(dayItem => {
        for (const [channelName, ch] of Object.entries(dayItem.channels)) {
          if (!summary[channelName]) continue;
          for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
            summary[channelName].plan[metric] += ch.plan[metric] || 0;
            summary[channelName].fact[metric] += ch.fact[metric] || 0;
            total.plan[metric] += ch.plan[metric] || 0;
            total.fact[metric] += ch.fact[metric] || 0;
          }
        }
      });

      for (const channelName of activeChannels) {
        const ch = summary[channelName];
        for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
          ch.plan[metric] = Math.round(ch.plan[metric] * 100) / 100;
          ch.fact[metric] = Math.round(ch.fact[metric] * 100) / 100;
        }
        ch.plan.calculated = computeFormulas(ch.plan);
        ch.fact.calculated = computeFormulas(ch.fact);
      }

      for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
        total.plan[metric] = Math.round(total.plan[metric] * 100) / 100;
        total.fact[metric] = Math.round(total.fact[metric] * 100) / 100;
      }
      total.plan.calculated = computeFormulas(total.plan);
      total.fact.calculated = computeFormulas(total.fact);

      const cityData = getMockCityData(total.fact.leads, total.fact.sales);

      res.statusCode = 200;
      res.end(JSON.stringify({
        status: 'success',
        roistatConnected: roistatSuccess,
        data: { channels: activeChannels, daily: mergedData, summary, total, city: cityData }
      }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // GET /api/rnp-plans?month=YYYY-MM
  if (pathname === '/api/rnp-plans' && req.method === 'GET') {
    const monthStr = query.month;
    const dashId = query.dashboardId;
    if (!monthStr) {
      res.statusCode = 400;
      res.end(JSON.stringify({ status: 'error', message: 'Parameter month is required (YYYY-MM)' }));
      return;
    }

    try {
      const dashPlans = dbData.dashboardsData[dashId]?.plans || {};
      if (dashPlans[monthStr]) {
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', month: monthStr, plans: dashPlans[monthStr] }));
        return;
      }

      // Sum plans from excel
      const [year, month] = monthStr.split('-').map(Number);
      const fromStr = `${year}-${String(month).padStart(2, '0')}-01`;
      const lastDay = new Date(year, month, 0).getDate();
      const toStr = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
      const isDemo = (dashId === "dash_rnp_kwt");
      const baseChannels = isDemo ? [...KILOWATT_CHANNELS] : [...DEFAULT_CHANNELS];
      const excelData = getExcelData(fromStr, toStr, null, true, baseChannels);

      const sumPlans = {};
      for (const channelName of baseChannels) {
        sumPlans[channelName] = { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 };
      }
      excelData.forEach(dayItem => {
        for (const [channelName, ch] of Object.entries(dayItem.channels)) {
          for (const metric of Object.keys(sumPlans[channelName])) {
            sumPlans[channelName][metric] += ch.plan[metric] || 0;
          }
        }
      });
      for (const channelName of baseChannels) {
        for (const metric of Object.keys(sumPlans[channelName])) {
          sumPlans[channelName][metric] = Math.round(sumPlans[channelName][metric]);
        }
      }

      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', month: monthStr, plans: sumPlans, source: 'excel' }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // POST /api/rnp-plans
  if (pathname === '/api/rnp-plans' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { month, plans, dashboardId } = body;
      if (!dbData.dashboardsData[dashboardId]) {
        dbData.dashboardsData[dashboardId] = { plans: {}, changes: [], abTests: [] };
      }
      dbData.dashboardsData[dashboardId].plans[month] = plans;
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', message: 'Plans saved successfully' }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  // POST /api/save-daily-plan
  if (pathname === '/api/save-daily-plan' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { date, channel, metric, val, dashboardId } = body;
      if (!dbData.dashboardsData[dashboardId]) {
        dbData.dashboardsData[dashboardId] = { plans: {}, changes: [], abTests: [] };
      }
      if (!dbData.dashboardsData[dashboardId].plans[date]) {
        dbData.dashboardsData[dashboardId].plans[date] = {};
      }
      if (!dbData.dashboardsData[dashboardId].plans[date][channel]) {
        dbData.dashboardsData[dashboardId].plans[date][channel] = {};
      }
      dbData.dashboardsData[dashboardId].plans[date][channel][metric] = val === null ? null : parseFloat(val);
      saveDB();
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success' }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message }));
    }
    return;
  }

  res.statusCode = 404;
  res.end(JSON.stringify({ status: 'error', message: `Endpoint ${req.method} ${pathname} not found` }));
}

// Request dispatcher
const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  if (pathname.startsWith('/api/')) {
    handleApi(req, res, pathname, parsedUrl.query);
    return;
  }

  // Normalize static file requests
  let filePath = path.join(PUBLIC_DIR, pathname === '/' ? 'index.html' : pathname);

  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.statusCode = 403;
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.end('Access Denied');
    return;
  }

  const ext = path.extname(filePath);
  let contentType = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (err.code === 'ENOENT') {
        res.statusCode = 404;
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.end('<h1>404 Not Found</h1><p>The requested file was not found on this server.</p>');
      } else {
        res.statusCode = 500;
        res.setHeader('Content-Type', 'text/plain; charset=utf-8');
        res.end(`Internal Server Error: ${err.code}`);
      }
    } else {
      res.statusCode = 200;
      res.setHeader('Content-Type', contentType);
      res.end(data);
    }
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n==================================================`);
  console.log(`🚀 Showcase Server is successfully running!`);
  console.log(`🔗 Localhost link: http://localhost:${PORT}`);
  console.log(`==================================================\n`);
});
