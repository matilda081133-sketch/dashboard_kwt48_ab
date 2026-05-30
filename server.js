const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const https = require('https');

const PORT = 5000;
const PUBLIC_DIR = path.join(__dirname, 'public');
const PLANS_FILE = path.join(__dirname, 'plans.json');

// Dynamically check parent directory or current directory for the Excel template
let excelPath = path.join(__dirname, '..', 'rnp_template.xlsx');
if (!fs.existsSync(excelPath)) {
  const localPath = path.join(__dirname, 'rnp_template.xlsx');
  if (fs.existsSync(localPath)) {
    excelPath = localPath;
  }
}

let XLSX;

// Dynamically import XLSX module at startup
import('../node_modules/xlsx/xlsx.mjs')
  .then(module => {
    XLSX = module;
    console.log('✅ XLSX module dynamically loaded.');
  })
  .catch(err => {
    console.error('❌ Failed to load XLSX module:', err);
  });

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

  // 1. МЕССЕНДЖЕРЫ (высший приоритет в dashboard_kilovatt.gs) -> Телеграм в РНП
  if (
    t.indexOf("tgapi") !== -1 ||
    t.indexOf("whatsapp") !== -1 ||
    t === "телеграм" ||
    ((t.indexOf("max") !== -1 || t.indexOf("макс") !== -1) && /7\d{9}/.test(t)) ||
    /79\d{9}/.test(t)
  ) {
    return "Телеграм";
  }

  // 2. ЯНДЕКС КАРТЫ -> Сайт, органика в РНП
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

  // 3. ДИРЕКТ -> Контекстная реклама в РНП
  if (
    t.indexOf("direct") !== -1 ||
    t.indexOf("директ") !== -1 ||
    t.indexOf("звонок из директа") !== -1
  ) {
    return "Контекстная реклама";
  }

  // 4. КАНАЛ MAX (без номера телефона) -> МАХ в РНП
  if (
    t === "max" ||
    t === "макс" ||
    t === "мах" ||
    t === "макс канал" ||
    t === "max канал"
  ) {
    return "МАХ";
  }

  // 5. SEO -> Сайт, органика в РНП
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

  // 6. ВЫСТАВКА -> Специализированные выставки в РНП
  if (t.indexOf("выставк") !== -1) {
    return "Специализированные выставки";
  }

  // 7. ВК -> Вконтакте в РНП
  if (
    t.indexOf("vk") !== -1 ||
    t.indexOf("вк") !== -1 ||
    t.indexOf("vkontakte") !== -1
  ) {
    return "Вконтакте";
  }

  // 8. ТГ КАНАЛ -> Телеграм в РНП
  if (
    t.indexOf("telegram") !== -1 ||
    t === "tg" ||
    t === "тг"
  ) {
    return "Телеграм";
  }

  // 9. СТАТЬИ В СМИ -> Статьи в СМИ в РНП (не зафиксировано в getGroup киловатт, но нужно для РНП)
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
function getExcelData(fromStr, toStr) {
  if (!XLSX) {
    console.warn('⚠️ XLSX module is not loaded yet. Returning empty array.');
    return [];
  }
  if (!fs.existsSync(excelPath)) {
    console.error('❌ Spreadsheet rnp_template.xlsx not found at:', excelPath);
    return [];
  }

  const fileBuffer = fs.readFileSync(excelPath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const dates = {}; // dateStr -> { sheetName, colIdx }

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

    for (const channelName of Object.keys(GROUPS_ROWS)) {
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
    const PID = '294460';
    const KEY = '87ed258c066c98668b595bafa0365e56';

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
async function handleApi(req, res, pathname, query) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

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

  // POST /api/add-month
  if (pathname === '/api/add-month' && req.method === 'POST') {
    try {
      const body = await parseBody(req);
      const { month } = body;
      
      if (!month || !month.match(/^\d{4}-\d{2}$/)) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Parameter month is required in YYYY-MM format' }));
        return;
      }
      
      const customMonthsFile = path.join(__dirname, 'custom_months.json');
      let custom = [];
      if (fs.existsSync(customMonthsFile)) {
        try {
          custom = JSON.parse(fs.readFileSync(customMonthsFile, 'utf8'));
        } catch (e) {}
      }
      if (!custom.includes(month)) {
        custom.push(month);
        fs.writeFileSync(customMonthsFile, JSON.stringify(custom, null, 2), 'utf8');
      }
      
      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', message: `Month ${month} added successfully` }));
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

    if (!fromStr || !toStr) {
      res.statusCode = 400;
      res.end(JSON.stringify({ status: 'error', message: 'Parameters from and to are required (YYYY-MM-DD)' }));
      return;
    }

    try {
      // 1. Get base data from Excel (both Plan and Fact)
      let mergedData = getExcelData(fromStr, toStr);

      // 2. Fetch from Roistat
      let roistatSuccess = false;
      try {
        const roistatItems = await fetchRoistat(fromStr, toStr);
        console.log(`🌐 Fetched ${roistatItems.length} items from Roistat API.`);
        
        if (roistatItems.length > 0) {
          roistatSuccess = true;
          const roistatByDate = {};
          
          roistatItems.forEach(item => {
            const d = item.dimensions || {};
            if (!d.date || !d.date.title) return;
            const dateStr = d.date.title.split('T')[0];
            
            const t1 = d.marker_level_1 ? d.marker_level_1.title || "" : "";
            const t2 = d.marker_level_2 ? d.marker_level_2.title || "" : "";
            const t3 = d.marker_level_3 ? d.marker_level_3.title || "" : "";
            const t4 = d.marker_level_4 ? d.marker_level_4.title || "" : "";
            const title = [t1, t2, t3, t4].join(" ").trim();
            
            const group = getGroup(title);
            if (!group) return;

            const m = {};
            item.metrics.forEach(x => { m[x.metric_name] = x.value; });

            const cost = m.visitsCost || m.marketing_cost || 0;
            const visits = m.visitCount || m.visits || 0;
            const isCall = title.toLowerCase().indexOf("звонок") !== -1;
            const leads = isCall ? 0 : (m.leadCount || m.leads || 0);
            const qual = m.custom_2 || 0;
            const kp = m.custom_5 || 0;
            const sales = m.paidLeadCount || m.sales || 0;
            const rev = m.paidLeadsPrice || m.revenue || 0;

            if (!roistatByDate[dateStr]) roistatByDate[dateStr] = {};
            if (!roistatByDate[dateStr][group]) {
              roistatByDate[dateStr][group] = { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 };
            }

            const g = roistatByDate[dateStr][group];
            g.cost += cost;
            g.visits += visits;
            g.leads += leads;
            g.qual += qual;
            g.kp += kp;
            g.sales += sales;
            g.rev += rev;
          });

          // Overwrite Fact data
          mergedData.forEach(dayItem => {
            const rDay = roistatByDate[dayItem.date];
            if (rDay) {
              for (const channelName of Object.keys(GROUPS_ROWS)) {
                const rCh = rDay[channelName];
                if (rCh) {
                  dayItem.channels[channelName].fact = rCh;
                } else {
                  dayItem.channels[channelName].fact = { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 };
                }
              }
            }
          });
        }
      } catch (err) {
        console.warn('⚠️ Roistat API connection bypassed/failed. Reason:', err.message);
        
        // Local fallback to mock Roistat data for April 2026 if available
        const mockFile = path.join(__dirname, 'roistat_april_mock.json');
        if (fs.existsSync(mockFile)) {
          try {
            const mockData = JSON.parse(fs.readFileSync(mockFile, 'utf8'));
            const mockMap = {};
            mockData.forEach(d => { mockMap[d.date] = d.channels; });
            
            let mockApplied = false;
            mergedData.forEach(dayItem => {
              const dateStr = dayItem.date;
              if (mockMap[dateStr]) {
                mockApplied = true;
                const mockChs = mockMap[dateStr];
                for (const channelName of Object.keys(GROUPS_ROWS)) {
                  if (mockChs[channelName]) {
                    dayItem.channels[channelName].fact = mockChs[channelName];
                  }
                }
              }
            });
            if (mockApplied) {
              console.log('✓ Successfully loaded mock Roistat facts for April 2026 from roistat_april_mock.json');
              roistatSuccess = true;
            }
          } catch (e) {
            console.error('Error loading mock Roistat facts:', e);
          }
        }
      }

      // 3. Apply custom daily plans overrides from plans.json
      const plans = readPlans();
      mergedData.forEach(dayItem => {
        const dateStr = dayItem.date;
        
        // 3a. Check if there are daily overrides for this specific date
        if (plans[dateStr]) {
          const dailyOver = plans[dateStr];
          for (const channelName of Object.keys(GROUPS_ROWS)) {
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
        
        // 3b. Check if there are monthly plan overrides (legacy fallback)
        const monthStr = dateStr.slice(0, 7); // "YYYY-MM"
        if (plans[monthStr] && !plans[dateStr]) {
          const monthlyPlan = plans[monthStr];
          const year = parseInt(dateStr.split('-')[0]);
          const month = parseInt(dateStr.split('-')[1]);
          const daysInMonth = new Date(year, month, 0).getDate();

          for (const channelName of Object.keys(GROUPS_ROWS)) {
            if (monthlyPlan[channelName]) {
              const chPlan = monthlyPlan[channelName];
              for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
                if (chPlan[metric] !== undefined) {
                  dayItem.channels[channelName].plan[metric] = (chPlan[metric] || 0) / daysInMonth;
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

      for (const channelName of Object.keys(GROUPS_ROWS)) {
        summary[channelName] = {
          plan: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 },
          fact: { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 }
        };
      }

      mergedData.forEach(dayItem => {
        for (const [channelName, ch] of Object.entries(dayItem.channels)) {
          for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
            summary[channelName].plan[metric] += ch.plan[metric] || 0;
            summary[channelName].fact[metric] += ch.fact[metric] || 0;

            total.plan[metric] += ch.plan[metric] || 0;
            total.fact[metric] += ch.fact[metric] || 0;
          }
        }
      });

      // Format summary and add calculations
      for (const channelName of Object.keys(GROUPS_ROWS)) {
        const ch = summary[channelName];
        
        // Round raw totals
        for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
          ch.plan[metric] = Math.round(ch.plan[metric] * 100) / 100;
          ch.fact[metric] = Math.round(ch.fact[metric] * 100) / 100;
        }

        ch.plan.calculated = computeFormulas(ch.plan);
        ch.fact.calculated = computeFormulas(ch.fact);
      }

      // Round total raw metrics
      for (const metric of ['cost', 'visits', 'leads', 'qual', 'kp', 'sales', 'rev']) {
        total.plan[metric] = Math.round(total.plan[metric] * 100) / 100;
        total.fact[metric] = Math.round(total.fact[metric] * 100) / 100;
      }
      total.plan.calculated = computeFormulas(total.plan);
      total.fact.calculated = computeFormulas(total.fact);

      // 5. Generate City data
      const cityData = getMockCityData(total.fact.leads, total.fact.sales);

      res.statusCode = 200;
      res.end(JSON.stringify({
        status: 'success',
        roistatConnected: roistatSuccess,
        data: {
          daily: mergedData,
          summary,
          total,
          city: cityData
        }
      }));
    } catch (err) {
      res.statusCode = 500;
      res.end(JSON.stringify({ status: 'error', message: err.message, stack: err.stack }));
    }
    return;
  }

  // GET /api/rnp-plans?month=YYYY-MM
  if (pathname === '/api/rnp-plans' && req.method === 'GET') {
    const monthStr = query.month; // e.g. "2026-04"
    if (!monthStr) {
      res.statusCode = 400;
      res.end(JSON.stringify({ status: 'error', message: 'Parameter month is required (YYYY-MM)' }));
      return;
    }

    try {
      const plans = readPlans();
      if (plans[monthStr]) {
        res.statusCode = 200;
        res.end(JSON.stringify({ status: 'success', month: monthStr, plans: plans[monthStr] }));
        return;
      }

      // Fallback: sum from Excel
      console.log(`Summing plans for ${monthStr} from Excel...`);
      const [year, month] = monthStr.split('-').map(Number);
      const fromStr = `${year}-${String(month).padStart(2, '0')}-01`;
      const lastDay = new Date(year, month, 0).getDate();
      const toStr = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;

      const excelData = getExcelData(fromStr, toStr);

      const sumPlans = {};
      for (const channelName of Object.keys(GROUPS_ROWS)) {
        sumPlans[channelName] = { cost: 0, visits: 0, leads: 0, qual: 0, kp: 0, sales: 0, rev: 0 };
      }

      excelData.forEach(dayItem => {
        for (const [channelName, ch] of Object.entries(dayItem.channels)) {
          for (const metric of Object.keys(sumPlans[channelName])) {
            sumPlans[channelName][metric] += ch.plan[metric] || 0;
          }
        }
      });

      // Round sums
      for (const channelName of Object.keys(GROUPS_ROWS)) {
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
      const { month, plans } = body;

      if (!month || !plans) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Parameters month (YYYY-MM) and plans are required' }));
        return;
      }

      const allPlans = readPlans();
      allPlans[month] = plans;
      writePlans(allPlans);

      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', message: `Plans for ${month} saved successfully` }));
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
      const { date, channel, metric, val } = body;

      if (!date || !channel || !metric) {
        res.statusCode = 400;
        res.end(JSON.stringify({ status: 'error', message: 'Parameters date (YYYY-MM-DD), channel, and metric are required' }));
        return;
      }

      const allPlans = readPlans();
      if (!allPlans[date]) {
        allPlans[date] = {};
      }
      if (!allPlans[date][channel]) {
        allPlans[date][channel] = {};
      }
      allPlans[date][channel][metric] = val === null ? null : parseFloat(val);
      writePlans(allPlans);

      res.statusCode = 200;
      res.end(JSON.stringify({ status: 'success', message: `Plan for ${channel} on ${date} saved successfully` }));
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

server.listen(PORT, () => {
  console.log(`\n==================================================`);
  console.log(`🚀 Showcase Server is successfully running!`);
  console.log(`🔗 Localhost link: http://localhost:${PORT}`);
  console.log(`==================================================\n`);
});
