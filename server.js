// Electricity Price Server
// Dohvaca cijene svakih sat vremena, cuva ih i servira web stranicu
// Start: node server.js

const https      = require("https");
const express    = require("express");
const cron       = require("node-cron");
const path       = require("path");
const fs         = require("fs");
const nodemailer = require("nodemailer");
const XLSX       = require("xlsx");

const app  = express();
const PORT = process.env.PORT || 4000;

// ENTSO-E token za CROPEX (HR) - opciono
const ENTSOE_TOKEN = process.env.ENTSOE_TOKEN || "";

// ─────────────────────────────────────────────
// Email konfiguracija
// ─────────────────────────────────────────────
const EMAIL_FROM  = process.env.EMAIL_FROM  || "";
const EMAIL_PASS  = process.env.EMAIL_PASS  || "";
const EMAIL_TO    = process.env.EMAIL_TO    || "";

// Putanja gdje se cuvaju podaci
const DATA_FILE = path.join(__dirname, "data.json");

// ─────────────────────────────────────────────
// HTTP helper
// ─────────────────────────────────────────────
function get(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith("https") ? https : require("http");
    const req = lib.get(url, {
      headers: { "User-Agent": "Mozilla/5.0", "Accept": "application/json,*/*" },
      timeout: 15000,
    }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
        return get(res.headers.location).then(resolve).catch(reject);
      }
      let data = "";
      res.on("data", (c) => (data += c));
      res.on("end", () => resolve({ status: res.statusCode, body: data }));
    });
    req.on("error", reject);
    req.on("timeout", () => { req.destroy(); reject(new Error("Timeout")); });
  });
}

// ─────────────────────────────────────────────
// CET/CEST konverzija (Europe/Budapest = HR, HU, RS, SI, DE)
// ─────────────────────────────────────────────
function getCET(dt) {
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Budapest",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", hour12: false,
  });
  const parts = fmt.formatToParts(dt);
  const get = type => parts.find(p => p.type === type).value;
  const hour = parseInt(get("hour")) % 24; // 24 -> 0 (ponoc)
  return {
    datum: `${get("year")}-${get("month")}-${get("day")}`,
    sat:   `${String(hour).padStart(2, "0")}:00`,
  };
}

// ─────────────────────────────────────────────
// Dohvacanje podataka
// ─────────────────────────────────────────────
async function fetchZone(zone, label, daysBack = 1) {
  const today    = new Date();
  const startDay = new Date(today);
  startDay.setDate(startDay.getDate() - daysBack);
  const startStr = startDay.toISOString().slice(0, 10);

  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().slice(0, 10);
  const url = `https://api.energy-charts.info/price?bzn=${zone}&start=${startStr}T00:00Z&end=${tomorrowStr}T23:00Z`;
  const res = await get(url);
  if (res.status !== 200) throw new Error(`HTTP ${res.status}`);

  const { unix_seconds, price } = JSON.parse(res.body);

  // Grupisanje u satne prosjeke (API vraca 15-min intervale), sati u CET/CEST
  const byHour = {};
  unix_seconds.forEach((ts, i) => {
    if (price[i] === null) return;
    const { datum, sat } = getCET(new Date(ts * 1000));
    const key = `${datum}_${sat}`;
    if (!byHour[key]) byHour[key] = { datum, sat, sum: 0, count: 0 };
    byHour[key].sum += price[i];
    byHour[key].count++;
  });

  const data = Object.values(byHour)
    .map(r => ({ datum: r.datum, sat: r.sat, value: Number((r.sum / r.count).toFixed(2)) }))
    .sort((a, b) => `${a.datum}${a.sat}`.localeCompare(`${b.datum}${b.sat}`));

  return { label, zone, data };
}

async function fetchCROPEX() {
  if (!ENTSOE_TOKEN || ENTSOE_TOKEN === "TVOJ_TOKEN_OVDJE") return null;

  const today     = new Date();
  const dateStr   = today.toISOString().slice(0, 10).replace(/-/g, "");
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const yStr = yesterday.toISOString().slice(0, 10).replace(/-/g, "");

  const url = `https://web-api.tp.entsoe.eu/api?documentType=A44&in_Domain=10YHR-HEP------M&out_Domain=10YHR-HEP------M&periodStart=${yStr}2300&periodEnd=${dateStr}2300&securityToken=${ENTSOE_TOKEN}`;
  const res = await get(url);
  if (res.status !== 200) return null;

  const rows = [];
  const periods = res.body.match(/<Period>([\s\S]*?)<\/Period>/g) || [];
  periods.forEach((period) => {
    const startMatch = period.match(/<start>(.*?)<\/start>/);
    if (!startMatch) return;
    const periodStart = new Date(startMatch[1]);
    const points = period.match(/<Point>([\s\S]*?)<\/Point>/g) || [];
    points.forEach((point) => {
      const pos   = point.match(/<position>(\d+)<\/position>/);
      const price = point.match(/<price\.amount>([\d.]+)<\/price\.amount>/);
      if (!pos || !price) return;
      const dt = new Date(periodStart.getTime() + (parseInt(pos[1]) - 1) * 3600000);
      rows.push({
        datum: dt.toISOString().slice(0, 10),
        sat:   `${String(dt.getUTCHours()).padStart(2, "0")}:00`,
        value: Number(Number(price[1]).toFixed(2)),
      });
    });
  });

  return { label: "CROPEX (HR)", zone: "HR", data: rows };
}

// ENTSO-E genericki fetcher (za CROPEX i HUPX)
async function fetchENTSOE(domain, label, zoneCode) {
  if (!ENTSOE_TOKEN || ENTSOE_TOKEN === "TVOJ_TOKEN_OVDJE") {
    console.log(`  ${label}: preskocen — ENTSOE_TOKEN nije postavljen`);
    return null;
  }

  const today    = new Date();
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
  // Idemo 2 dana unazad da sigurno uhvatimo danasnje sate (ENTSO-E periodicno iskljucuje tacnu granicu)
  const yStr = new Date(today.getTime() - 2 * 86400000).toISOString().slice(0, 10).replace(/-/g, "");
  const tStr = tomorrow.toISOString().slice(0, 10).replace(/-/g, "");

  const url = `https://web-api.tp.entsoe.eu/api?documentType=A44&in_Domain=${domain}&out_Domain=${domain}&periodStart=${yStr}2300&periodEnd=${tStr}2300&securityToken=${ENTSOE_TOKEN}`;
  const res = await get(url);
  if (res.status !== 200) {
    console.error(`  ${label}: HTTP ${res.status} — ${res.body.slice(0, 200)}`);
    return null;
  }

  const byHour = {};
  const periods = res.body.match(/<Period>([\s\S]*?)<\/Period>/g) || [];
  periods.forEach((period) => {
    const startMatch      = period.match(/<start>(.*?)<\/start>/);
    const resolutionMatch = period.match(/<resolution>(.*?)<\/resolution>/);
    if (!startMatch) return;
    const periodStart = new Date(startMatch[1]);
    const resolution  = resolutionMatch ? resolutionMatch[1] : "PT60M";
    const intervalMs  = resolution === "PT15M" ? 900000 : resolution === "PT30M" ? 1800000 : 3600000;

    const points = period.match(/<Point>([\s\S]*?)<\/Point>/g) || [];
    points.forEach((point) => {
      const pos   = point.match(/<position>(\d+)<\/position>/);
      const price = point.match(/<price\.amount>([\d.]+)<\/price\.amount>/);
      if (!pos || !price) return;
      const dt          = new Date(periodStart.getTime() + (parseInt(pos[1]) - 1) * intervalMs);
      const { datum, sat } = getCET(dt);
      const key         = `${datum}_${sat}`;
      if (!byHour[key]) byHour[key] = { datum, sat, sum: 0, count: 0 };
      byHour[key].sum += parseFloat(price[1]);
      byHour[key].count++;
    });
  });

  const rows = Object.values(byHour)
    .map(r => ({ datum: r.datum, sat: r.sat, value: Number((r.sum / r.count).toFixed(2)) }))
    .sort((a, b) => `${a.datum}${a.sat}`.localeCompare(`${b.datum}${b.sat}`));

  return { label, zone: zoneCode, data: rows };
}

async function fetchAllPrices() {
  console.log(`[${new Date().toISOString()}] Dohvacam cijene struje...`);
  // Krenemo od postojecih podataka — ako market padne (timeout), cuvamo stare vrijednosti
  const results = Object.assign({}, cachedData);
  delete results.updatedAt;

  try {
    const epex = await fetchZone("DE-LU", "EPEX SPOT (DE-LU)");
    results.epex = epex;
    console.log(`  EPEX: ${epex.data.length} sati`);
  } catch (e) { console.error("  EPEX greška:", e.message); }

  try {
    const seepex = await fetchZone("RS", "SEEPEX (RS)");
    results.seepex = seepex;
    console.log(`  SEEPEX: ${seepex.data.length} sati`);
  } catch (e) { console.error("  SEEPEX greška:", e.message); }

  try {
    const cropex = await fetchENTSOE("10YHR-HEP------M", "CROPEX (HR)", "HR");
    if (cropex) {
      results.cropex = cropex;
      console.log(`  CROPEX: ${cropex.data.length} sati`);
    }
  } catch (e) { console.error("  CROPEX greška:", e.message); }

  try {
    const hupx = await fetchENTSOE("10YHU-MAVIR----U", "HUPX (HU)", "HU");
    if (hupx) {
      results.hupx = hupx;
      console.log(`  HUPX: ${hupx.data.length} sati`);
    }
  } catch (e) { console.error("  HUPX greška:", e.message); }

  try {
    const mepx = await fetchZone("ME", "MEPX (ME)");
    results.mepx = mepx;
    console.log(`  MEPX: ${mepx.data.length} sati`);
  } catch (e) { console.error("  MEPX greška:", e.message); }

  try {
    const bsp = await fetchZone("SI", "BSP SouthPool (SI)");
    results.bsp = bsp;
    console.log(`  BSP: ${bsp.data.length} sati`);
  } catch (e) { console.error("  BSP greška:", e.message); }

  try {
    // HU podaci imaju ~5 dana kašnjenja, pa uzimamo 7 dana unazad
    const epexhu = await fetchZone("HU", "EPEX (HU)", 7);
    results.epexhu = epexhu;
    console.log(`  EPEX HU: ${epexhu.data.length} sati`);
  } catch (e) { console.error("  EPEX HU greška:", e.message); }

  results.updatedAt = new Date().toISOString();
  fs.writeFileSync(DATA_FILE, JSON.stringify(results, null, 2));
  console.log(`  Sacuvano u ${DATA_FILE}`);
  return results;
}

// ─────────────────────────────────────────────
// Email — Day Ahead izvještaj
// ─────────────────────────────────────────────
function buildDayAheadEmail(data) {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = getCET(tomorrow).datum; // CET datum, slaze se s pohranjenim podacima

  const markets = [
    { key: "epex",   label: "EPEX SPOT (DE-LU)", color: "#58a6ff" },
    { key: "seepex", label: "SEEPEX (RS)",        color: "#3fb950" },
    { key: "mepx",   label: "MEPX (ME)",          color: "#a371f7" },
    { key: "bsp",    label: "BSP SouthPool (SI)", color: "#39d353" },
    { key: "epexhu", label: "EPEX (HU)",           color: "#ffa657" },
    { key: "cropex", label: "CROPEX (HR)",         color: "#d29922" },
    { key: "hupx",   label: "HUPX (HU)",           color: "#f78166" },
  ];

  // Skupi sve sate (0-23)
  const hours = Array.from({ length: 24 }, (_, i) => `${String(i).padStart(2, "0")}:00`);

  // Za svaki market napravi mapu sat -> cijena za sutra
  const marketMaps = {};
  const marketStats = {};
  markets.forEach(m => {
    const src = data[m.key];
    if (!src) return;
    const dayData = src.data.filter(r => r.datum === tomorrowStr && r.value !== null);
    if (!dayData.length) return;
    const map = {};
    dayData.forEach(r => { map[r.sat] = r.value; });
    marketMaps[m.key] = map;
    const vals = dayData.map(r => r.value);
    marketStats[m.key] = {
      min: Math.min(...vals).toFixed(2),
      max: Math.max(...vals).toFixed(2),
      avg: (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2),
    };
  });

  const availableMarkets = markets.filter(m => marketMaps[m.key]);

  if (!availableMarkets.length) return null;

  // Statistike sekcija
  const statsRows = availableMarkets.map(m => {
    const s = marketStats[m.key];
    return `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #21262d;">
          <span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;background:${m.color}22;color:${m.color}">${m.label}</span>
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid #21262d;color:#3fb950;font-weight:700;text-align:right">${s.min}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #21262d;color:#f85149;font-weight:700;text-align:right">${s.max}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #21262d;color:#58a6ff;font-weight:700;text-align:right">${s.avg}</td>
      </tr>`;
  }).join("");

  // Satna tabela
  const headerCells = availableMarkets.map(m =>
    `<th style="padding:8px 12px;text-align:right;color:${m.color};font-size:12px;white-space:nowrap">${m.label}</th>`
  ).join("");

  const hourRows = hours.map(h => {
    const cells = availableMarkets.map(m => {
      const val = marketMaps[m.key][h];
      const color = val === undefined ? "#555" : val < 0 ? "#f85149" : val < 50 ? "#3fb950" : val < 100 ? "#d29922" : "#f85149";
      const txt = val !== undefined ? val.toFixed(2) : "—";
      return `<td style="padding:6px 12px;border-bottom:1px solid #21262d;text-align:right;color:${color};font-weight:${val !== undefined ? "600" : "400"}">${txt}</td>`;
    }).join("");
    return `<tr><td style="padding:6px 12px;border-bottom:1px solid #21262d;color:#8b949e;font-family:monospace">${h}</td>${cells}</tr>`;
  }).join("");

  const html = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;font-family:'Segoe UI',sans-serif;color:#e6edf3">
<div style="max-width:800px;margin:0 auto;padding:24px">

  <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;padding:24px;margin-bottom:20px">
    <h1 style="margin:0 0 4px;font-size:1.3rem;color:#f0f6fc">⚡ Day-Ahead cijene struje</h1>
    <p style="margin:0;color:#8b949e;font-size:0.85rem">Sutra — ${tomorrowStr}</p>
  </div>

  <!-- Statistike -->
  <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;overflow:hidden;margin-bottom:20px">
    <div style="padding:14px 16px;border-bottom:1px solid #21262d">
      <p style="margin:0;font-size:0.72rem;text-transform:uppercase;letter-spacing:1.5px;color:#8b949e">Pregled Min / Max / Prosjek (EUR/MWh)</p>
    </div>
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="background:#1c2128">
          <th style="padding:8px 12px;text-align:left;font-size:12px;color:#8b949e">Berza</th>
          <th style="padding:8px 12px;text-align:right;font-size:12px;color:#3fb950">Min</th>
          <th style="padding:8px 12px;text-align:right;font-size:12px;color:#f85149">Max</th>
          <th style="padding:8px 12px;text-align:right;font-size:12px;color:#58a6ff">Prosjek</th>
        </tr>
      </thead>
      <tbody>${statsRows}</tbody>
    </table>
  </div>

  <!-- Satna tabela -->
  <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;overflow:hidden;margin-bottom:20px">
    <div style="padding:14px 16px;border-bottom:1px solid #21262d">
      <p style="margin:0;font-size:0.72rem;text-transform:uppercase;letter-spacing:1.5px;color:#8b949e">Satni pregled (EUR/MWh)</p>
    </div>
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="background:#1c2128">
          <th style="padding:8px 12px;text-align:left;font-size:12px;color:#8b949e">Sat</th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>${hourRows}</tbody>
    </table>
  </div>

  <p style="text-align:center;color:#8b949e;font-size:0.75rem">
    Generirano automatski · <a href="http://localhost:4000" style="color:#58a6ff">Cijene struje</a>
  </p>
</div>
</body>
</html>`;

  return { html, tomorrowStr, availableMarkets: availableMarkets.map(m => m.label) };
}

async function sendDayAheadEmail(data) {
  if (EMAIL_FROM === "TVOJ_GMAIL@gmail.com" || EMAIL_PASS === "TVOJA_APP_LOZINKA") {
    console.log("  Email preskocen — konfigurisi EMAIL_FROM i EMAIL_PASS u server.js");
    return;
  }

  const result = buildDayAheadEmail(data);
  if (!result) {
    console.log("  Email preskocen — nema day-ahead podataka za sutra");
    return;
  }

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: EMAIL_FROM, pass: EMAIL_PASS },
  });

  await transporter.sendMail({
    from: `"Cijene Struje" <${EMAIL_FROM}>`,
    to: EMAIL_TO,
    subject: `⚡ Day-Ahead cijene struje — ${result.tomorrowStr}`,
    html: result.html,
  });

  console.log(`  Email poslan na ${EMAIL_TO} (${result.availableMarkets.join(", ")})`);
  dayAheadEmailSentDate = result.tomorrowStr; // Oznaci da je mail poslan
}

// ─────────────────────────────────────────────
// Alarm — negativne cijene
// ─────────────────────────────────────────────

// Prati koje berze su vec dobile alarm (reset kad cijena poraste iznad 0)
const negativeAlertSent = {};

// Prati je li day-ahead mail vec poslan danas
let dayAheadEmailSentDate = "";

async function checkNegativePrices(data) {
  if (EMAIL_FROM === "TVOJ_GMAIL@gmail.com" || EMAIL_PASS === "TVOJA_APP_LOZINKA") return;

  const now   = new Date();
  const today = now.toISOString().slice(0, 10);
  const hour  = `${String(now.getHours()).padStart(2, "0")}:00`;

  const markets = [
    { key: "epex",   label: "EPEX SPOT (DE-LU)" },
    { key: "seepex", label: "SEEPEX (RS)"        },
    { key: "mepx",   label: "MEPX (ME)"          },
    { key: "bsp",    label: "BSP SouthPool (SI)" },
    { key: "epexhu", label: "EPEX (HU)"           },
    { key: "cropex", label: "CROPEX (HR)"         },
    { key: "hupx",   label: "HUPX (HU)"           },
  ];

  const negativni = [];

  markets.forEach(m => {
    const src = data[m.key];
    if (!src) return;
    const row = src.data.find(r => r.datum === today && r.sat === hour);
    if (!row || row.value === null) return;

    if (row.value < 0) {
      if (!negativeAlertSent[m.key]) {
        negativeAlertSent[m.key] = true;
        negativni.push({ label: m.label, value: row.value, sat: hour });
      }
    } else {
      // Reset alarma kad cijena ponovo postane >= 0
      negativeAlertSent[m.key] = false;
    }
  });

  if (!negativni.length) return;

  const listaHTML = negativni.map(n =>
    `<tr>
      <td style="padding:8px 14px;border-bottom:1px solid #21262d;font-weight:600">${n.label}</td>
      <td style="padding:8px 14px;border-bottom:1px solid #21262d;color:#8b949e">${n.sat}</td>
      <td style="padding:8px 14px;border-bottom:1px solid #21262d;color:#f85149;font-weight:700;font-size:1.1rem">${n.value.toFixed(2)} EUR/MWh</td>
    </tr>`
  ).join("");

  const html = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;font-family:'Segoe UI',sans-serif;color:#e6edf3">
<div style="max-width:600px;margin:0 auto;padding:24px">

  <div style="background:#2d1b1b;border:1px solid #f8514966;border-radius:12px;padding:24px;margin-bottom:20px">
    <h1 style="margin:0 0 6px;font-size:1.3rem;color:#f85149">⚠️ Alarm: Negativne cijene struje!</h1>
    <p style="margin:0;color:#8b949e;font-size:0.85rem">${today} u ${hour} — cijena pala ispod 0</p>
  </div>

  <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;overflow:hidden;margin-bottom:20px">
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="background:#1c2128">
          <th style="padding:8px 14px;text-align:left;font-size:12px;color:#8b949e">Berza</th>
          <th style="padding:8px 14px;text-align:left;font-size:12px;color:#8b949e">Sat</th>
          <th style="padding:8px 14px;text-align:left;font-size:12px;color:#f85149">Cijena</th>
        </tr>
      </thead>
      <tbody>${listaHTML}</tbody>
    </table>
  </div>

  <p style="text-align:center;color:#8b949e;font-size:0.75rem">
    Automatski alarm · <a href="http://localhost:4000" style="color:#58a6ff">Cijene struje</a>
  </p>
</div>
</body>
</html>`;

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: EMAIL_FROM, pass: EMAIL_PASS },
  });

  await transporter.sendMail({
    from: `"Cijene Struje" <${EMAIL_FROM}>`,
    to: EMAIL_TO,
    subject: `⚠️ ALARM: Negativne cijene struje — ${today} ${hour}`,
    html,
  });

  console.log(`  [ALARM] Negativne cijene! Poslan mail za: ${negativni.map(n => n.label).join(", ")}`);
}

// ─────────────────────────────────────────────
// Inicijalizacija podataka
// ─────────────────────────────────────────────
let cachedData = {};

function loadCache() {
  if (fs.existsSync(DATA_FILE)) {
    try {
      cachedData = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
      console.log(`Cache ucitan iz ${DATA_FILE}`);
    } catch (e) { cachedData = {}; }
  }
}

// ─────────────────────────────────────────────
// Routes
// ─────────────────────────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());

app.get("/api/prices", (req, res) => {
  res.json(cachedData);
});

app.get("/api/refresh", async (req, res) => {
  try {
    cachedData = await fetchAllPrices();
    res.json({ ok: true, updatedAt: cachedData.updatedAt });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.get("/api/status", (req, res) => {
  const status = {
    token: ENTSOE_TOKEN ? `postavljen (${ENTSOE_TOKEN.slice(0, 8)}...)` : "NIJE POSTAVLJEN",
    email: EMAIL_FROM || "nije postavljen",
    updatedAt: cachedData.updatedAt || "nikad",
    markets: {},
  };
  ["epex", "seepex", "cropex", "hupx", "mepx", "bsp", "epexhu"].forEach(k => {
    const m = cachedData[k];
    status.markets[k] = m ? `${m.data.length} sati` : "nema podataka";
  });
  res.json(status);
});

app.get("/api/send-email", async (req, res) => {
  try {
    await sendDayAheadEmail(cachedData);
    res.json({ ok: true, message: `Email poslan na ${EMAIL_TO}` });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.post("/api/send-xls", async (req, res) => {
  try {
    const { email, markets } = req.body;
    if (!email)                 return res.status(400).json({ ok: false, error: "Email je obavezan" });
    if (!markets?.length)       return res.status(400).json({ ok: false, error: "Odaberite bar jednu berzu" });
    if (!EMAIL_FROM || !EMAIL_PASS) return res.status(500).json({ ok: false, error: "Email konfiguracija nije postavljena na serveru" });

    const tomorrow    = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = getCET(tomorrow).datum;

    const marketLabels = {
      seepex: "SEEPEX (RS)",
      hupx:   "HUPX (HU)",
      cropex: "CROPEX (HR)",
      epex:   "EPEX SPOT (DE-LU)",
      mepx:   "MEPX (ME)",
      bsp:    "BSP SouthPool (SI)",
      epexhu: "EPEX (HU)",
    };

    const selected = markets.filter(k => cachedData[k] && marketLabels[k]);
    if (!selected.length) return res.status(400).json({ ok: false, error: "Nema podataka za odabrane berze" });

    const hours = Array.from({ length: 24 }, (_, i) => `${String(i).padStart(2, "0")}:00`);

    // Satni podaci
    const rows = [["Sat (CET)", ...selected.map(k => marketLabels[k])]];
    hours.forEach(h => {
      const row = [h];
      selected.forEach(k => {
        const rec = cachedData[k]?.data.find(r => r.datum === tomorrowStr && r.sat === h);
        row.push(rec?.value !== undefined && rec?.value !== null ? rec.value : null);
      });
      rows.push(row);
    });

    // Statistike
    rows.push([]);
    rows.push(["Statistike", ...selected.map(k => marketLabels[k])]);
    ["Min", "Max", "Prosjek"].forEach(stat => {
      const row = [stat];
      selected.forEach(k => {
        const vals = (cachedData[k]?.data || [])
          .filter(r => r.datum === tomorrowStr && r.value !== null)
          .map(r => r.value);
        if (!vals.length) { row.push(null); return; }
        if (stat === "Min")    row.push(Math.min(...vals));
        else if (stat === "Max") row.push(Math.max(...vals));
        else row.push(Number((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2)));
      });
      rows.push(row);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, ...selected.map(() => ({ wch: 22 }))];
    XLSX.utils.book_append_sheet(wb, ws, `Sutra ${tomorrowStr}`);
    const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: EMAIL_FROM, pass: EMAIL_PASS },
    });

    await transporter.sendMail({
      from: `"Cijene Struje" <${EMAIL_FROM}>`,
      to: email,
      subject: `⚡ Day-Ahead cijene struje — ${tomorrowStr}`,
      text: `U prilogu su day-ahead cijene struje za sutra (${tomorrowStr}) za berze: ${selected.map(k => marketLabels[k]).join(", ")}.`,
      attachments: [{
        filename: `cijene-struje-${tomorrowStr}.xlsx`,
        content: buffer,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }],
    });

    console.log(`  XLS poslan na ${email} (${selected.join(", ")})`);
    res.json({ ok: true, message: `Excel fajl poslan na ${email}` });
  } catch (e) {
    console.error("Send XLS greška:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ─────────────────────────────────────────────
// Start
// ─────────────────────────────────────────────
loadCache();

app.listen(PORT, async () => {
  console.log(`\nServer pokrenut na http://localhost:${PORT}`);

  // Dohvati odmah pri startu ako nema cache-a
  if (!cachedData.updatedAt) {
    cachedData = await fetchAllPrices();
  }

  // Cron: svakih sat u :00
  cron.schedule("0 * * * *", async () => {
    cachedData = await fetchAllPrices();
    try { await checkNegativePrices(cachedData); } catch (e) { console.error("  Alarm greška:", e.message); }

    // Catch-up: ako je server restartovan i propustio 22:00 UTC (23:00 CET) mail
    const now = new Date();
    const tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = getCET(tomorrow).datum;
    if (now.getUTCHours() >= 22 && dayAheadEmailSentDate !== tomorrowStr) {
      console.log(`[${now.toISOString()}] Catch-up: saljem propusteni day-ahead email...`);
      try { await sendDayAheadEmail(cachedData); } catch (e) { console.error("  Email greška:", e.message); }
    }
  });

  // Cron: svaki dan u 22:00 UTC (= 23:00 CET zimsko / 00:00 CEST ljetno)
  cron.schedule("0 22 * * *", async () => {
    console.log(`[${new Date().toISOString()}] Saljem day-ahead email...`);
    try {
      await sendDayAheadEmail(cachedData);
    } catch (e) {
      console.error("  Email greška:", e.message);
    }
  });

  console.log("Cron job: dohvaca podatke svakih sat u :00");
  console.log("Cron job: salje day-ahead email svaki dan u 22:00 UTC (23:00 CET)\n");
});
