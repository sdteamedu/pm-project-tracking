// ─── CONFIG ───────────────────────────────────────────────────────────────────
// กรอก Client ID จาก Azure AD App Registration
const MSAL_CONFIG = {
  auth: {
    clientId: "234635d5-0043-4b51-a99f-123dc2582323",   // ← เปลี่ยนตรงนี้
    authority: "https://login.microsoftonline.com/fa211b0f-864d-4b4f-ba33-62d3f613ce0c", // ← และตรงนี้
    redirectUri: "https://sdteamedu.github.io/pm-project-tracking",
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"];
const EXCEL_FILE_NAME = "PM_Dashboard_Data.xlsx";
const SHEET_NAME = "Projects";

// ─── SHAREPOINT CONFIG ────────────────────────────────────────────────────────
// เก็บไฟล์ใน SharePoint Team Site แทน OneDrive ส่วนตัว
// ทุกคนในทีมที่มีสิทธิ์เข้า Site นี้จะเห็นข้อมูลเดียวกัน
const SP_HOSTNAME    = "iwiredco.sharepoint.com";           // domain ของ SharePoint
const SP_SITE_PATH   = "/sites/Iwired-SolutionsDevelopment"; // path ของ Team Site
const SP_FOLDER_PATH = "Shared Documents/Products/PMTracking"; // folder ใน Site (Document Library)

// ─── STATE ────────────────────────────────────────────────────────────────────
let msalInstance = null;
let account = null;
let projects = [];
let editingId = null;
let excelFileId = null;
let currentPage = 'dashboard';

// ─── INIT ─────────────────────────────────────────────────────────────────────
async function initApp() {
  // ตรวจว่า MSAL โหลดแล้วหรือยัง
  if (typeof msal === 'undefined') {
    console.error('MSAL library not loaded');
    return;
  }
  try {
    msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    await msalInstance.initialize();

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      await bootApp();
    }
  } catch (e) {
    console.error('initApp error:', e);
  }
}

// รอให้ทุก script โหลดเสร็จก่อนค่อยรัน
window.addEventListener('load', initApp);

// ─── AUTH ─────────────────────────────────────────────────────────────────────
async function signIn() {
  try {
    const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
    account = result.account;
    await bootApp();
  } catch (e) {
    showToast('เข้าสู่ระบบไม่สำเร็จ: ' + e.message, 'error');
  }
}

async function signOut() {
  await msalInstance.logoutPopup();
  account = null;
  document.getElementById('loginScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
}

async function getToken() {
  const req = { scopes: GRAPH_SCOPES, account };
  try {
    const res = await msalInstance.acquireTokenSilent(req);
    return res.accessToken;
  } catch {
    const res = await msalInstance.acquireTokenPopup(req);
    return res.accessToken;
  }
}

// ─── BOOT ─────────────────────────────────────────────────────────────────────
async function bootApp() {
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('appShell').classList.remove('hidden');
  document.getElementById('userName').textContent = account.name || account.username;
  document.getElementById('userAvatar').textContent = (account.name || 'U')[0].toUpperCase();

  await ensureExcelFile();
  await loadProjects();
  showPage('dashboard');
}

// ─── GRAPH API HELPERS ────────────────────────────────────────────────────────
async function graphGet(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

async function graphPost(path, body) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

async function graphPatch(path, body) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

async function graphDelete(path) {
  const token = await getToken();
  await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${token}` }
  });
}

// ─── EXCEL FILE MANAGEMENT (SharePoint) ──────────────────────────────────────
const HEADERS = ['id','code','name','description','pm','client','startDate','endDate',
  'status','progress','budget','spent','risks','milestones','notes','updatedAt'];

// siteId cache
let _siteId = null;

// ดึง SharePoint Site ID จาก hostname + site path
async function getSiteId() {
  if (_siteId) return _siteId;
  const data = await graphGet(`/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`);
  _siteId = data.id;
  return _siteId;
}

// helper: workbook base path สำหรับ SharePoint
async function wbPath() {
  const siteId = await getSiteId();
  return `/sites/${siteId}/drive/items/${excelFileId}/workbook`;
}

// Graph API path prefix สำหรับ SharePoint site drive
async function siteDrivePrefix() {
  const siteId = await getSiteId();
  return `/sites/${siteId}/drive`;
}

// full path ของไฟล์ Excel ใน SharePoint
function spFilePath() {
  const folder = SP_FOLDER_PATH.replace(/^\/|\/$/g, '');
  return folder ? `${folder}/${EXCEL_FILE_NAME}` : EXCEL_FILE_NAME;
}

async function ensureExcelFile() {
  try {
    const prefix = await siteDrivePrefix();
    // ตรวจว่าไฟล์มีอยู่แล้วหรือยัง
    try {
      const file = await graphGet(`${prefix}/root:/${spFilePath()}`);
      excelFileId = file.id;
    } catch {
      // ยังไม่มีไฟล์ — สร้างใหม่ (folder Shared Documents มีอยู่แล้วใน SharePoint)
      await createExcelFile();
    }
  } catch (e) {
    console.error('ensureExcelFile error:', e);
    showToast('ไม่สามารถเชื่อมต่อ SharePoint', 'error');
  }
}

async function createExcelFile() {
  const token = await getToken();
  const siteId = await getSiteId();
  // อัพโหลดไฟล์เปล่าไปที่ SharePoint folder
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${spFilePath()}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      body: new ArrayBuffer(0)
    }
  );
  const file = await res.json();
  excelFileId = file.id;

  // เพิ่ม header row
  await graphPost(
    `/sites/${siteId}/drive/items/${excelFileId}/workbook/worksheets('Sheet1')/tables/add`,
    { address: 'Sheet1!A1:P1', hasHeaders: true }
  );
  await graphPatch(
    `/sites/${siteId}/drive/items/${excelFileId}/workbook/worksheets('Sheet1')/range(address='A1:P1')`,
    { values: [HEADERS] }
  );
}

// ─── LOAD PROJECTS ────────────────────────────────────────────────────────────
async function loadProjects() {
  if (!excelFileId) return;
  try {
    const data = await graphGet(
      `${await wbPath()}/worksheets('Sheet1')/usedRange`
    );
    const rows = data.values || [];
    if (rows.length <= 1) { projects = []; return; }

    const headers = rows[0];
    projects = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
      return obj;
    }).filter(p => p.id);
  } catch (e) {
    console.error('loadProjects error:', e);
    // Fallback: use localStorage for demo
    const saved = localStorage.getItem('pm_projects_local');
    if (saved) projects = JSON.parse(saved);
  }
}

// ─── SAVE PROJECT ─────────────────────────────────────────────────────────────
async function saveProjectToExcel(project) {
  const row = HEADERS.map(h => project[h] ?? '');
  if (editingId) {
    // Find row index in sheet (row 2+ = data, row 1 = header)
    const idx = projects.findIndex(p => p.id === editingId);
    const rowNum = idx + 2; // 1-indexed, +1 for header
    try {
      await graphPatch(
        `${await wbPath()}/worksheets('Sheet1')/range(address='A${rowNum}:P${rowNum}')`,
        { values: [row] }
      );
    } catch (e) {
      console.error(e);
      saveLocal();
    }
  } else {
    // Append new row
    try {
      await graphPost(
        `${await wbPath()}/worksheets('Sheet1')/tables/0/rows/add`,
        { values: [row] }
      );
    } catch (e) {
      // Fallback: find next empty row
      const nextRow = projects.length + 2;
      try {
        await graphPatch(
          `${await wbPath()}/worksheets('Sheet1')/range(address='A${nextRow}:P${nextRow}')`,
          { values: [row] }
        );
      } catch (e2) {
        console.error(e2);
        saveLocal();
      }
    }
  }
}

function saveLocal() {
  localStorage.setItem('pm_projects_local', JSON.stringify(projects));
}

// ─── SYNC ─────────────────────────────────────────────────────────────────────
async function syncData() {
  showToast('กำลังซิงค์ข้อมูล...', '');
  await loadProjects();
  renderCurrentPage();
  showToast('ซิงค์เรียบร้อย ✓', 'success');
}

// ─── PAGES ────────────────────────────────────────────────────────────────────
function showPage(page) {
  currentPage = page;
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const navMap = { dashboard: 0, projects: 1, financials: 3, report: 4 };
  const navItems = document.querySelectorAll('.nav-item');
  if (navMap[page] !== undefined) navItems[navMap[page]]?.classList.add('active');
  renderCurrentPage();
}

function renderCurrentPage() {
  const titles = { dashboard: 'ภาพรวม', projects: 'โครงการทั้งหมด', report: 'สร้างรายงาน / สไลด์', financials: 'Financial & Resources' };
  document.getElementById('pageTitle').textContent = titles[currentPage] || '';
  switch (currentPage) {
    case 'dashboard':   renderDashboard();   break;
    case 'projects':    renderProjects();    break;
    case 'detail':      renderDetail();      break;
    case 'report':      renderReport();      break;
    case 'financials':  renderFinancials();  break;
  }
}

// ─── DASHBOARD (BI Style) ─────────────────────────────────────────────────────
function renderDashboard() {
  const total     = projects.length;
  const onTrack   = projects.filter(p => p.status === 'on-track').length;
  const atRisk    = projects.filter(p => p.status === 'at-risk').length;
  const delayed   = projects.filter(p => p.status === 'delayed').length;
  const completed = projects.filter(p => p.status === 'completed').length;
  const planning  = projects.filter(p => p.status === 'planning').length;

  const totalBudget  = projects.reduce((a, p) => a + (parseInt(p.budget) || 0), 0);
  const totalSpent   = projects.reduce((a, p) => a + (parseInt(p.spent)  || 0), 0);
  const budgetPct    = totalBudget ? Math.round(totalSpent / totalBudget * 100) : 0;
  const avgProgress  = total ? Math.round(projects.reduce((a, p) => a + (parseInt(p.progress)||0), 0) / total) : 0;

  const recent     = [...projects].sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt)).slice(0, 5);
  const needsAttn  = projects.filter(p => p.status === 'at-risk' || p.status === 'delayed');
  const duesSoon   = [...projects].filter(p => p.endDate && p.status !== 'completed')
                       .sort((a, b) => new Date(a.endDate) - new Date(b.endDate)).slice(0, 4);

  // ── Donut chart SVG ──
  const statuses  = ['on-track','at-risk','delayed','completed','planning'];
  const colors    = ['#16A34A','#D97706','#DC2626','#2563EB','#9CA3AF'];
  const counts    = statuses.map(s => projects.filter(p => p.status === s).length);
  const circ      = 2 * Math.PI * 38;
  let offset = 0;
  const arcs = counts.map((c, i) => {
    if (!c || !total) return '';
    const len = circ * (c / total);
    const arc = `<circle cx="44" cy="44" r="38" fill="none" stroke="${colors[i]}" stroke-width="11"
      stroke-dasharray="${len} ${circ - len}" stroke-dashoffset="${-offset}"
      transform="rotate(-90 44 44)" stroke-linecap="butt"/>`;
    offset += len;
    return arc;
  }).join('');

  // ── Budget bar chart ──
  const topBudget = [...projects].filter(p => p.budget).sort((a, b) => (b.budget||0)-(a.budget||0)).slice(0,5);
  const maxBgt    = Math.max(...topBudget.map(p => parseInt(p.budget)||0), 1);

  // ── Progress heatmap data ──
  const pmGroups = {};
  projects.forEach(p => {
    if (!p.pm) return;
    if (!pmGroups[p.pm]) pmGroups[p.pm] = [];
    pmGroups[p.pm].push(p);
  });

  // ── Days until deadline ──
  const today = new Date();
  const daysDiff = p => {
    if (!p.endDate) return null;
    return Math.ceil((new Date(p.endDate) - today) / (1000 * 60 * 60 * 24));
  };

  document.getElementById('pageContent').innerHTML = `
  <style>
    .bi-grid { display: grid; gap: 16px; }
    .bi-row2 { grid-template-columns: 1fr 1fr; }
    .bi-row3 { grid-template-columns: 1fr 1fr 1fr; }
    .bi-row-main { grid-template-columns: 2fr 1fr; }
    .bi-row-side { grid-template-columns: 1fr 1fr 1fr 1fr; }
    .bi-card {
      background: var(--card); border: 1px solid var(--border);
      border-radius: var(--radius); padding: 18px;
      box-shadow: var(--shadow);
    }
    .bi-card-title {
      font-size: 12px; font-weight: 700; color: var(--text-dim);
      text-transform: uppercase; letter-spacing: 1px; margin-bottom: 14px;
      display: flex; align-items: center; justify-content: space-between;
    }
    .kpi-val { font-size: 36px; font-weight: 700; line-height: 1; margin: 4px 0; }
    .kpi-sub { font-size: 12px; color: var(--text-dim); margin-top: 4px; }
    .kpi-chip {
      display: inline-flex; align-items: center; gap: 3px;
      padding: 2px 8px; border-radius: 20px; font-size: 11px; font-weight: 600;
    }
    .chip-green { background: #DCFCE7; color: #16A34A; }
    .chip-red { background: #FEE2E2; color: #DC2626; }
    .chip-amber { background: #FEF3C7; color: #D97706; }
    .chip-blue { background: #DBEAFE; color: #2563EB; }

    .gauge-wrap { display: flex; align-items: center; gap: 6px; margin-top: 10px; }
    .gauge-bar { flex: 1; height: 8px; background: var(--border); border-radius: 4px; overflow: hidden; }
    .gauge-fill { height: 100%; border-radius: 4px; transition: width .4s; }

    .donut-legend { display: flex; flex-direction: column; gap: 7px; }
    .dl-row { display: flex; align-items: center; gap: 8px; font-size: 12px; }
    .dl-dot { width: 8px; height: 8px; border-radius: 2px; flex-shrink: 0; }
    .dl-label { flex: 1; color: var(--text-dim); }
    .dl-count { font-weight: 700; color: var(--text); font-family: 'IBM Plex Mono', monospace; }
    .dl-pct { font-size: 10px; color: var(--text-light); }

    .budget-row { display: flex; align-items: center; gap: 10px; margin-bottom: 11px; }
    .budget-row:last-child { margin-bottom: 0; }
    .budget-name { width: 110px; font-size: 12px; font-weight: 500; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .budget-bar-wrap { flex: 1; display: flex; flex-direction: column; gap: 3px; }
    .budget-bar-track { height: 7px; background: var(--border); border-radius: 4px; overflow: hidden; }
    .budget-bar-fill { height: 100%; border-radius: 4px; transition: width .5s; }
    .budget-nums { font-size: 10px; color: var(--text-dim); font-family: 'IBM Plex Mono', monospace; text-align: right; white-space: nowrap; }

    .attn-item { display: flex; align-items: flex-start; gap: 10px; padding: 10px 0; border-bottom: 1px solid var(--border); cursor: pointer; transition: background .12s; }
    .attn-item:last-child { border-bottom: none; }
    .attn-item:hover { background: var(--primary-light); margin: 0 -18px; padding: 10px 18px; border-radius: 6px; }
    .attn-dot { width: 8px; height: 8px; border-radius: 50%; margin-top: 5px; flex-shrink: 0; }
    .attn-name { font-size: 13px; font-weight: 600; color: var(--text); margin-bottom: 2px; }
    .attn-desc { font-size: 11px; color: var(--text-dim); line-height: 1.4; }

    .due-item { display: flex; align-items: center; gap: 10px; padding: 9px 0; border-bottom: 1px solid var(--border); cursor: pointer; }
    .due-item:last-child { border-bottom: none; }
    .due-item:hover { background: var(--primary-light); margin: 0 -18px; padding: 9px 18px; border-radius: 6px; }
    .due-days { width: 44px; text-align: center; flex-shrink: 0; }
    .due-days-num { font-size: 18px; font-weight: 700; line-height: 1; }
    .due-days-label { font-size: 9px; color: var(--text-dim); }
    .due-name { flex: 1; font-size: 12.5px; font-weight: 500; }
    .due-code { font-size: 10px; color: var(--text-dim); font-family: 'IBM Plex Mono', monospace; }

    .pm-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
    .pm-row:last-child { margin-bottom: 0; }
    .pm-av { width: 28px; height: 28px; border-radius: 50%; background: var(--primary); color: white; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; flex-shrink: 0; }
    .pm-name { font-size: 12px; font-weight: 600; flex-shrink: 0; width: 90px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .pm-bars { flex: 1; display: flex; gap: 3px; }
    .pm-bar { height: 20px; border-radius: 3px; display: flex; align-items: center; justify-content: center; font-size: 9px; font-weight: 700; color: white; min-width: 20px; transition: width .4s; }
    .pm-total { font-size: 11px; color: var(--text-dim); font-family: 'IBM Plex Mono', monospace; width: 30px; text-align: right; }

    .prog-list-item { display: flex; align-items: center; gap: 10px; padding: 8px 0; border-bottom: 1px solid var(--border); cursor: pointer; }
    .prog-list-item:last-child { border-bottom: none; }
    .prog-list-item:hover { background: var(--primary-light); margin: 0 -18px; padding: 8px 18px; border-radius: 6px; }
    .prog-list-name { flex: 1; font-size: 12.5px; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .prog-list-bar { width: 120px; height: 6px; background: var(--border); border-radius: 3px; overflow: hidden; flex-shrink: 0; }
    .prog-list-fill { height: 100%; border-radius: 3px; }
    .prog-list-pct { width: 34px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'IBM Plex Mono', monospace; flex-shrink: 0; }
  </style>

  <!-- ROW 1: KPI Cards -->
  <div class="bi-grid bi-row-side" style="margin-bottom:16px">
    <div class="bi-card">
      <div class="bi-card-title">โครงการทั้งหมด</div>
      <div class="kpi-val" style="color:var(--primary)">${total}</div>
      <div class="kpi-sub">${completed} เสร็จแล้ว · ${total - completed} กำลังดำเนินการ</div>
      <div class="gauge-wrap"><div class="gauge-bar"><div class="gauge-fill" style="width:${total?Math.round(completed/total*100):0}%;background:var(--primary)"></div></div><span style="font-size:11px;color:var(--text-dim)">${total?Math.round(completed/total*100):0}%</span></div>
    </div>
    <div class="bi-card">
      <div class="bi-card-title">ความคืบหน้าเฉลี่ย</div>
      <div class="kpi-val" style="color:${avgProgress>=70?'#16A34A':avgProgress>=40?'#D97706':'#DC2626'}">${avgProgress}%</div>
      <div class="kpi-sub">ทุกโครงการที่ Active</div>
      <div class="gauge-wrap"><div class="gauge-bar"><div class="gauge-fill" style="width:${avgProgress}%;background:${avgProgress>=70?'#16A34A':avgProgress>=40?'#D97706':'#DC2626'}"></div></div></div>
    </div>
    <div class="bi-card">
      <div class="bi-card-title">งบประมาณรวม</div>
      <div class="kpi-val" style="color:var(--text);font-size:26px;margin-top:6px">${(totalBudget/1000000).toFixed(1)}M</div>
      <div class="kpi-sub">ใช้ไป ${(totalSpent/1000000).toFixed(1)}M ฿ (${budgetPct}%)</div>
      <div class="gauge-wrap"><div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(budgetPct,100)}%;background:${budgetPct>90?'#DC2626':budgetPct>70?'#D97706':'#16A34A'}"></div></div><span style="font-size:11px;color:var(--text-dim)">${budgetPct}%</span></div>
    </div>
    <div class="bi-card">
      <div class="bi-card-title">Health Status</div>
      <div style="display:flex;gap:8px;margin-top:4px;flex-wrap:wrap">
        <span class="kpi-chip chip-green">✓ ${onTrack} On Track</span>
        <span class="kpi-chip chip-amber">⚠ ${atRisk} At Risk</span>
        <span class="kpi-chip chip-red">✕ ${delayed} Delayed</span>
        <span class="kpi-chip chip-blue">○ ${planning} Planning</span>
      </div>
      <div style="margin-top:12px;display:flex;gap:2px;height:8px;border-radius:4px;overflow:hidden">
        ${total ? [
          [onTrack,'#16A34A'],[atRisk,'#D97706'],[delayed,'#DC2626'],[planning,'#9CA3AF'],[completed,'#2563EB']
        ].map(([c,col]) => c ? `<div style="flex:${c};background:${col}"></div>` : '').join('') : '<div style="flex:1;background:var(--border)"></div>'}
      </div>
    </div>
  </div>

  <!-- ROW 2: Donut + Budget + Needs Attention -->
  <div class="bi-grid bi-row3" style="margin-bottom:16px">
    <!-- Donut -->
    <div class="bi-card">
      <div class="bi-card-title">สัดส่วนสถานะ <span style="font-weight:400;text-transform:none;letter-spacing:0">${total} โครงการ</span></div>
      <div style="display:flex;align-items:center;gap:18px">
        <svg width="88" height="88" viewBox="0 0 88 88" style="flex-shrink:0">
          <circle cx="44" cy="44" r="38" fill="none" stroke="#F3F4F6" stroke-width="11"/>
          ${total ? arcs : ''}
          <text x="44" y="40" text-anchor="middle" fill="#1F2937" font-size="18" font-weight="700" font-family="sans-serif">${total}</text>
          <text x="44" y="54" text-anchor="middle" fill="#9CA3AF" font-size="9" font-family="sans-serif">projects</text>
        </svg>
        <div class="donut-legend">
          ${[['on-track','On Track','#16A34A',onTrack],['at-risk','At Risk','#D97706',atRisk],
             ['delayed','Delayed','#DC2626',delayed],['completed','Done','#2563EB',completed],
             ['planning','Planning','#9CA3AF',planning]].map(([,label,color,count]) => count > 0 ? `
          <div class="dl-row">
            <div class="dl-dot" style="background:${color}"></div>
            <span class="dl-label">${label}</span>
            <span class="dl-count">${count}</span>
            <span class="dl-pct">${total?Math.round(count/total*100):0}%</span>
          </div>` : '').join('')}
        </div>
      </div>
    </div>

    <!-- Budget Breakdown -->
    <div class="bi-card">
      <div class="bi-card-title">งบประมาณ — Top 5 โครงการ</div>
      ${topBudget.length === 0
        ? `<div class="empty" style="padding:20px"><div class="empty-icon">💰</div><div class="empty-text" style="font-size:12px">ยังไม่มีข้อมูล</div></div>`
        : topBudget.map(p => {
            const bgt = parseInt(p.budget)||0;
            const spt = parseInt(p.spent)||0;
            const bw  = Math.round(bgt/maxBgt*100);
            const sw  = bgt ? Math.round(spt/bgt*100) : 0;
            const bc  = sw > 90 ? '#DC2626' : sw > 70 ? '#D97706' : '#CC1E2B';
            return `<div class="budget-row">
              <div class="budget-name" title="${p.name}">${p.code}</div>
              <div class="budget-bar-wrap">
                <div class="budget-bar-track"><div class="budget-bar-fill" style="width:${bw}%;background:#E5E7EB"></div></div>
                <div class="budget-bar-track"><div class="budget-bar-fill" style="width:${Math.min(sw,100)}%;background:${bc}"></div></div>
              </div>
              <div class="budget-nums">${(bgt/1000).toFixed(0)}K</div>
            </div>`;
          }).join('')
      }
      ${topBudget.length > 0 ? `<div style="display:flex;gap:12px;margin-top:10px;font-size:10px;color:var(--text-dim)"><span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:4px;background:#E5E7EB;border-radius:2px;display:inline-block"></span>งบประมาณ</span><span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:4px;background:var(--primary);border-radius:2px;display:inline-block"></span>ใช้จริง</span></div>` : ''}
    </div>

    <!-- Needs Attention -->
    <div class="bi-card">
      <div class="bi-card-title" style="color:${needsAttn.length?'#DC2626':'inherit'}">
        ${needsAttn.length ? '🚨 ต้องดำเนินการด่วน' : '⚡ สถานะโครงการ'}
        <span style="background:${needsAttn.length?'#FEE2E2':'#DCFCE7'};color:${needsAttn.length?'#DC2626':'#16A34A'};padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;text-transform:none;letter-spacing:0">${needsAttn.length} รายการ</span>
      </div>
      ${needsAttn.length === 0
        ? `<div class="empty" style="padding:20px"><div class="empty-icon">✅</div><div class="empty-text" style="font-size:12px">ทุกโครงการเป็นไปตามแผน</div></div>`
        : needsAttn.slice(0,4).map(p => `
          <div class="attn-item" onclick="viewProject('${p.id}')">
            <div class="attn-dot" style="background:${p.status==='delayed'?'#DC2626':'#D97706'}"></div>
            <div style="flex:1">
              <div class="attn-name">${p.name}</div>
              <div class="attn-desc">${(p.risks||'ไม่มีหมายเหตุ').slice(0,55)}${(p.risks||'').length>55?'...':''}</div>
              <div style="margin-top:4px">${statusBadge(p.status)}</div>
            </div>
          </div>`).join('')
      }
    </div>
  </div>

  <!-- ROW 3: Progress List + Due Soon + PM Workload -->
  <div class="bi-grid" style="grid-template-columns:1.2fr 1fr 1fr;margin-bottom:16px">
    <!-- Progress per project -->
    <div class="bi-card">
      <div class="bi-card-title">ความคืบหน้าแต่ละโครงการ <button class="btn btn-primary" onclick="openAddProject()" style="padding:3px 10px;font-size:11px;text-transform:none;letter-spacing:0;font-weight:600">+ เพิ่ม</button></div>
      ${projects.length === 0
        ? `<div class="empty" style="padding:20px"><div class="empty-icon">📁</div><div class="empty-text" style="font-size:12px">ยังไม่มีโครงการ</div></div>`
        : [...projects].sort((a,b)=>(parseInt(b.progress)||0)-(parseInt(a.progress)||0)).map(p => {
            const pct = parseInt(p.progress)||0;
            const col = pct>=70?'#16A34A':pct>=40?'#D97706':'#DC2626';
            return `<div class="prog-list-item" onclick="viewProject('${p.id}')">
              <div style="flex:1;min-width:0">
                <div class="prog-list-name">${p.name}</div>
                <div style="font-size:10px;color:var(--text-dim);font-family:'IBM Plex Mono',monospace">${p.code}</div>
              </div>
              <div class="prog-list-bar"><div class="prog-list-fill" style="width:${pct}%;background:${col}"></div></div>
              <div class="prog-list-pct" style="color:${col}">${pct}%</div>
            </div>`;
          }).join('')
      }
    </div>

    <!-- Due Soon -->
    <div class="bi-card">
      <div class="bi-card-title">กำหนดส่งใกล้ถึง</div>
      ${duesSoon.length === 0
        ? `<div class="empty" style="padding:20px"><div class="empty-icon">📅</div><div class="empty-text" style="font-size:12px">ไม่มีกำหนดที่ใกล้ถึง</div></div>`
        : duesSoon.map(p => {
            const d = daysDiff(p);
            const dc = d <= 7 ? '#DC2626' : d <= 30 ? '#D97706' : '#16A34A';
            return `<div class="due-item" onclick="viewProject('${p.id}')">
              <div class="due-days">
                <div class="due-days-num" style="color:${dc}">${d !== null ? (d < 0 ? `+${Math.abs(d)}` : d) : '—'}</div>
                <div class="due-days-label">${d < 0 ? 'เกินกำหนด' : 'วัน'}</div>
              </div>
              <div style="flex:1;min-width:0">
                <div class="due-code">${p.code}</div>
                <div class="due-name">${p.name.length>26?p.name.slice(0,26)+'…':p.name}</div>
                <div style="margin-top:3px">${statusBadge(p.status)}</div>
              </div>
            </div>`;
          }).join('')
      }
    </div>

    <!-- PM Workload -->
    <div class="bi-card">
      <div class="bi-card-title">PM Workload</div>
      ${Object.keys(pmGroups).length === 0
        ? `<div class="empty" style="padding:20px"><div class="empty-icon">👤</div><div class="empty-text" style="font-size:12px">ยังไม่มีข้อมูล PM</div></div>`
        : Object.entries(pmGroups).map(([pm, projs]) => {
            const active   = projs.filter(p => p.status !== 'completed').length;
            const done     = projs.filter(p => p.status === 'completed').length;
            const atRiskPM = projs.filter(p => p.status === 'at-risk' || p.status === 'delayed').length;
            const initials = pm.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
            return `<div class="pm-row">
              <div class="pm-av">${initials}</div>
              <div class="pm-name" title="${pm}">${pm.split(' ')[0]}</div>
              <div class="pm-bars">
                ${active ? `<div class="pm-bar" style="flex:${active};background:#CC1E2B" title="Active: ${active}">${active}</div>` : ''}
                ${done  ? `<div class="pm-bar" style="flex:${done};background:#16A34A" title="Done: ${done}">${done}</div>` : ''}
              </div>
              <div class="pm-total">${projs.length}p</div>
            </div>`;
          }).join('')
      }
      ${Object.keys(pmGroups).length > 0 ? `<div style="display:flex;gap:10px;margin-top:10px;font-size:10px;color:var(--text-dim)"><span style="display:flex;align-items:center;gap:4px"><span style="width:8px;height:8px;background:var(--primary);border-radius:2px;display:inline-block"></span>Active</span><span style="display:flex;align-items:center;gap:4px"><span style="width:8px;height:8px;background:#16A34A;border-radius:2px;display:inline-block"></span>Done</span></div>` : ''}
    </div>
  </div>

  <!-- ROW 4: Recent Table -->
  <div class="bi-card" style="padding:0">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--border)">
      <span class="bi-card-title" style="margin:0">อัปเดตล่าสุด</span>
      <button class="btn btn-secondary" onclick="showPage('projects')" style="font-size:12px;padding:5px 12px">ดูทั้งหมด →</button>
    </div>
    ${recent.length === 0
      ? `<div class="empty"><div class="empty-icon">📁</div><div class="empty-text">ยังไม่มีโครงการ กด "+ เพิ่มโครงการ" เพื่อเริ่มต้น</div></div>`
      : `<table class="project-table">
          <thead><tr><th>รหัส</th><th>ชื่อโครงการ</th><th>PM</th><th>สถานะ</th><th>ความคืบหน้า</th><th>วันสิ้นสุด</th></tr></thead>
          <tbody>${recent.map(p => projectRow(p)).join('')}</tbody>
        </table>`
    }
  </div>`;
}

// ─── PROJECTS LIST ────────────────────────────────────────────────────────────
function renderProjects() {
  document.getElementById('pageContent').innerHTML = `
    <div class="section-header">
      <div class="section-title">โครงการทั้งหมด (${projects.length})</div>
      <button class="btn btn-primary" onclick="openAddProject()">+ เพิ่มโครงการ</button>
    </div>
    <div class="table-card">
      ${projects.length === 0 ? `<div class="empty"><div class="empty-icon">📁</div><div class="empty-text">ยังไม่มีโครงการ</div></div>` :
        `<table class="project-table">
          <thead><tr>
            <th>รหัส</th><th>ชื่อโครงการ</th><th>PM</th><th>ลูกค้า</th><th>สถานะ</th><th>ความคืบหน้า</th><th>งบประมาณ</th><th>วันสิ้นสุด</th><th></th>
          </tr></thead>
          <tbody>
            ${projects.map(p => projectRowFull(p)).join('')}
          </tbody>
        </table>`
      }
    </div>
  `;
}

function projectRow(p) {
  const pct = parseInt(p.progress) || 0;
  const cls = pct >= 70 ? '' : pct >= 40 ? 'gold' : 'red';
  return `<tr onclick="viewProject('${p.id}')" style="cursor:pointer;">
    <td><span style="font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--text-dim)">${p.code}</span></td>
    <td style="font-weight:500">${p.name}</td>
    <td>${p.pm}</td>
    <td>${statusBadge(p.status)}</td>
    <td><div class="progress-wrap"><div class="progress-bar"><div class="progress-fill ${cls}" style="width:${pct}%"></div></div><span class="progress-pct">${pct}%</span></div></td>
    <td style="font-family:'IBM Plex Mono',monospace;font-size:12px">${p.endDate || '-'}</td>
  </tr>`;
}

function projectRowFull(p) {
  const pct = parseInt(p.progress) || 0;
  const budget = p.budget ? parseInt(p.budget).toLocaleString() : '-';
  const cls = pct >= 70 ? '' : pct >= 40 ? 'gold' : 'red';
  return `<tr>
    <td style="cursor:pointer;font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--text-dim)" onclick="viewProject('${p.id}')">${p.code}</td>
    <td style="cursor:pointer;font-weight:500" onclick="viewProject('${p.id}')">${p.name}</td>
    <td>${p.pm}</td>
    <td>${p.client || '-'}</td>
    <td>${statusBadge(p.status)}</td>
    <td><div class="progress-wrap"><div class="progress-bar"><div class="progress-fill ${cls}" style="width:${pct}%"></div></div><span class="progress-pct">${pct}%</span></div></td>
    <td style="font-family:'IBM Plex Mono',monospace;font-size:12px">${budget}</td>
    <td style="font-family:'IBM Plex Mono',monospace;font-size:12px">${p.endDate || '-'}</td>
    <td>
      <div class="flex gap-2">
        <button class="btn btn-secondary" onclick="openEditProject('${p.id}')" style="padding:4px 10px;font-size:12px">แก้ไข</button>
        <button class="btn btn-danger" onclick="deleteProject('${p.id}')" style="padding:4px 10px;font-size:12px">ลบ</button>
      </div>
    </td>
  </tr>`;
}

function statusBadge(status) {
  const map = {
    'on-track': ['on-track', 'On Track'],
    'at-risk': ['at-risk', 'At Risk'],
    'delayed': ['delayed', 'Delayed'],
    'completed': ['completed', 'Completed'],
    'planning': ['planning', 'Planning'],
  };
  const [cls, label] = map[status] || ['planning', status];
  return `<span class="badge ${cls}"><span class="badge-dot"></span>${label}</span>`;
}

// ─── PROJECT DETAIL ───────────────────────────────────────────────────────────
let viewingId = null;
function viewProject(id) {
  viewingId = id;
  currentPage = 'detail';
  renderDetail();
}

function renderDetail() {
  const p = projects.find(x => x.id === viewingId);
  if (!p) { showPage('projects'); return; }
  document.getElementById('pageTitle').textContent = p.name;

  const pct = parseInt(p.progress) || 0;
  const budget = p.budget ? parseInt(p.budget).toLocaleString() : '-';
  const spent = p.spent ? parseInt(p.spent).toLocaleString() : '-';
  const budgetPct = p.budget && p.spent ? Math.round((parseInt(p.spent) / parseInt(p.budget)) * 100) : 0;

  let milestonesHtml = '';
  if (p.milestones) {
    const ms = p.milestones.split(';').map(s => s.trim()).filter(Boolean);
    milestonesHtml = `<div class="milestone-list">${ms.map(m => `
      <div class="milestone-item">
        <div class="milestone-dot"></div>
        <div class="milestone-name">${m}</div>
      </div>`).join('')}</div>`;
  }

  document.getElementById('pageContent').innerHTML = `
    <div class="detail-header">
      <div>
        <div class="detail-title">${p.name}</div>
        <div class="detail-code">${p.code}</div>
      </div>
      <div class="flex gap-2">
        <button class="btn btn-secondary" onclick="showPage('projects')">← กลับ</button>
        <button class="btn btn-secondary" onclick="openEditProject('${p.id}')">แก้ไข</button>
        <button class="btn btn-primary" onclick="generateSingleProjectSlide('${p.id}')">📊 สร้างสไลด์</button>
      </div>
    </div>

    <div class="detail-grid">
      <div class="detail-item">
        <div class="detail-item-label">สถานะ</div>
        <div>${statusBadge(p.status)}</div>
      </div>
      <div class="detail-item">
        <div class="detail-item-label">PM รับผิดชอบ</div>
        <div class="detail-item-value">${p.pm || '-'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-item-label">ลูกค้า</div>
        <div class="detail-item-value">${p.client || '-'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-item-label">วันเริ่ม – สิ้นสุด</div>
        <div class="detail-item-value" style="font-size:13px">${p.startDate || '-'} → ${p.endDate || '-'}</div>
      </div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px">
      <div class="table-card" style="padding:20px">
        <div class="detail-item-label" style="margin-bottom:12px">ความคืบหน้า</div>
        <div style="font-size:36px;font-weight:700;color:var(--teal);margin-bottom:10px">${pct}%</div>
        <div class="progress-bar" style="height:10px"><div class="progress-fill" style="width:${pct}%"></div></div>
      </div>
      <div class="table-card" style="padding:20px">
        <div class="detail-item-label" style="margin-bottom:12px">งบประมาณ</div>
        <div style="font-size:22px;font-weight:700;margin-bottom:4px">${budget} บาท</div>
        <div style="font-size:13px;color:var(--text-dim);margin-bottom:10px">ใช้ไปแล้ว: ${spent} บาท (${budgetPct}%)</div>
        <div class="progress-bar" style="height:10px"><div class="progress-fill ${budgetPct > 90 ? 'red' : budgetPct > 70 ? 'gold' : ''}" style="width:${Math.min(budgetPct,100)}%"></div></div>
      </div>
    </div>

    ${p.description ? `<div class="table-card" style="padding:20px;margin-bottom:16px">
      <div class="detail-item-label" style="margin-bottom:8px">รายละเอียดโครงการ</div>
      <div style="font-size:14px;line-height:1.7;color:var(--text)">${p.description}</div>
    </div>` : ''}

    ${p.risks ? `<div class="table-card" style="padding:20px;margin-bottom:16px;border-color:rgba(239,68,68,0.2)">
      <div class="detail-item-label" style="margin-bottom:8px;color:var(--red)">⚠️ ความเสี่ยง</div>
      <div style="font-size:14px;line-height:1.7">${p.risks}</div>
    </div>` : ''}

    ${p.notes ? `<div class="table-card" style="padding:20px;margin-bottom:16px">
      <div class="detail-item-label" style="margin-bottom:8px">📝 Update ล่าสุด</div>
      <div style="font-size:14px;line-height:1.7">${p.notes}</div>
    </div>` : ''}

    ${milestonesHtml ? `<div>
      <div class="section-title" style="margin-bottom:12px">🏁 Milestones</div>
      ${milestonesHtml}
    </div>` : ''}
  `;
}

// ─── REPORT PAGE ──────────────────────────────────────────────────────────────
let selectedReportType = 'weekly';
let selectedProjects = new Set();

function renderReport() {
  const projectCheckboxes = projects.map(p => `
    <label style="display:flex;align-items:center;gap:10px;padding:10px;background:var(--card);border:1px solid var(--border);border-radius:8px;cursor:pointer;font-size:14px">
      <input type="checkbox" id="chk_${p.id}" onchange="toggleSelectProject('${p.id}')" style="accent-color:var(--teal);width:16px;height:16px;">
      <span style="flex:1">${p.code} — <strong>${p.name}</strong></span>
      ${statusBadge(p.status)}
    </label>
  `).join('');

  document.getElementById('pageContent').innerHTML = `
    <div style="max-width:800px">
      <p style="color:var(--text-dim);margin-bottom:24px;font-size:14px">เลือกประเภทรายงาน จากนั้นเลือกโครงการที่ต้องการ แล้วกด "สร้าง PowerPoint"</p>

      <div class="section-title" style="margin-bottom:12px">ประเภทรายงาน</div>
      <div class="report-options">
        <div class="report-card selected" id="rt_weekly" onclick="selectReportType('weekly')">
          <div class="report-icon">📅</div>
          <div class="report-name">Weekly Report</div>
          <div class="report-desc">สรุปความคืบหน้าประจำสัปดาห์ แสดงสถานะและปัญหาของแต่ละโครงการ</div>
        </div>
        <div class="report-card" id="rt_monthly" onclick="selectReportType('monthly')">
          <div class="report-icon">📊</div>
          <div class="report-name">Monthly Report</div>
          <div class="report-desc">รายงานประจำเดือน ภาพรวมพอร์ตโฟลิโอ และงบประมาณ</div>
        </div>
        <div class="report-card" id="rt_executive" onclick="selectReportType('executive')">
          <div class="report-icon">👔</div>
          <div class="report-name">Executive Summary</div>
          <div class="report-desc">สรุปสำหรับผู้บริหาร เน้น KPI และ highlight สำคัญ</div>
        </div>
      </div>

      <div class="section-header" style="margin-top:24px">
        <div class="section-title">เลือกโครงการ</div>
        <div class="flex gap-2">
          <button class="btn btn-secondary" onclick="selectAllProjects()" style="font-size:12px">เลือกทั้งหมด</button>
          <button class="btn btn-secondary" onclick="clearSelectProjects()" style="font-size:12px">ล้าง</button>
        </div>
      </div>
      ${projects.length === 0 
        ? `<div class="empty" style="padding:32px"><div class="empty-icon">📁</div><div class="empty-text">ยังไม่มีโครงการ</div></div>`
        : `<div style="display:flex;flex-direction:column;gap:8px;margin-bottom:24px">${projectCheckboxes}</div>`
      }

      <button class="btn btn-primary" onclick="generatePPTX()" style="padding:14px 28px;font-size:15px;width:100%;justify-content:center">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        สร้าง PowerPoint (.pptx)
      </button>
    </div>
  `;
}

function selectReportType(type) {
  selectedReportType = type;
  document.querySelectorAll('.report-card').forEach(el => el.classList.remove('selected'));
  document.getElementById('rt_' + type)?.classList.add('selected');
}

function toggleSelectProject(id) {
  if (selectedProjects.has(id)) selectedProjects.delete(id);
  else selectedProjects.add(id);
}

function selectAllProjects() {
  projects.forEach(p => {
    selectedProjects.add(p.id);
    const chk = document.getElementById('chk_' + p.id);
    if (chk) chk.checked = true;
  });
}

function clearSelectProjects() {
  selectedProjects.clear();
  projects.forEach(p => {
    const chk = document.getElementById('chk_' + p.id);
    if (chk) chk.checked = false;
  });
}

// ─── CRUD ─────────────────────────────────────────────────────────────────────
function openAddProject() {
  editingId = null;
  document.getElementById('modalTitle').textContent = 'เพิ่มโครงการใหม่';
  ['code','name','desc','pm','client','start','end','risks','milestones','notes'].forEach(f => {
    document.getElementById('f_' + f).value = '';
  });
  document.getElementById('f_status').value = 'planning';
  document.getElementById('f_progress').value = '';
  document.getElementById('f_budget').value = '';
  document.getElementById('f_spent').value = '';
  openModal();
}

function openEditProject(id) {
  editingId = id;
  const p = projects.find(x => x.id === id);
  if (!p) return;
  document.getElementById('modalTitle').textContent = 'แก้ไขโครงการ';
  document.getElementById('f_code').value = p.code || '';
  document.getElementById('f_name').value = p.name || '';
  document.getElementById('f_desc').value = p.description || '';
  document.getElementById('f_pm').value = p.pm || '';
  document.getElementById('f_client').value = p.client || '';
  document.getElementById('f_start').value = p.startDate || '';
  document.getElementById('f_end').value = p.endDate || '';
  document.getElementById('f_status').value = p.status || 'planning';
  document.getElementById('f_progress').value = p.progress || '';
  document.getElementById('f_budget').value = p.budget || '';
  document.getElementById('f_spent').value = p.spent || '';
  document.getElementById('f_risks').value = p.risks || '';
  document.getElementById('f_milestones').value = p.milestones || '';
  document.getElementById('f_notes').value = p.notes || '';
  openModal();
}

async function saveProject() {
  const name = document.getElementById('f_name').value.trim();
  if (!name) { showToast('กรุณากรอกชื่อโครงการ', 'error'); return; }

  const project = {
    id: editingId || 'prj_' + Date.now(),
    code: document.getElementById('f_code').value.trim(),
    name,
    description: document.getElementById('f_desc').value.trim(),
    pm: document.getElementById('f_pm').value.trim(),
    client: document.getElementById('f_client').value.trim(),
    startDate: document.getElementById('f_start').value,
    endDate: document.getElementById('f_end').value,
    status: document.getElementById('f_status').value,
    progress: document.getElementById('f_progress').value,
    budget: document.getElementById('f_budget').value,
    spent: document.getElementById('f_spent').value,
    risks: document.getElementById('f_risks').value.trim(),
    milestones: document.getElementById('f_milestones').value.trim(),
    notes: document.getElementById('f_notes').value.trim(),
    updatedAt: new Date().toISOString(),
  };

  if (editingId) {
    const idx = projects.findIndex(p => p.id === editingId);
    if (idx >= 0) projects[idx] = project;
  } else {
    projects.push(project);
  }

  closeModal();
  showToast('กำลังบันทึก...', '');
  await saveProjectToExcel(project);
  saveLocal();
  showToast('บันทึกเรียบร้อย ✓', 'success');
  renderCurrentPage();
}

async function deleteProject(id) {
  if (!confirm('ยืนยันลบโครงการนี้?')) return;
  projects = projects.filter(p => p.id !== id);
  saveLocal();
  // Re-write all rows to Excel
  showToast('กำลังลบ...', '');
  await rewriteAllToExcel();
  showToast('ลบเรียบร้อย ✓', 'success');
  renderCurrentPage();
}

async function rewriteAllToExcel() {
  if (!excelFileId) return;
  try {
    const rows = projects.map(p => HEADERS.map(h => p[h] ?? ''));
    const endRow = rows.length + 1;
    // Clear existing data rows (keep header)
    if (rows.length > 0) {
      await graphPatch(
        `${await wbPath()}/worksheets('Sheet1')/range(address='A2:P${endRow}')`,
        { values: rows }
      );
    }
  } catch (e) {
    console.error('rewriteAllToExcel error:', e);
  }
}

// ─── MODAL HELPERS ────────────────────────────────────────────────────────────
function openModal() { document.getElementById('projectModal').classList.add('open'); }
function closeModal() { document.getElementById('projectModal').classList.remove('open'); }

// ─── PPTX GENERATION ─────────────────────────────────────────────────────────
async function generateSingleProjectSlide(id) {
  const p = projects.find(x => x.id === id);
  if (!p) return;
  generatePPTXFromProjects([p], 'project_detail');
}

async function generatePPTX() {
  const selected = selectedProjects.size > 0
    ? projects.filter(p => selectedProjects.has(p.id))
    : projects;

  if (selected.length === 0) {
    showToast('กรุณาเลือกโครงการอย่างน้อย 1 โครงการ', 'error');
    return;
  }
  generatePPTXFromProjects(selected, selectedReportType);
}

function generatePPTXFromProjects(selectedProjs, reportType) {
  showToast('กำลังสร้าง PowerPoint...', '');

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'PM Report';

  const NAVY = '0F1B35';
  const NAVY2 = '162040';
  const TEAL = '00BFA6';
  const GOLD = 'FFB300';
  const RED = 'EF4444';
  const GREEN = '22C55E';
  const BLUE = '3B82F6';
  const WHITE = 'FFFFFF';
  const LGRAY = 'E2E8F0';
  const GRAY = '94A3B8';

  const today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  const typeLabel = { weekly: 'Weekly Status Report', monthly: 'Monthly Report', executive: 'Executive Summary', project_detail: 'Project Detail' }[reportType] || 'PM Report';

  // ── SLIDE 1: Cover ─────────────────────────────────────────────────────────
  const cover = pptx.addSlide();
  cover.background = { color: NAVY };

  // Accent bar left
  cover.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: TEAL } });

  // Decorative rect
  cover.addShape(pptx.ShapeType.rect, { x: 6.5, y: 0.8, w: 3.3, h: 4.0, fill: { color: NAVY2 }, line: { color: TEAL, width: 1 } });
  cover.addShape(pptx.ShapeType.rect, { x: 6.6, y: 0.9, w: 3.1, h: 3.8, fill: { color: NAVY2 } });

  // Grid lines decoration
  for (let i = 0; i < 5; i++) {
    cover.addShape(pptx.ShapeType.line, { x: 6.7, y: 1.2 + i * 0.6, w: 2.8, h: 0, line: { color: TEAL, width: 0.3, transparency: 70 } });
  }

  cover.addText(typeLabel.toUpperCase(), {
    x: 0.4, y: 1.2, w: 5.8, h: 0.5,
    fontSize: 13, bold: true, color: TEAL, charSpacing: 4, fontFace: 'Calibri'
  });
  cover.addText('รายงานสถานะโครงการ', {
    x: 0.4, y: 1.8, w: 5.8, h: 1.0,
    fontSize: 38, bold: true, color: WHITE, fontFace: 'Calibri'
  });
  cover.addText(`วันที่ ${today}`, {
    x: 0.4, y: 2.9, w: 5.8, h: 0.5,
    fontSize: 15, color: GRAY, fontFace: 'Calibri'
  });

  // Stats on cover
  const total = selectedProjs.length;
  const onTrack = selectedProjs.filter(p => p.status === 'on-track').length;
  const atRisk = selectedProjs.filter(p => p.status === 'at-risk').length;
  const delayed = selectedProjs.filter(p => p.status === 'delayed').length;

  const statItems = [
    { label: 'Total', val: total, color: TEAL },
    { label: 'On Track', val: onTrack, color: GREEN },
    { label: 'At Risk', val: atRisk, color: GOLD },
    { label: 'Delayed', val: delayed, color: RED },
  ];
  statItems.forEach((s, i) => {
    const x = 0.4 + i * 2.4;
    cover.addShape(pptx.ShapeType.rect, { x, y: 4.0, w: 2.1, h: 1.1, fill: { color: NAVY2 }, line: { color: s.color, width: 1 } });
    cover.addText(String(s.val), { x, y: 4.05, w: 2.1, h: 0.55, fontSize: 28, bold: true, color: s.color, align: 'center', fontFace: 'Calibri', margin: 0 });
    cover.addText(s.label, { x, y: 4.6, w: 2.1, h: 0.4, fontSize: 10, color: GRAY, align: 'center', fontFace: 'Calibri', margin: 0 });
  });

  // ── SLIDE 2: Portfolio Overview ────────────────────────────────────────────
  if (selectedProjs.length > 1) {
    const overview = pptx.addSlide();
    overview.background = { color: NAVY };
    overview.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: NAVY2 } });
    overview.addText('PORTFOLIO OVERVIEW', { x: 0.4, y: 0.12, w: 9, h: 0.46, fontSize: 14, bold: true, color: TEAL, charSpacing: 3, fontFace: 'Calibri', margin: 0 });

    // Table
    const tableData = [
      [
        { text: 'รหัส', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
        { text: 'ชื่อโครงการ', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
        { text: 'PM', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
        { text: 'สถานะ', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
        { text: 'คืบหน้า', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
        { text: 'สิ้นสุด', options: { bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 } },
      ],
      ...selectedProjs.map(p => {
        const statusColors = { 'on-track': GREEN, 'at-risk': GOLD, 'delayed': RED, 'completed': BLUE, 'planning': GRAY };
        const sc = statusColors[p.status] || GRAY;
        const statusLabels = { 'on-track': 'On Track', 'at-risk': 'At Risk', 'delayed': 'Delayed', 'completed': 'Done', 'planning': 'Planning' };
        return [
          { text: p.code || '', options: { fontSize: 10, color: GRAY, fontFace: 'Courier New' } },
          { text: p.name || '', options: { fontSize: 11, color: WHITE, bold: true } },
          { text: p.pm || '', options: { fontSize: 10, color: LGRAY } },
          { text: statusLabels[p.status] || p.status, options: { fontSize: 10, color: sc, bold: true } },
          { text: (p.progress || '0') + '%', options: { fontSize: 11, color: TEAL, bold: true, align: 'center' } },
          { text: p.endDate || '-', options: { fontSize: 10, color: GRAY, fontFace: 'Courier New' } },
        ];
      })
    ];

    overview.addTable(tableData, {
      x: 0.4, y: 0.85, w: 9.2, h: Math.min(4.5, 0.5 + selectedProjs.length * 0.55),
      border: { pt: 0.5, color: '1E2D52' },
      fill: { color: NAVY },
      colW: [1.1, 2.8, 1.3, 1.1, 1.0, 1.1],
      rowH: 0.5,
    });
  }

  // ── SLIDES per project ─────────────────────────────────────────────────────
  selectedProjs.forEach((p, idx) => {
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };

    // Header bar
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY2 } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.08, h: 0.72, fill: { color: TEAL } });

    // Status pill color
    const statusColors = { 'on-track': GREEN, 'at-risk': GOLD, 'delayed': RED, 'completed': BLUE, 'planning': GRAY };
    const statusLabels = { 'on-track': '● ON TRACK', 'at-risk': '● AT RISK', 'delayed': '● DELAYED', 'completed': '● DONE', 'planning': '● PLANNING' };
    const sc = statusColors[p.status] || GRAY;

    slide.addText(p.code || '', { x: 0.2, y: 0.1, w: 1.2, h: 0.28, fontSize: 9, color: TEAL, fontFace: 'Courier New', bold: true, margin: 0 });
    slide.addText(p.name || 'ไม่มีชื่อ', { x: 0.2, y: 0.34, w: 7, h: 0.3, fontSize: 16, bold: true, color: WHITE, fontFace: 'Calibri', margin: 0 });
    slide.addText(statusLabels[p.status] || '', { x: 7.8, y: 0.2, w: 1.9, h: 0.32, fontSize: 10, bold: true, color: sc, align: 'right', fontFace: 'Calibri', margin: 0 });

    // Progress section (left)
    const pct = parseInt(p.progress) || 0;
    const progColor = pct >= 70 ? TEAL : pct >= 40 ? GOLD : RED;

    slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.85, w: 4.2, h: 1.8, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
    slide.addText('PROGRESS', { x: 0.45, y: 0.95, w: 2, h: 0.3, fontSize: 9, color: GRAY, charSpacing: 2, fontFace: 'Calibri', margin: 0 });
    slide.addText(pct + '%', { x: 0.45, y: 1.25, w: 1.4, h: 0.7, fontSize: 40, bold: true, color: progColor, fontFace: 'Calibri', margin: 0 });

    // Progress bar
    slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 2.15, w: 3.8, h: 0.12, fill: { color: '1E2D52' } });
    if (pct > 0) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 2.15, w: Math.max(3.8 * pct / 100, 0.1), h: 0.12, fill: { color: progColor } });
    }

    // PM & Client
    slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 2.8, w: 2.0, h: 0.9, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
    slide.addText('PM', { x: 0.45, y: 2.88, w: 1.7, h: 0.22, fontSize: 8, color: GRAY, charSpacing: 2, margin: 0, fontFace: 'Calibri' });
    slide.addText(p.pm || '-', { x: 0.45, y: 3.1, w: 1.7, h: 0.5, fontSize: 12, bold: true, color: WHITE, margin: 0, fontFace: 'Calibri' });

    slide.addShape(pptx.ShapeType.rect, { x: 2.45, y: 2.8, w: 2.05, h: 0.9, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
    slide.addText('CLIENT', { x: 2.6, y: 2.88, w: 1.7, h: 0.22, fontSize: 8, color: GRAY, charSpacing: 2, margin: 0, fontFace: 'Calibri' });
    slide.addText(p.client || '-', { x: 2.6, y: 3.1, w: 1.7, h: 0.5, fontSize: 12, bold: true, color: WHITE, margin: 0, fontFace: 'Calibri' });

    // Timeline
    slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 3.85, w: 4.2, h: 0.7, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
    slide.addText('📅  ' + (p.startDate || '-') + '  →  ' + (p.endDate || '-'), {
      x: 0.45, y: 3.93, w: 3.9, h: 0.54, fontSize: 11, color: LGRAY, fontFace: 'Courier New', margin: 0
    });

    // Right panel: Notes / Update
    slide.addShape(pptx.ShapeType.rect, { x: 4.9, y: 0.85, w: 4.8, h: 2.0, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
    slide.addText('UPDATE ล่าสุด', { x: 5.05, y: 0.93, w: 4.5, h: 0.28, fontSize: 9, color: GRAY, charSpacing: 2, margin: 0, fontFace: 'Calibri' });
    slide.addText(p.notes || 'ยังไม่มี update', {
      x: 5.05, y: 1.22, w: 4.5, h: 1.45,
      fontSize: 11, color: WHITE, fontFace: 'Calibri', valign: 'top', margin: 4, wrap: true
    });

    // Risks
    if (p.risks) {
      slide.addShape(pptx.ShapeType.rect, { x: 4.9, y: 3.0, w: 4.8, h: 1.55, fill: { color: NAVY2 }, line: { color: RED, width: 0.75, transparency: 50 } });
      slide.addText('⚠  RISKS & ISSUES', { x: 5.05, y: 3.07, w: 4.5, h: 0.28, fontSize: 9, color: RED, charSpacing: 2, margin: 0, fontFace: 'Calibri' });
      slide.addText(p.risks, {
        x: 5.05, y: 3.35, w: 4.5, h: 1.1,
        fontSize: 10, color: LGRAY, fontFace: 'Calibri', valign: 'top', margin: 4, wrap: true
      });
    }

    // Budget row
    if (p.budget) {
      const budget = parseInt(p.budget) || 0;
      const spent = parseInt(p.spent) || 0;
      const bpct = budget > 0 ? Math.round(spent / budget * 100) : 0;
      const bc = bpct > 90 ? RED : bpct > 70 ? GOLD : GREEN;
      slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 4.7, w: 9.4, h: 0.7, fill: { color: NAVY2 }, line: { color: '1E2D52', width: 1 } });
      slide.addText('BUDGET', { x: 0.5, y: 4.78, w: 1.2, h: 0.24, fontSize: 8, color: GRAY, charSpacing: 2, margin: 0, fontFace: 'Calibri' });
      slide.addText(budget.toLocaleString() + ' บาท', { x: 1.7, y: 4.75, w: 2.5, h: 0.55, fontSize: 13, bold: true, color: WHITE, margin: 0, fontFace: 'Calibri' });
      slide.addText('ใช้ไป: ' + spent.toLocaleString() + ' (' + bpct + '%)', { x: 4.5, y: 4.78, w: 4, h: 0.54, fontSize: 11, color: bc, margin: 0, fontFace: 'Calibri', align: 'right' });
    }

    // Slide number
    slide.addText(`${idx + 2}`, { x: 9.5, y: 5.3, w: 0.4, h: 0.25, fontSize: 9, color: GRAY, align: 'right', margin: 0, fontFace: 'Calibri' });
  });

  // ── FINAL SLIDE: Thank You ─────────────────────────────────────────────────
  const last = pptx.addSlide();
  last.background = { color: NAVY };
  last.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.12, fill: { color: TEAL } });
  last.addShape(pptx.ShapeType.rect, { x: 0, y: 5.505, w: 10, h: 0.12, fill: { color: TEAL } });
  last.addText('Thank You', { x: 1, y: 1.8, w: 8, h: 1.2, fontSize: 48, bold: true, color: WHITE, align: 'center', fontFace: 'Calibri' });
  last.addText('PM Dashboard  •  ' + today, { x: 1, y: 3.2, w: 8, h: 0.5, fontSize: 13, color: GRAY, align: 'center', fontFace: 'Calibri' });

  // Save
  const filename = `PM_${typeLabel.replace(/\s/g,'_')}_${new Date().toISOString().slice(0,10)}.pptx`;
  pptx.writeFile({ fileName: filename })
    .then(() => showToast('ดาวน์โหลด ' + filename + ' เรียบร้อย ✓', 'success'))
    .catch(e => showToast('เกิดข้อผิดพลาด: ' + e.message, 'error'));
}

// ─── TOAST ────────────────────────────────────────────────────────────────────
function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.innerHTML = type === '' ? `<div class="spinner"></div> ${msg}` : msg;
  t.className = 'show ' + type;
  clearTimeout(t._timeout);
  t._timeout = setTimeout(() => { t.classList.remove('show'); }, 3000);
}
// ─── FINANCIAL & RESOURCES ────────────────────────────────────────────────────
let budgetItems = JSON.parse(localStorage.getItem('fin_budget') || '[]');
let timesheetItems = JSON.parse(localStorage.getItem('fin_timesheet') || '[]');
let currentFinTab = 'budget';
let editingBudgetId = null;

function saveFin() {
  localStorage.setItem('fin_budget', JSON.stringify(budgetItems));
  localStorage.setItem('fin_timesheet', JSON.stringify(timesheetItems));
}

// ── RENDER FINANCIALS PAGE ────────────────────────────────────────────────────
function renderFinancials() {
  const totalBudget = projects.reduce((a, p) => a + (parseInt(p.budget) || 0), 0);
  const totalSpent  = projects.reduce((a, p) => a + (parseInt(p.spent)  || 0), 0);
  const totalHours  = timesheetItems.reduce((a, t) => a + (parseInt(t.total) || 0), 0);
  const variance    = totalSpent - totalBudget;
  const varColor    = variance > 0 ? 'var(--red)' : 'var(--green)';

  // Summary stats
  const statsHtml = `
    <div class="fin-grid">
      <div class="stat-card teal">
        <div class="stat-label">งบประมาณรวม</div>
        <div class="stat-value">${(totalBudget/1000000).toFixed(2)}M</div>
        <div class="stat-sub">ทุกโครงการ</div>
      </div>
      <div class="stat-card gold">
        <div class="stat-label">ใช้จริงรวม</div>
        <div class="stat-value">${(totalSpent/1000000).toFixed(2)}M</div>
        <div class="stat-sub">${totalBudget ? Math.round(totalSpent/totalBudget*100) : 0}% ของงบ</div>
      </div>
      <div class="stat-card ${variance > 0 ? 'red' : 'blue'}">
        <div class="stat-label">Variance</div>
        <div class="stat-value" style="color:${varColor}">${variance >= 0 ? '+' : ''}${(variance/1000).toFixed(0)}K</div>
        <div class="stat-sub">${variance > 0 ? 'เกินงบ' : 'ประหยัดงบ'}</div>
      </div>
      <div class="stat-card blue">
        <div class="stat-label">ชั่วโมงรวม</div>
        <div class="stat-value">${totalHours.toLocaleString()}</div>
        <div class="stat-sub">Man-hours</div>
      </div>
    </div>`;

  // Tabs
  const tabsHtml = `
    <div class="fin-tabs">
      <button class="fin-tab ${currentFinTab==='budget'?'active':''}" onclick="switchFinTab('budget')">💰 Budget vs Actual</button>
      <button class="fin-tab ${currentFinTab==='timesheet'?'active':''}" onclick="switchFinTab('timesheet')">🕐 Timesheet</button>
    </div>`;

  // Budget tab
  const CATS = ['Labour','Materials','Equipment','Subcontract','Overhead','Other'];
  const catTotals = CATS.map(cat => {
    const items = budgetItems.filter(b => b.cat === cat);
    return {
      cat,
      est: items.reduce((a, b) => a + (parseInt(b.est)||0), 0),
      act: items.reduce((a, b) => a + (parseInt(b.act)||0), 0),
    };
  }).filter(c => c.est > 0 || c.act > 0);

  const maxVal = Math.max(...catTotals.map(c => Math.max(c.est, c.act)), 1);

  const budgetTabHtml = `
    <div class="grid-2">
      <div class="table-card" style="padding:0">
        <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--border)">
          <span style="font-size:14px;font-weight:600">Budget vs Actual — รายหมวด</span>
          <button class="btn btn-primary" style="font-size:12px;padding:6px 12px" onclick="openBudgetModal()">+ เพิ่มรายการ</button>
        </div>
        <div style="padding:14px 18px">
          ${catTotals.length === 0 ? `<div class="empty" style="padding:32px"><div class="empty-icon">💰</div><div class="empty-text">ยังไม่มีรายการ กด "+ เพิ่มรายการ"</div></div>` :
            catTotals.map(c => {
              const estW = (c.est/maxVal*100).toFixed(0);
              const actW = (c.act/maxVal*100).toFixed(0);
              const actColor = c.act > c.est ? 'var(--red)' : 'var(--gold)';
              return `<div class="cost-row">
                <div class="cost-cat">${c.cat}</div>
                <div class="cost-bars">
                  <div class="cost-bar-row">
                    <div class="cost-bar-label" style="color:var(--blue)">Est.</div>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:${estW}%;background:var(--blue);opacity:0.6"></div></div>
                    <div class="cost-amount" style="color:var(--blue)">${(c.est/1000).toFixed(0)}K</div>
                  </div>
                  <div class="cost-bar-row">
                    <div class="cost-bar-label" style="color:${actColor}">Act.</div>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:${actW}%;background:${actColor}"></div></div>
                    <div class="cost-amount" style="color:${actColor}">${(c.act/1000).toFixed(0)}K</div>
                  </div>
                </div>
              </div>`;
            }).join('')
          }
        </div>
      </div>

      <div class="table-card" style="padding:0">
        <div style="padding:14px 18px;border-bottom:1px solid var(--border)">
          <span style="font-size:14px;font-weight:600">รายการทั้งหมด</span>
        </div>
        <table class="project-table">
          <thead><tr><th>โครงการ</th><th>หมวด</th><th>งบ (฿)</th><th>จริง (฿)</th><th>หมายเหตุ</th><th></th></tr></thead>
          <tbody>
            ${budgetItems.length === 0
              ? `<tr><td colspan="6"><div class="empty" style="padding:24px"><div class="empty-icon">📋</div><div class="empty-text">ยังไม่มีรายการ</div></div></td></tr>`
              : budgetItems.map(b => {
                  const proj = projects.find(p => p.id === b.projectId);
                  const over = (parseInt(b.act)||0) > (parseInt(b.est)||0);
                  return `<tr>
                    <td style="font-size:12px">${proj ? proj.code : '-'}</td>
                    <td><span class="badge" style="background:rgba(0,191,166,0.12);color:var(--teal)">${b.cat}</span></td>
                    <td style="font-family:'IBM Plex Mono',monospace;font-size:12px">${parseInt(b.est||0).toLocaleString()}</td>
                    <td style="font-family:'IBM Plex Mono',monospace;font-size:12px;color:${over?'var(--red)':'var(--green)'}">${parseInt(b.act||0).toLocaleString()}</td>
                    <td style="font-size:12px;color:var(--text-dim)">${b.note||'-'}</td>
                    <td><button class="btn btn-danger" style="padding:3px 8px;font-size:11px" onclick="deleteBudgetItem('${b.id}')">ลบ</button></td>
                  </tr>`;
                }).join('')
            }
          </tbody>
        </table>
      </div>
    </div>`;

  // Timesheet tab
  const weeks = [...new Set(timesheetItems.map(t => t.week))].sort().reverse().slice(0, 4);
  const timesheetTabHtml = `
    <div class="section-header" style="margin-bottom:16px">
      <div class="section-title">Timesheet — บันทึกชั่วโมงการทำงาน</div>
      <button class="btn btn-primary" style="font-size:12px;padding:6px 12px" onclick="openTimesheetModal()">+ บันทึก Timesheet</button>
    </div>
    ${timesheetItems.length === 0
      ? `<div class="table-card"><div class="empty"><div class="empty-icon">🕐</div><div class="empty-text">ยังไม่มีรายการ กด "+ บันทึก Timesheet"</div></div></div>`
      : `<div class="table-card">
          <table class="project-table">
            <thead><tr><th>โครงการ</th><th>พนักงาน</th><th>สัปดาห์</th><th>จ.</th><th>อ.</th><th>พ.</th><th>พฤ.</th><th>ศ.</th><th>รวม</th><th></th></tr></thead>
            <tbody>
              ${timesheetItems.map(t => {
                const proj = projects.find(p => p.id === t.projectId);
                return `<tr>
                  <td style="font-size:12px">${proj ? proj.code : '-'}</td>
                  <td style="font-size:12px;font-weight:500">${t.name}</td>
                  <td style="font-size:11px;font-family:'IBM Plex Mono',monospace;color:var(--text-dim)">${t.week}</td>
                  ${['mon','tue','wed','thu','fri'].map(d => `<td class="ts-cell">${t[d]>0?`<span class="ts-fill">${t[d]}</span>`:'-'}</td>`).join('')}
                  <td style="font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--teal);font-weight:600">${t.total}h</td>
                  <td><button class="btn btn-danger" style="padding:3px 8px;font-size:11px" onclick="deleteTimesheetItem('${t.id}')">ลบ</button></td>
                </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>`
    }`;

  document.getElementById('pageContent').innerHTML = `
    ${statsHtml}
    ${tabsHtml}
    <div id="fin-tab-content">
      ${currentFinTab === 'budget' ? budgetTabHtml : timesheetTabHtml}
    </div>`;
}

function switchFinTab(tab) {
  currentFinTab = tab;
  renderFinancials();
}

// ── BUDGET MODAL ──────────────────────────────────────────────────────────────
function openBudgetModal() {
  editingBudgetId = null;
  document.getElementById('budgetModalTitle').textContent = 'เพิ่มรายการงบประมาณ';
  const sel = document.getElementById('fb_project');
  sel.innerHTML = projects.map(p => `<option value="${p.id}">${p.code} — ${p.name}</option>`).join('');
  document.getElementById('fb_cat').value = 'Labour';
  document.getElementById('fb_est').value = '';
  document.getElementById('fb_act').value = '';
  document.getElementById('fb_note').value = '';
  document.getElementById('budgetModal').classList.add('open');
}

function closeBudgetModal() {
  document.getElementById('budgetModal').classList.remove('open');
}

function saveBudgetItem() {
  const item = {
    id: 'b_' + Date.now(),
    projectId: document.getElementById('fb_project').value,
    cat: document.getElementById('fb_cat').value,
    est: document.getElementById('fb_est').value,
    act: document.getElementById('fb_act').value,
    note: document.getElementById('fb_note').value.trim(),
  };
  budgetItems.push(item);
  saveFin();
  closeBudgetModal();
  showToast('บันทึกรายการงบแล้ว ✓', 'success');
  renderFinancials();
}

function deleteBudgetItem(id) {
  if (!confirm('ยืนยันลบรายการนี้?')) return;
  budgetItems = budgetItems.filter(b => b.id !== id);
  saveFin();
  renderFinancials();
  showToast('ลบแล้ว', 'success');
}

// ── TIMESHEET MODAL ───────────────────────────────────────────────────────────
function openTimesheetModal() {
  const sel = document.getElementById('ft_project');
  sel.innerHTML = projects.map(p => `<option value="${p.id}">${p.code} — ${p.name}</option>`).join('');
  document.getElementById('ft_name').value = '';
  document.getElementById('ft_week').value = new Date().toISOString().slice(0, 10);
  ['mon','tue','wed','thu','fri'].forEach(d => document.getElementById('ft_'+d).value = '');
  document.getElementById('ft_total').value = '';
  // Auto-calculate total
  ['mon','tue','wed','thu','fri'].forEach(d => {
    document.getElementById('ft_'+d).oninput = () => {
      const total = ['mon','tue','wed','thu','fri'].reduce((a, x) => a + (parseInt(document.getElementById('ft_'+x).value)||0), 0);
      document.getElementById('ft_total').value = total;
    };
  });
  document.getElementById('timesheetModal').classList.add('open');
}

function closeTimesheetModal() {
  document.getElementById('timesheetModal').classList.remove('open');
}

function saveTimesheetItem() {
  const name = document.getElementById('ft_name').value.trim();
  if (!name) { showToast('กรุณากรอกชื่อพนักงาน', 'error'); return; }
  const days = ['mon','tue','wed','thu','fri'];
  const vals = {};
  days.forEach(d => vals[d] = parseInt(document.getElementById('ft_'+d).value)||0);
  const total = days.reduce((a, d) => a + vals[d], 0);
  const item = {
    id: 'ts_' + Date.now(),
    projectId: document.getElementById('ft_project').value,
    name,
    week: document.getElementById('ft_week').value,
    ...vals,
    total,
  };
  timesheetItems.push(item);
  saveFin();
  closeTimesheetModal();
  showToast('บันทึก Timesheet แล้ว ✓', 'success');
  renderFinancials();
}

function deleteTimesheetItem(id) {
  if (!confirm('ยืนยันลบรายการนี้?')) return;
  timesheetItems = timesheetItems.filter(t => t.id !== id);
  saveFin();
  renderFinancials();
  showToast('ลบแล้ว', 'success');
}
