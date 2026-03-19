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
  const navMap = { dashboard: 0, projects: 1, report: 3 };
  const navItems = document.querySelectorAll('.nav-item');
  if (navMap[page] !== undefined) navItems[navMap[page]]?.classList.add('active');
  renderCurrentPage();
}

function renderCurrentPage() {
  const titles = { dashboard: 'ภาพรวม', projects: 'โครงการทั้งหมด', report: 'สร้างรายงาน / สไลด์' };
  document.getElementById('pageTitle').textContent = titles[currentPage] || '';
  switch (currentPage) {
    case 'dashboard': renderDashboard(); break;
    case 'projects': renderProjects(); break;
    case 'detail': renderDetail(); break;
    case 'report': renderReport(); break;
  }
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function renderDashboard() {
  const total = projects.length;
  const onTrack = projects.filter(p => p.status === 'on-track').length;
  const atRisk = projects.filter(p => p.status === 'at-risk').length;
  const delayed = projects.filter(p => p.status === 'delayed').length;
  const completed = projects.filter(p => p.status === 'completed').length;

  const recent = [...projects].sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt)).slice(0, 5);

  document.getElementById('pageContent').innerHTML = `
    <div class="stat-grid">
      <div class="stat-card teal">
        <div class="stat-label">โครงการทั้งหมด</div>
        <div class="stat-value">${total}</div>
        <div class="stat-sub">โครงการในระบบ</div>
      </div>
      <div class="stat-card blue">
        <div class="stat-label">On Track</div>
        <div class="stat-value">${onTrack}</div>
        <div class="stat-sub">เป็นไปตามแผน</div>
      </div>
      <div class="stat-card gold">
        <div class="stat-label">At Risk</div>
        <div class="stat-value">${atRisk}</div>
        <div class="stat-sub">ต้องติดตาม</div>
      </div>
      <div class="stat-card red">
        <div class="stat-label">Delayed</div>
        <div class="stat-value">${delayed}</div>
        <div class="stat-sub">ล่าช้ากว่าแผน</div>
      </div>
    </div>

    <div class="section-header">
      <div class="section-title">อัปเดตล่าสุด</div>
      <button class="btn btn-primary" onclick="openAddProject()">+ เพิ่มโครงการ</button>
    </div>
    <div class="table-card">
      ${recent.length === 0 ? `<div class="empty"><div class="empty-icon">📁</div><div class="empty-text">ยังไม่มีโครงการ กด "เพิ่มโครงการ" เพื่อเริ่มต้น</div></div>` :
        `<table class="project-table">
          <thead><tr>
            <th>รหัส</th><th>ชื่อโครงการ</th><th>PM</th><th>สถานะ</th><th>ความคืบหน้า</th><th>วันสิ้นสุด</th>
          </tr></thead>
          <tbody>
            ${recent.map(p => projectRow(p)).join('')}
          </tbody>
        </table>`
      }
    </div>
  `;
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