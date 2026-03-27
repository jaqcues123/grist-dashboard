// ═══════════════════════════════════════════════════════════
// NEWS BUILD DASHBOARD — JS BLOCK
// ═══════════════════════════════════════════════════════════

const charts = {};

let DB = {
  backlog:      [],
  pipeline:     [],
  bodyClass:    [],
  monthSummary: [],
  pipelineCycle:[],
  velocity:     [],
  nrcOrders:    [],
};

const localState = {};
function getLocal(id, key, fallback) {
  return (localState[id] && localState[id][key] !== undefined) ? localState[id][key] : fallback;
}
function setLocal(id, key, val) {
  if (!localState[id]) localState[id] = {};
  localState[id][key] = val;
}

let activePage = 'operations';
let opsFilters = { name: 'ANY', chassis: 'ANY', body: 'ANY' };
let invState = {
  snapDate: (() => { const d = new Date(); d.setMonth(d.getMonth()+6); d.setHours(0,0,0,0); return d; })(),
  body: 'ANY', chassis: 'ANY', color: 'ANY', groupBy: 'body', projMode: 'none'
};
let cfPeriod  = 'week';
let cfHorizon = 6;
let cfSelectedWeekStart = null;
let ganttRange    = 'short';
let ganttZoom     = 'day';
let ganttSortMode = 'date';
let ganttTasks    = [];
let ganttSelected = null;
let ganttCellWidth = 40;
let ganttDays     = 60;
let ganttStart    = null;

// Sales state — two independent filters
let chartYear     = 'ALL';   // filters monthly revenue + body class charts
let selectedMonth = new Date().getMonth() + 1;
let selectedYear  = new Date().getFullYear();

let nrcFilters = {
  from:   (() => { const d = new Date(); d.setHours(0,0,0,0); return d; })(),
  to:     null,
  status: 'ALL'
};

// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function escHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  if (Array.isArray(val) && val[0] === 'd') return new Date(val[1] * 1000);
  if (typeof val === 'number') {
    if (val > 100000000) return new Date(val * 1000);  // unix seconds
    if (val > 0)         return new Date(val * 86400 * 1000); // grist day count
  }
  if (typeof val === 'string' && val.trim()) { const d = new Date(val.trim()); return isNaN(d) ? null : d; }
  return null;
}

function fmtDate(val) {
  const d = parseDate(val);
  if (!d) return '';
  return (d.getMonth()+1) + '/' + d.getDate() + '/' + String(d.getFullYear()).slice(2);
}

function toInputDate(val) {
  const d = parseDate(val);
  if (!d) return '';
  return d.toISOString().slice(0,10);
}

function fmtCurrency(n, compact) {
  if (n == null || isNaN(n)) return '—';
  if (compact) {
    if (Math.abs(n) >= 1000000) return '$' + (n/1000000).toFixed(1) + 'M';
    if (Math.abs(n) >= 1000)    return '$' + (n/1000).toFixed(0) + 'K';
    return '$' + Math.round(n).toLocaleString();
  }
  return '$' + Math.round(n).toLocaleString();
}

function fmtPct(n) {
  if (n == null || isNaN(n)) return '—';
  return (n * 100).toFixed(1) + '%';
}

function showToast(msg, isError) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.className = 'show' + (isError ? ' error' : '');
  setTimeout(() => t.className = '', 2500);
}

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

// ═══════════════════════════════════════════════════════════
// SECONDARY TABLE FETCH
// ═══════════════════════════════════════════════════════════

async function fetchTable(tableId) {
  const result = await grist.docApi.fetchTable(tableId);
  const ids  = result.id || [];
  const cols = Object.keys(result).filter(k => k !== 'id' && k !== 'manualSort');
  return ids.map((id, i) => {
    const row = { id };
    cols.forEach(col => { row[col] = result[col][i]; });
    return row;
  });
}

async function loadSecondaryTables() {
  try {
    const [pipeline, bodyClass, monthSummary, pipelineCycle, velocity, nrcOrders] = await Promise.all([
      fetchTable('NEWS_Pipeline'),
      fetchTable('NEWS_Build_Backlog_summary_Body_Class'),
      fetchTable('NEWS_Build_Backlog_summary_Sales_Month_Year'),
      fetchTable('NEWS_Pipeline_summary_Sales_Cycle'),
      fetchTable('Sales_Velocity_summary_Body_Class'),
      fetchTable('NRC_Order_Tracking'),
    ]);

    DB.pipeline      = pipeline;
    DB.bodyClass     = bodyClass;
    DB.monthSummary  = monthSummary;
    DB.pipelineCycle = pipelineCycle;
    DB.velocity      = velocity;
    DB.nrcOrders = nrcOrders;
    document.getElementById('last-updated').textContent = 'Updated ' + new Date().toLocaleTimeString();
  } catch(e) {
    console.error('loadSecondaryTables failed:', e);
    showToast('Error loading tables: ' + e.message, true);
  }
}

// ═══════════════════════════════════════════════════════════
// WRITE BACK
// ═══════════════════════════════════════════════════════════

async function writeBoolToGrist(rowId, colId, value, toggleEl) {
  toggleEl.classList.add('saving');
  try {
    await grist.docApi.applyUserActions([['UpdateRecord', 'NEWS_Build_Backlog', rowId, { [colId]: value }]]);
    toggleEl.classList.remove('saving');
    const rec = DB.backlog.find(b => b.id === rowId);
    if (rec) rec[colId] = value;
    showToast('Saved');
  } catch(e) {
    toggleEl.classList.remove('saving');
    toggleEl.classList.add('error');
    setTimeout(() => toggleEl.classList.remove('error'), 2000);
    showToast('Save failed: ' + e.message, true);
    toggleEl.querySelector('input').checked = !value;
  }
}

// ═══════════════════════════════════════════════════════════
// SECTION LOGIC
// ═══════════════════════════════════════════════════════════

function getSection(rec) {
  const started = !!rec.Build_Started_, ended = !!rec.Build_Ended_;
  const delivered = !!rec.Delivered_, paid = !!rec.Paid_, toolbox = !!rec.Toolbox;
  if (started && !ended)            return 'inservice';
  if (ended && !delivered && !paid) return toolbox ? 'forsale' : 'notoolbox';
  if (delivered && !paid)           return 'awaitpay';
  if (paid && !delivered)           return 'awaitdel';
  return null;
}

const SECTIONS = [
  { key:'inservice',  label:'Units In Service / In Progress',          cls:'sec-inservice',  showStatus:true,  showValue:false },
  { key:'forsale',    label:'Completed Units For Sale W/ Toolbox',     cls:'sec-forsale',    showStatus:false, showValue:true  },
  { key:'notoolbox',  label:'Completed Units For Sale — W/O Toolbox',  cls:'sec-notoolbox',  showStatus:false, showValue:true  },
  { key:'awaitpay',   label:'Completed Units Awaiting Payment',        cls:'sec-awaitpay',   showStatus:false, showValue:true  },
  { key:'awaitdel',   label:'Completed Units Awaiting Delivery',       cls:'sec-awaitdel',   showStatus:false, showValue:true  },
];
const STATUS_OPTS = ['—','In Work','On Hold','In Coming','Completed','Complete'];
const LOC_OPTS    = ['','Lot','Shop','Outside','Storage','Other'];

// ═══════════════════════════════════════════════════════════
// PAGE 1 — OPERATIONS
// ═══════════════════════════════════════════════════════════

function renderOperations() {
  renderKPIs();
  populateOpsFilters();
  renderBacklog();
}

function renderKPIs() {
  const b = DB.backlog;
  const inprog   = b.filter(r => getSection(r) === 'inservice');
  const onlot    = b.filter(r => getSection(r) === 'forsale' || getSection(r) === 'notoolbox');
  const awaitPay = b.filter(r => getSection(r) === 'awaitpay');
  const awaitDel = b.filter(r => getSection(r) === 'awaitdel');
  const sum = (arr, col) => arr.reduce((s,r) => s + (r[col]||0), 0);
  document.getElementById('kpi-inprog-count').textContent   = inprog.length;
  document.getElementById('kpi-onlot-count').textContent    = onlot.length;
  document.getElementById('kpi-onlot-value').textContent    = fmtCurrency(sum(onlot,'TOTAL_RETAIL_SALE'), true);
  document.getElementById('kpi-onlot-cost').textContent =
  fmtCurrency(
    onlot.reduce((total, r) =>
      total +
      (r.Body_Cost || 0) +
      (r.Actual_Labor_Cost || 0) +
      (r.INTERNAL_BUILD_COST_PARTS_ || 0),
    0),true
  );
  document.getElementById('kpi-payment-count').textContent  = awaitPay.length;
  document.getElementById('kpi-payment-value').textContent  = fmtCurrency(sum(awaitPay,'TOTAL_RETAIL_SALE'), true);
  document.getElementById('kpi-delivery-count').textContent = awaitDel.length;
  document.getElementById('kpi-delivery-value').textContent = fmtCurrency(sum(awaitDel,'TOTAL_RETAIL_SALE'), true);
 // Deliveries month to date
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear  = now.getFullYear();
  const mtdSales = DB.backlog.filter(r => {
    const d = parseDate(r.Date_of_Delivery);
    return d && d.getMonth() + 1 === currentMonth && d.getFullYear() === currentYear;
  });
  const mtdTotal = mtdSales.reduce((s, r) => s + (r.TOTAL_RETAIL_SALE || 0), 0);
  document.getElementById('kpi-cash-net').textContent = fmtCurrency(mtdTotal, true);
  document.getElementById('kpi-cash-sub').textContent = `${mtdSales.length} delivered${mtdSales.length !== 1 ? '' : ''} in ${now.toLocaleString('en-US', { month: 'long' })}`;
  document.getElementById('kpi-cash-net').style.color = 'var(--kpi-cash)';
  }

function populateOpsFilters() {
  const names = new Set(), chassis = new Set(), bodies = new Set();
  DB.backlog.forEach(r => {
    if (r.NAME)          names.add(String(r.NAME));
    if (r.CHASSIS_MODEL) chassis.add(String(r.CHASSIS_MODEL));
    if (r.Body_Class)    bodies.add(String(r.Body_Class));
  });
  rebuildSelect('ops-filter-name',    names,   opsFilters.name,    '— All Customers —');
  rebuildSelect('ops-filter-chassis', chassis, opsFilters.chassis, '— All Chassis —');
  rebuildSelect('ops-filter-body',    bodies,  opsFilters.body,    '— All Bodies —');
}

function rebuildSelect(elId, set, current, placeholder) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = `<option value="ANY">${placeholder || '— Any —'}</option>`;
  [...set].sort().forEach(v => {
    const o = document.createElement('option');
    o.value = v; o.textContent = v; el.appendChild(o);
  });
  el.value = (current !== 'ANY' && set.has(current)) ? current : 'ANY';
}

function applyOpsFilters(records) {
  return records.filter(r => {
    if (opsFilters.name    !== 'ANY' && String(r.NAME||'')          !== opsFilters.name)    return false;
    if (opsFilters.chassis !== 'ANY' && String(r.CHASSIS_MODEL||'') !== opsFilters.chassis) return false;
    if (opsFilters.body    !== 'ANY' && String(r.Body_Class||'')    !== opsFilters.body)    return false;
    return true;
  });
}

function renderBacklog() {
  const filtered = applyOpsFilters(DB.backlog);
  const buckets  = {};
  SECTIONS.forEach(s => buckets[s.key] = []);
  filtered.forEach(r => { const sec = getSection(r); if (sec && buckets[sec]) buckets[sec].push(r); });
  const area = document.getElementById('backlog-area');
  area.innerHTML = SECTIONS.map(sec => buildSectionTable(sec, buckets[sec.key])).join('');
  area.querySelectorAll('.bool-toggle input').forEach(inp => {
    inp.addEventListener('change', async e => {
      const wrap = e.target.closest('.bool-toggle');
      const rowId = parseInt(wrap.dataset.row);
      const col   = wrap.dataset.col;
      const val   = e.target.checked;
      await writeBoolToGrist(rowId, col, val, wrap);
      setTimeout(() => { renderKPIs(); renderBacklog(); }, 300);
    });
  });
  area.querySelectorAll('[data-local-key]').forEach(el => {
    el.addEventListener('change', e => {
      setLocal(parseInt(e.target.dataset.id), e.target.dataset.localKey, e.target.value);
    });
  });
}

function buildSectionTable(section, records) {
  const { label, cls, showStatus, showValue } = section;
  const totalValue = showValue ? records.reduce((s,r) => s + (r.TOTAL_RETAIL_SALE||0), 0) : 0;
  const valueStr   = showValue && totalValue ? `<span class="s-value">${fmtCurrency(totalValue, true)} total retail</span>` : '';
  const hdr  = `<tr class="section-hdr"><td colspan="20">${escHtml(label)}<span class="s-count">${records.length}</span>${valueStr}</td></tr>`;
  const rows = records.length ? records.map(r => buildDataRow(r, showStatus)).join('') : `<tr class="empty-row"><td colspan="20">No records in this section</td></tr>`;
  return `<table class="section-table ${cls}"><thead>${hdr}${buildColHeader(showStatus)}</thead><tbody>${rows}</tbody></table>`;
}

function buildColHeader(showStatus) {
  let cols = '';
  if (showStatus) cols += '<th>Status</th>';
  cols += `<th>OP #</th><th>SO #</th><th>Tech</th><th>Customer</th>
    <th>Paid Date</th><th>Chassis / Model</th><th>VIN #</th>
    <th>Chassis Loc</th><th>Body Class</th><th>Body Loc</th>
    <th>Serial #</th><th>Build Start</th><th>PDI Date</th>
    <th>B.Started</th><th>B.Ended</th><th>Delivered</th><th>Paid</th>
    <th>Body P/U</th><th>Chassis Del</th>`;
  return `<tr class="col-hdr">${cols}</tr>`;
}

function buildDataRow(rec, showStatus) {
  const id = rec.id;
  const status = getLocal(id,'status','—'), tech = getLocal(id,'tech','');
  const chassisLoc = getLocal(id,'chassisLoc',''), bodyLoc = getLocal(id,'bodyLoc','');
  const pdiDate = getLocal(id,'pdiDate','');
  let cells = '';
  if (showStatus) {
    const opts = STATUS_OPTS.map(s => `<option value="${s}" ${s===status?'selected':''}>${s}</option>`).join('');
    cells += `<td><select class="status-sel" data-id="${id}" data-local-key="status">${opts}</select></td>`;
  }
  cells += `<td>${escHtml(rec.OP_||'')}</td><td>${escHtml(rec.SO_||'')}</td>`;
  cells += `<td><input class="cell-in" value="${escHtml(tech)}" data-id="${id}" data-local-key="tech" placeholder="—"/></td>`;
  cells += `<td title="${escHtml(rec.NAME||'')}">${escHtml(rec.NAME||'')}</td>`;
  cells += `<td>${fmtDate(rec.Paid_Date)}</td>`;
  cells += `<td title="${escHtml(rec.CHASSIS_MODEL||'')}">${escHtml(rec.CHASSIS_MODEL||'')}</td>`;
  cells += `<td>${escHtml(rec.CHASSIS_VIN_||'')}</td>`;
  const cLocOpts = LOC_OPTS.map(o => `<option value="${o}" ${o===chassisLoc?'selected':''}>${o||'—'}</option>`).join('');
  cells += `<td><select class="cell-sel" data-id="${id}" data-local-key="chassisLoc">${cLocOpts}</select></td>`;
  cells += `<td>${escHtml(rec.Body_Class||'')}</td>`;
  const bLocOpts = LOC_OPTS.map(o => `<option value="${o}" ${o===bodyLoc?'selected':''}>${o||'—'}</option>`).join('');
  cells += `<td><select class="cell-sel" data-id="${id}" data-local-key="bodyLoc">${bLocOpts}</select></td>`;
  cells += `<td>${escHtml(rec.SERIAL_||'')}</td>`;
  cells += `<td>${fmtDate(rec.Actual_Build_Start_Date)}</td>`;
  cells += `<td><input type="date" class="cell-in" value="${escHtml(pdiDate)}" data-id="${id}" data-local-key="pdiDate"/></td>`;
  [
    { col:'Build_Started_',    val:!!rec.Build_Started_    },
    { col:'Build_Ended_',      val:!!rec.Build_Ended_      },
    { col:'Delivered_',        val:!!rec.Delivered_        },
    { col:'Paid_',             val:!!rec.Paid_             },
    { col:'Body_Picked_Up',    val:!!rec.Body_Picked_Up    },
    { col:'Chassis_Delivered', val:!!rec.Chassis_Delivered },
  ].forEach(({ col, val }) => {
    cells += `<td><div class="toggle-wrap"><label class="bool-toggle" data-row="${id}" data-col="${col}">
      <input type="checkbox" ${val?'checked':''}/><span class="slider"></span></label></div></td>`;
  });
  return `<tr class="data-row">${cells}</tr>`;
}

// ═══════════════════════════════════════════════════════════
// PAGE 3 — SALES & PIPELINE
// ═══════════════════════════════════════════════════════════

function renderSales() {
  populateChartYearDropdown();
  populateSalesYearDropdown();
  renderPipelineFunnel();
  renderMonthlyRevenueChart();
  renderBodyClassChart();
  renderMonthlySalesWIP();
}

function renderPipelineFunnel() {
  const container = document.getElementById('pipeline-funnel');
  if (!container) return;
  const stages = DB.pipelineCycle;
  if (!stages.length) { container.innerHTML = '<p style="color:var(--text-muted);padding:10px;">No pipeline data</p>'; return; }
  const stageColors = { 'Secured':'secured','Quoted':'quoted','Prospect':'prospect','Closed':'closed' };
  container.innerHTML = stages.map(s => `
    <div class="funnel-stage ${stageColors[s.Sales_Cycle]||'prospect'}">
      <div class="funnel-stage-label">${escHtml(s.Sales_Cycle||'Unknown')}</div>
      <div class="funnel-stage-count">${s.count||0}</div>
      <div class="funnel-stage-value">${fmtCurrency(s.Value,true)}</div>
    </div>`).join('');
}

// ── Chart Year dropdown ──
function populateChartYearDropdown() {
  const years = new Set();
  DB.monthSummary.forEach(r => {
    const d = parseDate(r.Sales_Month_Year);
    if (d) years.add(d.getFullYear());
  });
  const el = document.getElementById('chart-year-sel');
  if (!el) return;
  const sorted = [...years].sort((a,b) => b - a);
  el.innerHTML = `<option value="ALL">— All Years —</option>` +
    sorted.map(y => `<option value="${y}" ${y == chartYear ? 'selected' : ''}>${y}</option>`).join('');
  el.value = chartYear;
}

function renderMonthlyRevenueChart() {
  destroyChart('monthly-revenue');
  let data = DB.monthSummary.filter(r => r.Sales_Month_Year);
  // Apply chart year filter
  if (chartYear !== 'ALL') {
    data = data.filter(r => {
      const d = parseDate(r.Sales_Month_Year);
      return d && d.getFullYear() === parseInt(chartYear);
    });
  }
  data = data.sort((a,b) => (parseDate(a.Sales_Month_Year)||0) - (parseDate(b.Sales_Month_Year)||0));
  if (!data.length) return;
  const labels = data.map(r => {
    const d = parseDate(r.Sales_Month_Year);
    return d ? new Intl.DateTimeFormat('en-US',{month:'short',year:'2-digit'}).format(d) : '';
  });
  const ctx = document.getElementById('chart-monthly-revenue');
  if (!ctx) return;
  charts['monthly-revenue'] = new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[
      { label:'Revenue', data:data.map(r=>r.TOTAL_RETAIL_SALE||0), backgroundColor:'rgba(30,58,95,0.75)', borderRadius:2 },
      { label:'Target',  data:data.map(r=>r.Sales_Target||0), type:'line', borderColor:'#e8a020', borderWidth:2, pointRadius:2, fill:false },
      { label:'Profit',  data:data.map(r=>r.TRUE_PROFIT||0), backgroundColor:'rgba(5,150,105,0.6)', borderRadius:2 },
    ]},
    options:{ responsive:true, maintainAspectRatio:true,
      plugins:{ legend:{ labels:{ font:{ family:'Barlow',size:10 }}}},
      scales:{ y:{ ticks:{ callback:v=>fmtCurrency(v,true), font:{family:'Barlow',size:9} }}, x:{ ticks:{ font:{family:'Barlow',size:9} }} }
    }
  });
}

function renderBodyClassChart() {
  destroyChart('body-class');
  // For body class chart, filter backlog by selected chart year
  let source = DB.backlog.filter(r => r.Date_of_Delivery);
  if (chartYear !== 'ALL') {
    source = source.filter(r => {
      const d = parseDate(r.Date_of_Delivery);
      return d && d.getFullYear() === parseInt(chartYear);
    });
  }
  // Aggregate by body class
  const byBody = {};
  source.forEach(r => {
    const bc = r.Body_Class || '(blank)';
    byBody[bc] = (byBody[bc] || 0) + (r.TOTAL_RETAIL_SALE || 0);
  });
  const entries = Object.entries(byBody).filter(([,v]) => v > 0);
  if (!entries.length) return;
  const ctx = document.getElementById('chart-body-class');
  if (!ctx) return;
  charts['body-class'] = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels: entries.map(([k]) => k),
      datasets:[{ data: entries.map(([,v]) => v),
        backgroundColor:['#1e3a5f','#1a5c38','#7a5200','#0e5c4a','#4a1a80','#0891b2','#e8a020'],
        borderWidth:2, borderColor:'#fff' }]
    },
    options:{ responsive:true, maintainAspectRatio:true,
      plugins:{
        legend:{ position:'right', labels:{ font:{family:'Barlow',size:10} }},
        tooltip:{ callbacks:{ label: ctx => ctx.label+': '+fmtCurrency(ctx.raw,true) }}
      }
    }
  });
}

// ── Sales table year dropdown ──
function populateSalesYearDropdown() {
  const years = new Set();
  DB.backlog.forEach(r => {
    const d1 = parseDate(r.Date_of_Delivery);
    if (d1) years.add(d1.getFullYear());
    const d2 = parseDate(r.Projected_Build_End_Date);
    if (d2) years.add(d2.getFullYear());
  });
  const el = document.getElementById('sales-year-sel');
  if (!el) return;
  const sorted = [...years].sort((a,b) => b - a);
  el.innerHTML = sorted.map(y =>
    `<option value="${y}" ${y === selectedYear ? 'selected' : ''}>${y}</option>`
  ).join('');
  el.value = String(selectedYear);
}

// ── Monthly Sales & WIP ──
const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];
function monthLabel(m, y) { return MONTH_NAMES[m-1] + ' ' + y; }

function inMonth(val, month, year) {
  const d = parseDate(val);
  if (!d) return false;
  return d.getMonth()+1 === month && d.getFullYear() === year;
}

function wasWipInMonth(rec, month, year) {
  const now          = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear  = now.getFullYear();
  const isCurrentMonth = (month === currentMonth && year === currentYear);

  if (isCurrentMonth) {
    // Live WIP — actively started and not yet ended
    return !!rec.Build_Started_ && !rec.Build_Ended_;
  }

  // Historical WIP — build was in progress at any point during this month
  // Build must have started on or before the last day of the month
  // AND ended on or after the first day of the month (or not ended at all)
  const firstOfMonth = new Date(year, month - 1, 1);
  const lastOfMonth  = new Date(year, month, 0, 23, 59, 59);

  const startDate = parseDate(rec.Actual_Build_Start_Date)
                 || parseDate(rec.Projected_Build_Start_Date)
                 || parseDate(rec.Manual_Override_Projected_Build_Start_Date);

  const endDate   = parseDate(rec.Actual_Build_End_Date)
                 || parseDate(rec.Projected_Build_End_Date);

  if (!startDate) return false;

  // Must have started on or before last day of selected month
  if (startDate > lastOfMonth) return false;

  // Must not have ended before the first day of selected month
  if (endDate && endDate < firstOfMonth) return false;

  return true;
}

function renderMonthlySalesWIP() {
  const container = document.getElementById('monthly-sales-wip-container');
  if (!container) return;
  const m = selectedMonth, y = selectedYear;
  const nextMonth = m === 12 ? 1     : m + 1;
  const nextYear  = m === 12 ? y + 1 : y;

  const sales     = DB.backlog.filter(r => inMonth(r.Date_of_Delivery, m, y));
  const wip       = DB.backlog.filter(r => wasWipInMonth(r, m, y));
  const projected = DB.backlog.filter(r => {
    if (r.Projected_Next_Month_Filter === true)  return true;
    if (r.Projected_Next_Month_Filter === false) return false;
    return inMonth(r.ETA_WK_SALE_MONTH_FINAL_SALE_DATE, nextMonth, nextYear);
  });

  const totalSalesRetail = sales.reduce((s,r) => s+(r.TOTAL_RETAIL_SALE||0), 0);
  const totalSalesProfit = sales.reduce((s,r) => s+(r.TRUE_PROFIT||0), 0);
  const totalSalesCost   = sales.reduce((s,r) => s+(r.Total_Costs||0), 0);
  const salesMargin      = totalSalesRetail > 0 ? totalSalesProfit/totalSalesRetail : 0;
  const totalWipCost     = wip.reduce((s,r) => s+(r.Total_Costs||0), 0);
  const totalWipRetail   = wip.reduce((s,r) => s+(r.TOTAL_RETAIL_SALE||0), 0);
  const totalProjRetail  = projected.reduce((s,r) => s+(r.TOTAL_RETAIL_SALE||0), 0);
  const totalProjProfit  = projected.reduce((s,r) => s+(r.TRUE_PROFIT||0), 0);
  const totalProjCost    = projected.reduce((s,r) => s+(r.Total_Costs||0), 0);

  const kpiHtml = `<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid var(--border);border-radius:6px;padding:12px 14px;border-top:3px solid var(--forsale-hdr);">
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);margin-bottom:4px;">Sales This Month</div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:700;color:var(--forsale-hdr);line-height:1;">${sales.length}</div>
      <div style="font-size:11px;color:var(--text-dim);margin-top:2px;">${fmtCurrency(totalSalesRetail,true)} retail</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid var(--border);border-radius:6px;padding:12px 14px;border-top:3px solid ${totalSalesProfit>=0?'#059669':'#dc2626'};">
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);margin-bottom:4px;">Month Profit</div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:700;color:${totalSalesProfit>=0?'#059669':'#dc2626'};line-height:1;">${fmtCurrency(totalSalesProfit,true)}</div>
      <div style="font-size:11px;color:var(--text-dim);margin-top:2px;">Margin: ${fmtPct(salesMargin)}</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid var(--border);border-radius:6px;padding:12px 14px;border-top:3px solid #1e3a5f;">
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);margin-bottom:4px;">WIP (Active Builds)</div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:700;color:#1e3a5f;line-height:1;">${wip.length}</div>
      <div style="font-size:11px;color:var(--text-dim);margin-top:2px;">${fmtCurrency(totalWipCost,true)} total cost</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid var(--border);border-radius:6px;padding:12px 14px;border-top:3px solid #7c3aed;">
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);margin-bottom:4px;">Projected Next Month</div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:700;color:#7c3aed;line-height:1;">${projected.length}</div>
      <div style="font-size:11px;color:var(--text-dim);margin-top:2px;">${fmtCurrency(totalProjRetail,true)} retail</div>
    </div>
  </div>`;

  const thS = "padding:6px 8px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#ffffff;border-bottom:1px solid rgba(255,255,255,0.2);white-space:nowrap;";
  const tblStyle = 'width:100%;border-collapse:collapse;font-size:11.5px;min-width:700px;box-shadow:0 1px 4px rgba(0,0,0,0.07);border-radius:4px;overflow:hidden;';

  function secHdr(label, count, valueStr, bg, subBg, colHeaders) {
    // Default column headers — used by Sales and Projected tables
   const defaultHeaders = `
     <th style="${thS}">OP #</th><th style="${thS}">SO #</th><th style="${thS}">Customer</th>
     <th style="${thS}">Chassis / Model</th><th style="${thS}">Body</th><th style="${thS}">Delivery Date</th>
     <th style="${thS};text-align:right;">Retail Sale</th><th style="${thS};text-align:right;">Total Cost</th>
     <th style="${thS};text-align:right;">True Profit</th><th style="${thS};text-align:right;">Margin %</th>`;

   const headers = colHeaders || defaultHeaders;

   return `<tr><td colspan="20" style="padding:7px 10px;background:${bg};font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:#fff;">
       ${escHtml(label)}<span style="display:inline-block;margin-left:8px;background:rgba(255,255,255,0.2);border-radius:10px;padding:1px 8px;font-size:10px;">${count}</span>
       ${valueStr?`<span style="float:right;font-size:11px;font-weight:400;opacity:0.85;">${valueStr}</span>`:''}
     </td></tr>
     <tr style="background:${subBg};">${headers}</tr>`;
 }

  function dataRow(r, oddBg, evenBg, idx) {
    const bg = idx%2===0?oddBg:evenBg;
    const profit = r.TRUE_PROFIT||0, retail = r.TOTAL_RETAIL_SALE||0, cost = r.Total_Costs||0;
    const margin = retail>0?((profit/retail)*100).toFixed(1)+'%':'—';
    const pc = profit>=0?'#059669':'#dc2626';
    return `<tr style="background:${bg};" onmouseover="this.style.filter='brightness(0.96)'" onmouseout="this.style.filter=''">
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.OP_||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.SO_||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.NAME||'')}">${escHtml(r.NAME||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.CHASSIS_MODEL||'')}">${escHtml(r.CHASSIS_MODEL||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.Body_Class||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${fmtDate(r.Date_of_Delivery||r.ETA_WK_SALE_MONTH_FINAL_SALE_DATE)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;white-space:nowrap;">${fmtCurrency(retail)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;white-space:nowrap;">${fmtCurrency(cost)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:600;color:${pc};white-space:nowrap;">${fmtCurrency(profit)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;color:${pc};white-space:nowrap;">${margin}</td>
    </tr>`;
  }

  function wipRow(r, idx) {
    const bg = idx%2===0?'#f5f7fb':'#edf0f7';
    const cost = r.Total_Costs||0, retail = r.TOTAL_RETAIL_SALE||0;
    const start = parseDate(r.Actual_Build_Start_Date)||parseDate(r.Projected_Build_Start_Date)||parseDate(r.Manual_Override_Projected_Build_Start_Date);
    const end   = parseDate(r.Actual_Build_End_Date)||parseDate(r.Projected_Build_End_Date);
    return `<tr style="background:${bg};" onmouseover="this.style.filter='brightness(0.96)'" onmouseout="this.style.filter=''">
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.OP_||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.SO_||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.NAME||'')}">${escHtml(r.NAME||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.CHASSIS_MODEL||'')}">${escHtml(r.CHASSIS_MODEL||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.Body_Class||'')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${start?fmtDate(start)+' → '+(end?fmtDate(end):'TBD'):'—'}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;white-space:nowrap;">${fmtCurrency(retail)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;font-family:'Barlow Condensed',sans-serif;white-space:nowrap;">${fmtCurrency(cost)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid var(--border);text-align:right;color:var(--text-muted);font-size:10px;" colspan="2">In Progress</td>
    </tr>`;
  }

  function totalsRow(label, retail, cost, profit, bg) {
    const margin = retail>0?((profit/retail)*100).toFixed(1)+'%':'—';
    const pc = profit>=0?'#6ee7b7':'#fca5a5';
    return `<tr style="background:${bg};">
      <td colspan="6" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;letter-spacing:0.06em;">${label}</td>
      <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;">${fmtCurrency(retail)}</td>
      <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;">${fmtCurrency(cost)}</td>
      <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:${pc};white-space:nowrap;">${fmtCurrency(profit)}</td>
      <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:${pc};white-space:nowrap;">${margin}</td>
    </tr>`;
  }

  const emptyRow = msg => `<tr><td colspan="10" style="padding:14px;text-align:center;font-style:italic;color:var(--text-muted);background:#fafafa;">${msg}</td></tr>`;

  const salesRows = [...sales].sort((a,b)=>(parseDate(a.Date_of_Delivery)||0)-(parseDate(b.Date_of_Delivery)||0))
    .map((r,i)=>dataRow(r,'#f5fff8','#e8f5ec',i)).join('') || emptyRow('No deliveries recorded for this month');

  const wipRows = [...wip].sort((a,b)=>{
    const da=parseDate(a.Actual_Build_Start_Date||a.Projected_Build_Start_Date)||new Date(0);
    const db=parseDate(b.Actual_Build_Start_Date||b.Projected_Build_Start_Date)||new Date(0);
    return da-db;
  }).map((r,i)=>wipRow(r,i)).join('') || emptyRow('No active WIP builds for this month');

  const projRows = [...projected].sort((a,b)=>{
    const da=parseDate(a.ETA_WK_SALE_MONTH_FINAL_SALE_DATE||a.Projected_Date_of_Delivery)||new Date(0);
    const db=parseDate(b.ETA_WK_SALE_MONTH_FINAL_SALE_DATE||b.Projected_Date_of_Delivery)||new Date(0);
    return da-db;
  }).map((r,i)=>dataRow(r,'#f3eeff','#e8deff',i)).join('') || emptyRow(`No projected sales for ${monthLabel(nextMonth,nextYear)}`);

  container.innerHTML = kpiHtml + `
    <div style="overflow-x:auto;margin-bottom:16px;"><table style="${tblStyle}"><tbody>
      ${secHdr(`Sales — ${monthLabel(m,y)}`,sales.length,`${fmtCurrency(totalSalesRetail,true)} retail · ${fmtCurrency(totalSalesProfit,true)} profit`,'var(--forsale-hdr)','#236b45')}
      ${salesRows}
      ${sales.length?totalsRow('TOTAL SALES',totalSalesRetail,totalSalesCost,totalSalesProfit,'var(--forsale-hdr)'):''}
    </tbody></table></div>
    <div style="overflow-x:auto;margin-bottom:16px;"><table style="${tblStyle}"><tbody>
      ${secHdr(`WIP — ${monthLabel(m,y)}`, wip.length, `${fmtCurrency(totalWipCost,true)} cost · ${fmtCurrency(totalWipRetail,true)} retail`, 'var(--inservice-hdr)', '#2a4f7a', `
        <th style="${thS}">OP #</th><th style="${thS}">SO #</th><th style="${thS}">Customer</th>
        <th style="${thS}">Chassis / Model</th><th style="${thS}">Body</th><th style="${thS}">Build Dates</th>
        <th style="${thS};text-align:right;">Retail Sale</th><th style="${thS};text-align:right;">Total Cost</th>
        <th style="${thS};text-align:right;"></th><th style="${thS};text-align:right;"></th>`)}
      ${wipRows}
      ${wip.length?`<tr style="background:var(--inservice-hdr);">
        <td colspan="6" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;">TOTAL WIP</td>
        <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;">${fmtCurrency(totalWipRetail)}</td>
        <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;">${fmtCurrency(totalWipCost)}</td>
        <td colspan="2" style="padding:7px 10px;"></td>
      </tr>`:''}
    </tbody></table></div>
    <div style="overflow-x:auto;margin-bottom:16px;"><table style="${tblStyle}"><tbody>
      ${secHdr(`Projected Sales — ${monthLabel(nextMonth,nextYear)}`,projected.length,`${fmtCurrency(totalProjRetail,true)} retail · ${fmtCurrency(totalProjProfit,true)} projected profit`,'#4a1a80','#5c2898')}
      ${projRows}
      ${projected.length?totalsRow('TOTAL PROJECTED',totalProjRetail,totalProjCost,totalProjProfit,'#4a1a80'):''}
    </tbody></table></div>`;
}

// ═══════════════════════════════════════════════════════════
// PAGE 4 — CASH FLOW
// ═══════════════════════════════════════════════════════════

function renderCashFlow() {
  const now = new Date(); now.setHours(0,0,0,0);
  const horizon = new Date(now); horizon.setMonth(horizon.getMonth() + cfHorizon);

  // ── Generate cash flow events from NEWS_Build_Backlog ──
  //   Each OP produces up to 3 events:
  //   Body Delivery  — outflow — Body_Cost       — Body_Pickup_Date || NRC_EDD
  //   Chassis Payment— outflow — CHASSIS_COST    — Chassis_Delivery_Date || Chassis_EDD
  //   Sale           — inflow  — TOTAL_RETAIL_SALE — Date_of_Delivery || Projected_Date_of_Delivery

  function buildEvents(rec) {
  const events = [];
  const name = String(rec.NAME || rec.OP_ || '');

  const bodyDate    = parseDate(rec.Body_Pickup_Date) || parseDate(rec.NRC_EDD);
  const saleDate    = parseDate(rec.Date_of_Delivery) || parseDate(rec.Projected_Date_of_Delivery);

  // Chassis payment goes out on the same date as the sale
  const chassisDate = saleDate;

  if (bodyDate && rec.Body_Cost) {
    events.push({
      op:     String(rec.OP_ || ''),
      name,
      type:   'Body_Delivery',
      date:   bodyDate,
      amount: -(rec.Body_Cost || 0),
    });
  }
  if (chassisDate && rec.CHASSIS_COST) {
    events.push({
      op:     String(rec.OP_ || ''),
      name,
      type:   'Chassis_Payment',
      date:   chassisDate,
      amount: -(rec.CHASSIS_COST || 0),
    });
  }
  if (saleDate && rec.TOTAL_RETAIL_SALE) {
    events.push({
      op:     String(rec.OP_ || ''),
      name,
      type:   'Sale',
      date:   saleDate,
      amount: rec.TOTAL_RETAIL_SALE || 0,
    });
  }
  return events;
}

  const testEvents = DB.backlog.flatMap(buildEvents);
   console.log('Total generated events:', testEvents.length);
   console.log('Sample backlog rec:', DB.backlog[0]);
   console.log('Sample events from rec 0:', buildEvents(DB.backlog[0]));

  const allEvents = DB.backlog.flatMap(buildEvents)
    .sort((a, b) => a.date - b.date);

  // Future events only for charts (today → horizon)
  const futureEvents = allEvents.filter(r => r.date >= now && r.date <= horizon);

  // ── Period buckets for bar/line charts ──
  const buckets = {};
  futureEvents.forEach(r => {
    let key;
    if (cfPeriod === 'week') {
      const wk = new Date(r.date); wk.setDate(wk.getDate() - wk.getDay());
      key = wk.toISOString().slice(0,10);
    } else {
      key = r.date.getFullYear() + '-' + String(r.date.getMonth()+1).padStart(2,'0');
    }
    if (!buckets[key]) buckets[key] = { inflow:0, outflow:0 };
    if (r.amount > 0) buckets[key].inflow  += r.amount;
    else              buckets[key].outflow += r.amount;
  });

  const periods = Object.keys(buckets).sort();
  const labels  = periods.map(p => {
    if (cfPeriod === 'week') {
      const d = new Date(p+'T00:00:00');
      return 'Wk '+(d.getMonth()+1)+'/'+d.getDate();
    }
    const [y,m] = p.split('-');
    return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(m)-1]+' '+y.slice(2);
  });

  const cfPeriodLabel = document.getElementById('cf-period-label');
  if (cfPeriodLabel) cfPeriodLabel.textContent = cfPeriod === 'week' ? 'Weekly' : 'Monthly';

 // ── Inflow / Outflow chart with net labels ──
  destroyChart('cashflow');
  destroyChart('cashflow-cumulative');
  const ctx1 = document.getElementById('chart-cashflow');
  if (ctx1 && periods.length) {
    const netData = periods.map(p => buckets[p].inflow + buckets[p].outflow);
    charts['cashflow'] = new Chart(ctx1, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label:           'Inflow',
            data:            periods.map(p => buckets[p].inflow),
            backgroundColor: 'rgba(5,150,105,0.7)',
            borderRadius:    2,
            datalabels:      { display: false },
          },
          {
            label:           'Outflow',
            data:            periods.map(p => buckets[p].outflow),
            backgroundColor: 'rgba(220,38,38,0.6)',
            borderRadius:    2,
            // Net label sits above the inflow bar (tallest positive bar)
            datalabels: {
              display: true,
              align:   'end',
              anchor:  'end',
              formatter: (_, ctx) => {
                const net = netData[ctx.dataIndex];
                const abs = fmtCurrency(Math.abs(net), true);
                return net >= 0 ? '+' + abs : '-' + abs;
              },
              color: idx => netData[idx.dataIndex] >= 0 ? '#059669' : '#dc2626',
              font:  { family:'Barlow Condensed', size:9, weight:'700' },
              offset: 2,
            },
          },
        ]
      },
      options: {
        responsive:          true,
        maintainAspectRatio: true,
        layout: { padding: { top: 20 } }, // room for net labels
        plugins: {
          legend:     { labels: { font: { family:'Barlow', size:10 } } },
          tooltip:    { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmtCurrency(ctx.raw, true) } },
          datalabels: { display: false }, // default off, overridden per dataset above
        },
        scales: {
          y: { ticks: { callback: v => fmtCurrency(v, true), font: { family:'Barlow', size:9 } } },
          x: { ticks: { font: { family:'Barlow', size:9 } } }
        }
      },
      plugins: [ChartDataLabels],
    });
  }

  // ── Week navigation events table ──
  if (!cfSelectedWeekStart) {
    cfSelectedWeekStart = new Date(now);
    const day  = cfSelectedWeekStart.getDay();
    const diff = day === 0 ? -6 : 1 - day; // snap to Monday
    cfSelectedWeekStart.setDate(cfSelectedWeekStart.getDate() + diff);
  }

  const weekStart = new Date(cfSelectedWeekStart);
  const weekEnd   = new Date(weekStart); weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23,59,59,999);

  function fmtWeekLabel(d) {
    return new Intl.DateTimeFormat('en-US', { month:'short', day:'numeric' }).format(d);
  }
  const weekLabel = `${fmtWeekLabel(weekStart)} – ${fmtWeekLabel(weekEnd)}, ${weekEnd.getFullYear()}`;

  // All events (past and future) filtered to selected week
  const weekEvents = allEvents.filter(r => r.date >= weekStart && r.date <= weekEnd);

  const weekInflow  = weekEvents.filter(r => r.amount > 0).reduce((s,r) => s + r.amount, 0);
  const weekOutflow = weekEvents.filter(r => r.amount < 0).reduce((s,r) => s + r.amount, 0);
  const weekNet     = weekInflow + weekOutflow;
  const netColor    = weekNet >= 0 ? '#059669' : '#dc2626';

  const eventColors = {
    'Sale':            '#059669',
    'Body_Delivery':   '#dc2626',
    'Chassis_Payment': '#d97706',
  };

  const eventsHtml = weekEvents.length
    ? weekEvents.map(r => {
        const color = r.amount >= 0 ? '#059669' : '#dc2626';
        const ec    = eventColors[r.type] || '#6b7280';
        const label = r.type === 'Body_Delivery'   ? 'Body Delivery'
                    : r.type === 'Chassis_Payment'  ? 'Chassis Payment'
                    : 'Sale';
        return `<tr class="data-row">
          <td>${escHtml(r.op)}</td>
          <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.name)}">${escHtml(r.name)}</td>
          <td><span style="font-size:10px;padding:1px 6px;border-radius:2px;background:${ec}22;color:${ec};font-family:'Barlow Condensed',sans-serif;letter-spacing:0.06em;">${label}</span></td>
          <td>${fmtDate(r.date)}</td>
          <td style="color:${color};font-family:'Barlow Condensed',sans-serif;font-weight:600;">${fmtCurrency(r.amount)}</td>
        </tr>`;
      }).join('')
    : `<tr class="empty-row"><td colspan="5">No events for this week</td></tr>`;

  const totalsHtml = weekEvents.length
    ? `<tr style="background:#f0f4f8;border-top:2px solid var(--border);">
        <td colspan="4" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:700;letter-spacing:0.08em;color:var(--text-dim);">
          INFLOW &nbsp;<span style="color:#059669;font-weight:700;">${fmtCurrency(weekInflow)}</span>
          &nbsp;&nbsp; OUTFLOW &nbsp;<span style="color:#dc2626;font-weight:700;">${fmtCurrency(weekOutflow)}</span>
        </td>
        <td style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:${netColor};text-align:right;">
          NET ${fmtCurrency(weekNet)}
        </td>
      </tr>`
    : '';

  const eventsContainer = document.getElementById('cf-events-container');
  if (eventsContainer) {
    eventsContainer.innerHTML = `
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;flex-wrap:wrap;">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);">Cash Flow Events</div>
        <div style="display:flex;align-items:center;gap:0;border:1px solid var(--border);border-radius:3px;overflow:hidden;margin-left:auto;">
          <button id="cf-week-prev" style="padding:5px 14px;background:var(--surface);border:none;border-right:1px solid var(--border);color:var(--text-dim);font-size:13px;cursor:pointer;">◀</button>
          <span style="padding:5px 16px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:600;color:var(--text);background:var(--white);white-space:nowrap;">${weekLabel}</span>
          <button id="cf-week-next" style="padding:5px 14px;background:var(--surface);border:none;border-left:1px solid var(--border);color:var(--text-dim);font-size:13px;cursor:pointer;">▶</button>
        </div>
        <button id="cf-week-today" style="padding:5px 10px;background:var(--surface);border:1px solid var(--border);border-radius:3px;color:var(--text-dim);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.08em;text-transform:uppercase;cursor:pointer;">This Week</button>
      </div>
      <div style="overflow-x:auto;">
        <table class="section-table sec-inservice" style="min-width:700px;">
          <thead>
            <tr class="col-hdr">
              <th>OP #</th><th>Customer</th><th>Event</th><th>Date</th><th>Amount</th>
            </tr>
          </thead>
          <tbody>${eventsHtml}</tbody>
          <tfoot>${totalsHtml}</tfoot>
        </table>
      </div>`;

    document.getElementById('cf-week-prev').addEventListener('click', () => {
      cfSelectedWeekStart.setDate(cfSelectedWeekStart.getDate() - 7);
      renderCashFlow();
    });
    document.getElementById('cf-week-next').addEventListener('click', () => {
      cfSelectedWeekStart.setDate(cfSelectedWeekStart.getDate() + 7);
      renderCashFlow();
    });
    document.getElementById('cf-week-today').addEventListener('click', () => {
      cfSelectedWeekStart = null;
      renderCashFlow();
    });
  }
}

// ═══════════════════════════════════════════════════════════
// PAGE 2 — GANTT
// ═══════════════════════════════════════════════════════════

const GANTT_ZOOM = { day:{cellWidth:40}, week:{cellWidth:20}, month:{cellWidth:8} };
const SECS_PER_DAY = 86400;

function ganttToDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val)?null:val;
  if (Array.isArray(val) && val[0]==='d') return new Date(val[1]*1000);
  const n = Number(val);
  if (!isNaN(n)&&n!==0) {
    if (n<100000)       return new Date(n*SECS_PER_DAY*1000);
    if (n<100000000000) return new Date(n*1000);
    return new Date(n);
  }
  if (typeof val==='string'&&val.trim()) { const d=new Date(val.trim()); return isNaN(d)?null:d; }
  return null;
}
function ganttDaysDiff(d1,d2) { return Math.round((d2.getTime()-d1.getTime())/86400000); }
function ganttAddDays(date,n)  { const r=new Date(date); r.setDate(r.getDate()+n); return r; }
function ganttFmtDate(d)       { if(!d) return '—'; return new Intl.DateTimeFormat('en-US',{month:'short',day:'numeric',year:'numeric'}).format(d); }
function ganttIsWeekend(d)     { const day=d.getDay(); return day===0||day===6; }
function ganttIsSameDay(a,b)   { return a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate(); }

function resolveGanttDates(r) {
  const tailStart = ganttToDate(r.Body_Pickup_Date)||ganttToDate(r.NRC_EDD)||null;
  const barStart  = ganttToDate(r.Actual_Build_Start_Date)||ganttToDate(r.Projected_Build_Start_Date)||ganttToDate(r.Manual_Override_Projected_Build_Start_Date)||null;
  const barEnd    = ganttToDate(r.Actual_Build_End_Date)||ganttToDate(r.Projected_Build_End_Date)||null;
  return { tailStart, barStart, barEnd };
}

function buildGanttTasks(range) {
  const today=new Date(); today.setHours(0,0,0,0);
  const in3weeks=ganttAddDays(today,21), in1year=ganttAddDays(today,365);
  let records;
  if (range==='short') {
    records=DB.backlog.filter(r=>{
      const {barStart}=resolveGanttDates(r);
      return (r.Build_Started_&&!r.Build_Ended_)||(barStart&&barStart>=ganttAddDays(today,-1)&&barStart<=in3weeks);
    });
  } else {
    records=DB.backlog.filter(r=>{
      if(r.Delivered_) return false;
      const {tailStart,barStart,barEnd}=resolveGanttDates(r);
      const earliest=tailStart||barStart, latest=barEnd||barStart;
      if(!earliest&&!latest) return false;
      if(latest&&latest>in1year) return false;
      return true;
    });
  }
  return records.map(r=>{
    const {tailStart,barStart,barEnd}=resolveGanttDates(r);
    let tier=4;
    if(r.Build_Ended_) tier=1;
    else if(r.Build_Started_) tier=2;
    else if(barStart&&barStart<=in3weeks) tier=3;
    return { id:r.id, titre:r.NAME||r.OP_||'—', op:String(r.OP_||''),
      bodyClass:String(r.Body_Class||''), chassis:String(r.CHASSIS_MODEL||''),
      tailStart, barStart, barEnd, tier, overflow:!!(r.Overflow||r.Overflow_Check),
      tailSource:r.Body_Pickup_Date?'Body Pickup':(r.NRC_EDD?'NRC EDD':null),
      startSource:r.Actual_Build_Start_Date?'Actual':(r.Projected_Build_Start_Date?'Projected':'Manual Override'),
      endSource:r.Actual_Build_End_Date?'Actual':'Projected' };
  }).filter(t=>t.barStart||t.barEnd);
}

function sortGanttTasks() {
  if (ganttSortMode==='date') ganttTasks.sort((a,b)=>((a.barStart||a.barEnd||a.tailStart||new Date(0)).getTime())-((b.barStart||b.barEnd||b.tailStart||new Date(0)).getTime()));
  else if (ganttSortMode==='name') ganttTasks.sort((a,b)=>a.titre.localeCompare(b.titre));
  else if (ganttSortMode==='status') ganttTasks.sort((a,b)=>a.tier-b.tier||((a.barStart||new Date(0))-(b.barStart||new Date(0))));
}

function calcGanttRange() {
  if (!ganttTasks.length) { ganttStart=ganttAddDays(new Date(),-7); ganttDays=60; return; }
  let minD=null, maxD=null;
  ganttTasks.forEach(t=>[t.tailStart,t.barStart,t.barEnd].forEach(d=>{
    if(!d) return;
    if(!minD||d<minD) minD=d;
    if(!maxD||d>maxD) maxD=d;
  }));
  if(!minD) minD=new Date();
  if(!maxD) maxD=ganttAddDays(minD,30);
  ganttStart=ganttAddDays(minD,-7);
  ganttDays=ganttDaysDiff(ganttStart,ganttAddDays(maxD,14));
  if(ganttDays<30) ganttDays=30;
}

const GANTT_TIER = {
  1:{bar:'linear-gradient(135deg,#10b981,#34d399)',border:'#10b981',label:'Complete'},
  2:{bar:'linear-gradient(135deg,#3b82f6,#60a5fa)',border:'#3b82f6',label:'In Progress'},
  3:{bar:'linear-gradient(135deg,#f59e0b,#fbbf24)',border:'#f59e0b',label:'Starting Soon'},
  4:{bar:'linear-gradient(135deg,#6b7280,#9ca3af)',border:'#6b7280',label:'Upcoming'},
};

function renderGantt() {
  const container=document.getElementById('gantt-container');
  if(!container) return;
  ganttTasks=buildGanttTasks(ganttRange);
  sortGanttTasks(); calcGanttRange();
  ganttCellWidth=GANTT_ZOOM[ganttZoom].cellWidth;
  const totalWidth=ganttDays*ganttCellWidth;
  const today=new Date(); today.setHours(0,0,0,0);
  const ROW_H=50, todayOff=ganttDaysDiff(ganttStart,today);

  let monthHtml='',curMonth='',curMonthW=0;
  const monthCell=(label,w)=>`<div style="width:${w}px;min-width:${w}px;padding:0 6px;font-size:10px;font-family:'Barlow Condensed',sans-serif;letter-spacing:0.06em;color:var(--text-muted);border-right:1px solid var(--border);display:flex;align-items:center;justify-content:center;white-space:nowrap;overflow:hidden;box-sizing:border-box;flex-shrink:0;">${label}</div>`;
  for(let i=0;i<ganttDays;i++){
    const d=ganttAddDays(ganttStart,i);
    const label=new Intl.DateTimeFormat('en-US',{month:'long',year:'numeric'}).format(d);
    if(label!==curMonth){if(curMonth) monthHtml+=monthCell(curMonth,curMonthW); curMonth=label; curMonthW=0;}
    curMonthW+=ganttCellWidth;
  }
  if(curMonth) monthHtml+=monthCell(curMonth,curMonthW);

  let dayHtml='';
  for(let i=0;i<ganttDays;i++){
    const d=ganttAddDays(ganttStart,i);
    const isToday=ganttIsSameDay(d,today), weekend=ganttIsWeekend(d);
    let label='';
    if(ganttZoom==='day') label=String(d.getDate());
    else if(ganttZoom==='week') label=d.getDay()===1?(d.getMonth()+1)+'/'+d.getDate():'';
    else label=d.getDate()===1?new Intl.DateTimeFormat('en-US',{month:'short'}).format(d):'';
    const bg=isToday?'#dbeafe':weekend?'#fef2f2':'var(--surface)';
    const color=isToday?'#2563eb':'var(--text-muted)', fw=isToday?'700':'normal';
    dayHtml+=`<div style="min-width:${ganttCellWidth}px;width:${ganttCellWidth}px;height:24px;display:flex;align-items:center;justify-content:center;font-size:10px;font-family:'Barlow',sans-serif;color:${color};font-weight:${fw};background:${bg};border-right:1px solid var(--border);box-sizing:border-box;flex-shrink:0;">${label}</div>`;
  }

  let taskListHtml='';
  if(!ganttTasks.length) {
    taskListHtml=`<div style="display:flex;align-items:center;justify-content:center;padding:40px 20px;color:var(--text-muted);font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:0.1em;text-align:center;">NO BUILDS MATCH THIS VIEW</div>`;
  } else {
    ganttTasks.forEach(t=>{
      const tc=GANTT_TIER[t.tier], sel=t.id===ganttSelected;
      taskListHtml+=`<div class="g-task-row" data-id="${t.id}" style="height:${ROW_H}px;display:flex;align-items:center;padding:0 10px;gap:8px;border-bottom:1px solid var(--border);cursor:pointer;background:${sel?'#eef2ff':'#fff'};outline:${sel?'2px solid #4f46e5':''};outline-offset:-2px;transition:background 0.12s;box-sizing:border-box;">
        <div style="width:4px;height:28px;border-radius:2px;background:${tc.border};flex-shrink:0;"></div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:12px;font-weight:600;font-family:'Barlow',sans-serif;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--text);" title="${escHtml(t.titre)}">${escHtml(t.titre)}</div>
          ${t.op&&t.op!==t.titre?`<div style="font-size:10px;font-family:'Barlow Condensed',sans-serif;color:var(--text-muted);letter-spacing:0.06em;">OP# ${escHtml(t.op)}</div>`:''}
          <div style="font-size:10px;color:var(--text-muted);">${t.barStart?ganttFmtDate(t.barStart):'—'} → ${t.barEnd?ganttFmtDate(t.barEnd):'TBD'}</div>
        </div>
        ${t.overflow?`<div style="width:8px;height:8px;border-radius:50%;background:#ef4444;flex-shrink:0;" title="Overflow"></div>`:''}
      </div>`;
    });
  }

  let gridHtml='';
  ganttTasks.forEach((t,idx)=>{
    let cells='';
    for(let i=0;i<ganttDays;i++){
      const d=ganttAddDays(ganttStart,i);
      cells+=`<div style="min-width:${ganttCellWidth}px;width:${ganttCellWidth}px;height:100%;border-right:1px solid ${ganttIsWeekend(d)?'#fee2e2':'#f1f5f9'};background:${ganttIsWeekend(d)?'rgba(254,242,242,0.5)':'transparent'};box-sizing:border-box;flex-shrink:0;"></div>`;
    }
    gridHtml+=`<div style="height:${ROW_H}px;display:flex;border-bottom:1px solid var(--border);position:relative;box-sizing:border-box;">${cells}</div>`;
  });

  ganttTasks.forEach((t,idx)=>{
    const top=idx*ROW_H, tc=GANTT_TIER[t.tier], sel=t.id===ganttSelected;
    if(t.tailStart&&t.barStart&&t.tailStart<t.barStart){
      const tailOff=ganttDaysDiff(ganttStart,t.tailStart), tailDur=ganttDaysDiff(t.tailStart,t.barStart);
      if(tailOff+tailDur>0&&tailOff<ganttDays&&tailDur>0){
        const visStartPx=Math.max(0,tailOff*ganttCellWidth), visEndPx=Math.min(totalWidth,(tailOff+tailDur)*ganttCellWidth), tailWidth=visEndPx-visStartPx;
        if(tailWidth>0) gridHtml+=`<div class="g-bar" data-id="${t.id}" data-title="${escHtml(t.titre)} — ${escHtml(t.tailSource||'Tail')}" data-op="${escHtml(t.op)}" data-start="${ganttFmtDate(t.tailStart)}" data-end="${ganttFmtDate(t.barStart)}" data-body="${escHtml(t.bodyClass)}" data-chassis="${escHtml(t.chassis)}" data-tail="1" style="position:absolute;left:${visStartPx}px;width:${tailWidth}px;top:${top+19}px;height:12px;border-radius:4px;background:linear-gradient(135deg,#f97316,#fdba74);opacity:0.85;cursor:pointer;z-index:2;box-shadow:0 1px 2px rgba(0,0,0,0.15);"></div>`;
      }
    }
    if(t.barStart&&t.barEnd){
      const startOff=ganttDaysDiff(ganttStart,t.barStart), dur=ganttDaysDiff(t.barStart,t.barEnd)+1;
      if(startOff+dur>0&&startOff<ganttDays){
        const left=Math.max(0,startOff*ganttCellWidth)+2;
        const width=Math.min((startOff<0?dur+startOff:dur)*ganttCellWidth-4,totalWidth-left-4);
        if(width>0) gridHtml+=`<div class="g-bar" data-id="${t.id}" data-title="${escHtml(t.titre)}" data-op="${escHtml(t.op)}" data-start="${ganttFmtDate(t.barStart)}" data-end="${ganttFmtDate(t.barEnd)}" data-body="${escHtml(t.bodyClass)}" data-chassis="${escHtml(t.chassis)}" data-startsrc="${escHtml(t.startSource)}" data-endsrc="${escHtml(t.endSource)}" style="position:absolute;left:${left}px;width:${width}px;top:${top+13}px;height:24px;border-radius:6px;background:${tc.bar};display:flex;align-items:center;padding:0 8px;font-size:10px;color:#fff;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,0.2);${sel?'outline:3px solid #1e3a5f;outline-offset:1px;':''}z-index:3;">${width>60?escHtml(t.titre):''}</div>`;
      }
    } else if(t.barStart&&!t.barEnd){
      const startOff=ganttDaysDiff(ganttStart,t.barStart);
      if(startOff>=0&&startOff<ganttDays) gridHtml+=`<div class="g-bar" data-id="${t.id}" data-title="${escHtml(t.titre)} (no end date)" data-op="${escHtml(t.op)}" data-start="${ganttFmtDate(t.barStart)}" data-end="—" data-body="${escHtml(t.bodyClass)}" data-chassis="${escHtml(t.chassis)}" style="position:absolute;left:${startOff*ganttCellWidth+2}px;width:12px;top:${top+13}px;height:24px;border-radius:6px;background:${tc.bar};cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,0.2);z-index:3;opacity:0.7;"></div>`;
    }
  });

  if(todayOff>=0&&todayOff<ganttDays){
    const todayLeft=todayOff*ganttCellWidth+ganttCellWidth/2;
    gridHtml+=`<div style="position:absolute;left:${todayLeft}px;top:0;bottom:0;width:2px;background:#ef4444;z-index:5;pointer-events:none;"><div style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:9px;color:#ef4444;font-weight:700;font-family:'Barlow Condensed',sans-serif;letter-spacing:0.08em;white-space:nowrap;">TODAY</div></div>`;
  }

  container.innerHTML=`
    <div style="display:flex;align-items:center;gap:10px;padding:8px 0;flex-wrap:wrap;">
      <span style="font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.1em;color:var(--text-muted);text-transform:uppercase;">Zoom:</span>
      <div style="display:flex;border:1px solid var(--border);border-radius:3px;overflow:hidden;">
        ${['day','week','month'].map((z,i,arr)=>`<button onclick="setGanttZoom('${z}')" style="padding:4px 10px;background:${ganttZoom===z?'var(--navy)':'var(--surface)'};color:${ganttZoom===z?'#fff':'var(--text-dim)'};border:none;border-right:${i<arr.length-1?'1px solid var(--border)':'none'};font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.1em;text-transform:capitalize;cursor:pointer;">${z}</button>`).join('')}
      </div>
      <span style="font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.1em;color:var(--text-muted);text-transform:uppercase;margin-left:6px;">Sort:</span>
      <div style="display:flex;border:1px solid var(--border);border-radius:3px;overflow:hidden;">
        ${[['date','By Date'],['name','By Name'],['status','By Status']].map(([k,l],i,arr)=>`<button onclick="setGanttSort('${k}')" style="padding:4px 10px;background:${ganttSortMode===k?'var(--navy)':'var(--surface)'};color:${ganttSortMode===k?'#fff':'var(--text-dim)'};border:none;border-right:${i<arr.length-1?'1px solid var(--border)':'none'};font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.06em;cursor:pointer;">${l}</button>`).join('')}
      </div>
      <span style="font-family:'Barlow Condensed',sans-serif;font-size:10px;color:var(--text-muted);margin-left:auto;">${ganttTasks.length} build${ganttTasks.length!==1?'s':''}</span>
    </div>
    <div style="display:flex;gap:14px;padding-bottom:8px;flex-wrap:wrap;">
      ${Object.values(GANTT_TIER).map(tc=>`<div style="display:flex;align-items:center;gap:5px;font-family:'Barlow Condensed',sans-serif;font-size:10px;color:var(--text-muted);"><div style="width:14px;height:10px;border-radius:2px;background:${tc.border};"></div>${tc.label}</div>`).join('')}
      <div style="display:flex;align-items:center;gap:5px;font-family:'Barlow Condensed',sans-serif;font-size:10px;color:var(--text-muted);"><div style="width:14px;height:8px;border-radius:2px;background:#f97316;opacity:0.85;"></div>Body Pickup / NRC EDD</div>
    </div>
    <div style="display:flex;flex-direction:column;height:calc(100vh - 215px);min-height:300px;border:1px solid var(--border);border-radius:4px;overflow:hidden;background:#fff;">
      <div style="display:flex;flex-shrink:0;border-bottom:1px solid var(--border);">
        <div style="width:260px;min-width:260px;background:var(--bg);border-right:1px solid var(--border);padding:0 10px;display:flex;align-items:flex-end;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);">BUILD</div>
        <div id="g-month-header" style="flex:1;overflow:hidden;display:flex;height:24px;"><div style="display:flex;width:${totalWidth}px;flex-shrink:0;">${monthHtml}</div></div>
      </div>
      <div style="display:flex;flex-shrink:0;border-bottom:2px solid var(--border);">
        <div style="width:260px;min-width:260px;background:var(--bg);border-right:1px solid var(--border);"></div>
        <div id="g-day-header" style="flex:1;overflow:hidden;display:flex;height:24px;"><div style="display:flex;width:${totalWidth}px;flex-shrink:0;">${dayHtml}</div></div>
      </div>
      <div style="display:flex;flex:1;overflow:hidden;">
        <div id="g-task-list" style="width:260px;min-width:260px;overflow-y:auto;border-right:1px solid var(--border);background:#fff;">${taskListHtml}</div>
        <div id="g-timeline-body" style="flex:1;overflow:auto;position:relative;">
          <div id="g-timeline-grid" style="position:relative;width:${totalWidth}px;min-height:100%;">${gridHtml}</div>
        </div>
      </div>
    </div>
    <div id="g-tooltip" style="position:fixed;background:#1a2030;color:#fff;padding:10px 14px;border-radius:6px;font-size:11px;z-index:9999;max-width:300px;box-shadow:0 8px 24px rgba(0,0,0,0.25);pointer-events:none;opacity:0;transition:opacity 0.15s;display:none;">
      <div id="g-tt-title" style="font-weight:700;font-family:'Barlow Condensed',sans-serif;font-size:13px;letter-spacing:0.06em;margin-bottom:5px;"></div>
      <div id="g-tt-op" style="font-size:10px;color:rgba(255,255,255,0.55);margin-bottom:5px;font-family:'Barlow Condensed',sans-serif;letter-spacing:0.08em;"></div>
      <div style="display:grid;grid-template-columns:auto 1fr;gap:2px 10px;font-size:10px;opacity:0.9;">
        <span style="color:rgba(255,255,255,0.55);">Start:</span><span id="g-tt-start"></span>
        <span style="color:rgba(255,255,255,0.55);">End:</span><span id="g-tt-end"></span>
        <span id="g-tt-src-label" style="color:rgba(255,255,255,0.55);"></span><span id="g-tt-src"></span>
        <span style="color:rgba(255,255,255,0.55);">Body:</span><span id="g-tt-body"></span>
        <span style="color:rgba(255,255,255,0.55);">Chassis:</span><span id="g-tt-chassis"></span>
      </div>
    </div>`;

  const timelineBody=document.getElementById('g-timeline-body');
  const monthHdr=document.getElementById('g-month-header'), dayHdr=document.getElementById('g-day-header'), taskListEl=document.getElementById('g-task-list');
  if(timelineBody){
    timelineBody.addEventListener('scroll',()=>{
      if(monthHdr) monthHdr.scrollLeft=timelineBody.scrollLeft;
      if(dayHdr)   dayHdr.scrollLeft=timelineBody.scrollLeft;
      if(taskListEl) taskListEl.scrollTop=timelineBody.scrollTop;
    });
    setTimeout(()=>{ const px=todayOff*ganttCellWidth; if(px>0) timelineBody.scrollLeft=Math.max(0,px-timelineBody.clientWidth/3); },50);
  }
  document.querySelectorAll('.g-task-row').forEach(row=>{
    row.addEventListener('click',()=>{ ganttSelected=parseInt(row.dataset.id); renderGantt(); });
  });
  const tooltip=document.getElementById('g-tooltip');
  document.querySelectorAll('.g-bar').forEach(bar=>{
    bar.addEventListener('mouseenter',e=>{
      document.getElementById('g-tt-title').textContent=bar.dataset.title||'';
      document.getElementById('g-tt-op').textContent=bar.dataset.op?'OP# '+bar.dataset.op:'';
      document.getElementById('g-tt-start').textContent=bar.dataset.start||'—';
      document.getElementById('g-tt-end').textContent=bar.dataset.end||'—';
      document.getElementById('g-tt-body').textContent=bar.dataset.body||'—';
      document.getElementById('g-tt-chassis').textContent=bar.dataset.chassis||'—';
      if(bar.dataset.tail){ document.getElementById('g-tt-src-label').textContent='Type:'; document.getElementById('g-tt-src').textContent='NRC / Body Pickup Tail'; }
      else if(bar.dataset.startsrc){ document.getElementById('g-tt-src-label').textContent='Dates:'; document.getElementById('g-tt-src').textContent=`${bar.dataset.startsrc} start / ${bar.dataset.endsrc} end`; }
      else { document.getElementById('g-tt-src-label').textContent=''; document.getElementById('g-tt-src').textContent=''; }
      tooltip.style.display='block'; setTimeout(()=>tooltip.style.opacity='1',10);
    });
    bar.addEventListener('mousemove',e=>{ tooltip.style.left=(e.clientX+16)+'px'; tooltip.style.top=(e.clientY+16)+'px'; });
    bar.addEventListener('mouseleave',()=>{ tooltip.style.opacity='0'; setTimeout(()=>{ if(tooltip.style.opacity==='0') tooltip.style.display='none'; },160); });
    bar.addEventListener('click',()=>{ const id=parseInt(bar.dataset.id); if(id){ganttSelected=id;renderGantt();} });
  });
}
function setGanttZoom(z) { ganttZoom=z; renderGantt(); }
function setGanttSort(s) { ganttSortMode=s; renderGantt(); }

// ═══════════════════════════════════════════════════════════
// PAGE 5 — STOCK INVENTORY
// ═══════════════════════════════════════════════════════════

function renderInventory() {
  const snapInput = document.getElementById('inv-snap-date');
  if (snapInput) snapInput.value = toInputDate(invState.snapDate);

  const snapDate = invState.snapDate;

  // All stock records (NAME=STOCK, not delivered, not paid)
  const allStock = DB.backlog.filter(r => String(r.NAME||'').trim() === 'STOCK' && !r.Delivered_ && !r.Paid_);

  // Populate filter dropdowns from all stock
  const bodies  = new Set(allStock.map(r=>r.Body_Class).filter(Boolean).map(String));
  const chassis = new Set(allStock.map(r=>r.CHASSIS_MODEL).filter(Boolean).map(String));
  const colors  = new Set(allStock.map(r=>r.Color).filter(Boolean).map(String));
  rebuildSelect('inv-filter-body',    bodies,  invState.body,    '— Any —');
  rebuildSelect('inv-filter-chassis', chassis, invState.chassis, '— Any —');
  rebuildSelect('inv-filter-color',   colors,  invState.color,   '— Any —');

  // Apply user filters
  const filtered = allStock.filter(r => {
    if (invState.body    !== 'ANY' && String(r.Body_Class||'')    !== invState.body)    return false;
    if (invState.chassis !== 'ANY' && String(r.CHASSIS_MODEL||'') !== invState.chassis) return false;
    if (invState.color   !== 'ANY' && String(r.Color||'')         !== invState.color)   return false;
    return true;
  });

  // Available by snapshot date (NRC EDD <= snap date)
  // Available by snapshot date — Body_Pickup_Date if present, otherwise NRC_EDD
  const available = filtered.filter(r => {
   const arrivalDate = parseDate(r.Body_Pickup_Date) || parseDate(r.NRC_EDD);
   if (!arrivalDate) return false;
   arrivalDate.setHours(0,0,0,0);
   return arrivalDate <= snapDate;
  });

  // ── Inventory Value Chart ──
  renderInventoryValueChart(snapDate);
   
  // KPI
  const kpiEl = document.getElementById('inv-kpi');
   if (kpiEl) kpiEl.innerHTML = `
    <div class="kpi-card onlot" style="max-width:200px;">
      <div class="kpi-label">Bodies in Stock by Date</div>
      <div class="kpi-value" style="color:var(--kpi-onlot);">${available.length}</div>
      <div class="kpi-sub">of ${filtered.length} stock builds</div>
    </div>`;

 // ── Grouped detail table ──
  const area = document.getElementById('inv-table-area');
  if (!area) return;

  if (!available.length) {
    area.innerHTML = '<p style="color:var(--text-muted);font-family:Barlow Condensed,sans-serif;letter-spacing:0.1em;padding:20px;text-align:center;">NO STOCK AVAILABLE BY SELECTED DATE</p>';
    return;
  }

  function daysInInventory(rec) {
    const arrivalDate = parseDate(rec.Body_Pickup_Date) || parseDate(rec.NRC_EDD);
    if (!arrivalDate) return null;
    const diff = Math.floor((snapDate - arrivalDate) / 86400000);
    return diff >= 0 ? diff : null;
  }

  function daysChip(days) {
    if (days === null) return '—';
    const color = days <= 30 ? '#059669' : days <= 60 ? '#d97706' : '#dc2626';
    return `<span style="font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:${color};">${days}d</span>`;
  }

  function arrivalCell(rec) {
    const d   = parseDate(rec.Body_Pickup_Date) || parseDate(rec.NRC_EDD);
    const src = rec.Body_Pickup_Date ? 'Pickup' : 'NRC EDD';
    return d ? `${fmtDate(d)} <span style="font-size:9px;color:var(--text-muted);font-family:'Barlow Condensed',sans-serif;">${src}</span>` : '—';
  }

  // Group records by the selected groupBy key
  const grp = invState.groupBy;
  function getGroupKey(r) {
    if (grp === 'body')    return String(r.Body_Class    || '(blank)');
    if (grp === 'chassis') return String(r.CHASSIS_MODEL || '(blank)');
    if (grp === 'color')   return String(r.Color         || '(blank)');
    return 'All';
  }

  // Build group map preserving individual records
  const groupMap = {};
  available.forEach(r => {
    const key = getGroupKey(r);
    if (!groupMap[key]) groupMap[key] = [];
    groupMap[key].push(r);
  });

  // Sort each group oldest arrival first
  Object.values(groupMap).forEach(arr => arr.sort((a, b) => {
    const da = parseDate(a.Body_Pickup_Date) || parseDate(a.NRC_EDD) || new Date(0);
    const db = parseDate(b.Body_Pickup_Date) || parseDate(b.NRC_EDD) || new Date(0);
    return da - db;
  }));

  const thS = "padding:6px 8px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#ffffff;border-bottom:1px solid rgba(255,255,255,0.2);white-space:nowrap;";

  // Sub-row columns depend on groupBy — don't repeat the group key column
 function subColHeader() {
  const s = 'padding:6px 8px;text-align:left;font-family:"Barlow Condensed",sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#ffffff !important;border-bottom:1px solid rgba(255,255,255,0.2);white-space:nowrap;background:#ffffff;';
  const cols = [];
  if (grp !== 'body')    cols.push(`<td style="${s}">Body Class</td>`);
  if (grp !== 'chassis') cols.push(`<td style="${s}">Chassis / Model</td>`);
  if (grp !== 'color')   cols.push(`<td style="${s}">Color</td>`);
  cols.push(
    `<td style="${s}">OP #</td>`,
    `<td style="${s}">Arrival Date</td>`,
    `<td style="${s}">Proj Build End</td>`,
    `<td style="${s}">Days in Stock</td>`
  );
  return cols.join('');
}

  function subRow(r, idx) {
    const bg      = idx % 2 === 0 ? 'var(--inservice-odd)' : 'var(--inservice-even)';
    const projEnd = parseDate(r.Projected_Build_End_Date);
    const days    = daysInInventory(r);
    const cols    = [];
    if (grp !== 'body')    cols.push(`<td style="padding:5px 8px 5px 28px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.Body_Class||'')}</td>`);
    if (grp !== 'chassis') cols.push(`<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;max-width:150px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.CHASSIS_MODEL||'')}">${escHtml(r.CHASSIS_MODEL||'')}</td>`);
    if (grp !== 'color')   cols.push(`<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${escHtml(r.Color||'')}</td>`);
    cols.push(
      `<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;font-family:'Barlow Condensed',sans-serif;">${escHtml(r.OP_||'')}</td>`,
      `<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${arrivalCell(r)}</td>`,
      `<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;">${projEnd ? fmtDate(projEnd) : '—'}</td>`,
      `<td style="padding:5px 8px;border-bottom:1px solid var(--border);">${daysChip(days)}</td>`
    );
    return `<tr style="background:${bg};" onmouseover="this.style.filter='brightness(0.96)'" onmouseout="this.style.filter=''">${cols.join('')}</tr>`;
  }

  // Column count for colspan — 3 detail cols + 4 fixed = 7, minus the group key col = 6
  const colCount = (grp === 'flat' ? 3 : 2) + 4; // body+chassis+color minus groupby key + op+arrival+projend+days

  let tableRows = '', grand = 0;
  Object.entries(groupMap).sort((a,b) => a[0].localeCompare(b[0])).forEach(([key, records]) => {
    grand += records.length;

    // Group header row
    tableRows += `<tr style="background:var(--navy-mid);">
      <td colspan="${colCount}" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:0.08em;color:#fff;">
        ${escHtml(key)}
        <span style="display:inline-block;margin-left:8px;background:rgba(255,255,255,0.2);border-radius:10px;padding:1px 8px;font-size:10px;">${records.length}</span>
      </td>
    </tr>`;

    // Sub-column header row
    tableRows += `<tr style="background:#1a3050;color:#ffffff;">${subColHeader()}</tr>`;

    // Individual build rows
    records.forEach((r, i) => { tableRows += subRow(r, i); });
  });

  const groupByLabel = grp === 'body' ? 'BODY CLASS' : grp === 'chassis' ? 'CHASSIS' : grp === 'color' ? 'COLOR' : 'ALL';

  area.innerHTML = `
    <div style="font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:0.08em;color:var(--text-muted);margin-bottom:6px;">
      ${available.length} unit${available.length !== 1 ? 's' : ''} in stock by selected date &nbsp;·&nbsp;
      <span style="color:#059669;">■</span> ≤30 days &nbsp;
      <span style="color:#d97706;">■</span> 31–60 days &nbsp;
      <span style="color:#dc2626;">■</span> 60+ days in stock
    </div>
    <table style="width:100%;border-collapse:collapse;min-width:700px;box-shadow:0 1px 4px rgba(0,0,0,0.08);border-radius:4px;overflow:hidden;">
      <thead>
        <tr style="background:var(--navy);">
          <td colspan="${colCount}" style="padding:6px 8px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;background:#1a2e4a;color:#ffffff;border-bottom:1px solid rgba(255,255,255,0.2);     white-space:nowrap;">GROUPED BY ${groupByLabel}</td>
        </tr>
      </thead>
      <tbody>${tableRows}</tbody>
      <tfoot>
        <tr style="background:var(--navy);">
          <td colspan="${colCount - 1}" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:var(--gold);">TOTAL STOCK INVENTORY</td>
          <td style="padding:7px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:var(--gold);">${grand}</td>
        </tr>
      </tfoot>
    </table>`;
}

// ── Inventory Value Line Chart ──
  
function renderInventoryValueChart(snapDate) {
  destroyChart('inv-value');
  const ctx = document.getElementById('chart-inv-value');
  if (!ctx) return;

  const projMode   = invState.projMode;
  const today      = new Date(); today.setHours(0,0,0,0);
  const startPoint = new Date(today); startPoint.setMonth(startPoint.getMonth() - 6);
  const endPoint   = new Date(today); endPoint.setMonth(endPoint.getMonth() + 6);

  const labels = [], projData = [];
  let cursor = new Date(startPoint);

  while (cursor <= endPoint) {
    const snap = new Date(cursor);
    labels.push(new Intl.DateTimeFormat('en-US', { month:'short', day:'numeric' }).format(snap));

    let projVal = 0;
    if (projMode === 'none') {
      // No sales projected — body enters inventory on arrival, never leaves
      projVal = DB.backlog
        .filter(r => {
          if (r.Delivered_) return false;
          const arrivalDate = parseDate(r.Body_Pickup_Date) || parseDate(r.NRC_EDD);
          return arrivalDate && arrivalDate <= snap;
        })
        .reduce((s, r) => s + (r.Body_Cost || 0), 0);
    } else if (projMode === 'pending') {
      // With projected sales — body leaves when projected delivery date passes
      projVal = DB.backlog
        .filter(r => {
          if (r.Delivered_) return false;
          const arrivalDate = parseDate(r.Body_Pickup_Date) || parseDate(r.NRC_EDD);
          if (!arrivalDate || arrivalDate > snap) return false;
          const projDel = parseDate(r.Projected_Date_of_Delivery)
                       || parseDate(r.Manual_Expected_Delivery_Date);
          if (projDel && projDel <= snap) return false;
          const delDate = parseDate(r.Date_of_Delivery);
          if (delDate && delDate <= snap) return false;
          return true;
        })
        .reduce((s, r) => s + (r.Body_Cost || 0), 0);
    }
    projData.push(projVal);

    cursor.setDate(cursor.getDate() + 7);
  }

  charts['inv-value'] = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: projMode === 'none' ? 'Inventory Value — No Sales Projected' : 'Inventory Value — With Projected Sales',
        data: projData,
        borderColor: '#1e3a5f',
        backgroundColor: 'rgba(30,58,95,0.12)',
        fill: true,
        tension: 0.3,
        pointRadius: 2,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { labels: { font: { family:'Barlow', size:10 } } },
        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmtCurrency(ctx.raw, true) } }
      },
      scales: {
        y: { ticks: { callback: v => fmtCurrency(v, true), font: { family:'Barlow', size:9 } } },
        x: { ticks: { font: { family:'Barlow', size:9 }, maxTicksLimit: 12 } }
      }
    }
  });
}

// ═══════════════════════════════════════════════════════════
// PAGE 6 — NRC ORDER TRACKING
// ═══════════════════════════════════════════════════════════

function renderNRC() {
  console.log('DB.nrcOrders length:', (DB.nrcOrders || []).length);
  console.log('sample NRC row:', DB.nrcOrders?.[0]);
  const today = new Date(); today.setHours(0,0,0,0);

  // Build a lookup of Body_Cost from backlog by OP_ number
  const bodyCostByOP = {};
  DB.backlog.forEach(r => {
    if (r.OP_) bodyCostByOP[String(r.OP_).trim()] = r.Body_Cost || 0;
  });

  const bodyConfirmedByOP = {};
  const rowIdByOP = {};
  DB.backlog.forEach(r => {
   if (r.OP_) {
    const key = String(r.OP_).trim();
    bodyConfirmedByOP[key] = !!r.Body_Confirmed;
    rowIdByOP[key]         = r.id;
    }
  });

  // Clean display name
  function nrcName(r) {
    const raw = String(r.Name || '').trim().toUpperCase();
    return raw === 'NEW ENGLAND WRECKER SALES' ? 'STOCK' : (r.Name || '—');
  }

  // Total drift: NRC_EDD vs NRC_EDD_Original (null original = no drift)
  function totalDrift(r) {
    const current  = parseDate(r.NRC_EDD);
    const original = parseDate(r.NRC_EDD_Original);
    if (!current || !original) return 0;
    return Math.round((current - original) / 86400000);
  }

  // Filter: must have NRC_EDD, must be today or future
  const base = (DB.nrcOrders || []).filter(r => {
    const edd = parseDate(r.NRC_EDD);
    return edd && edd >= today;
  });

  // Populate status dropdown
  const statuses = new Set(base.map(r => r.NRC_Status).filter(Boolean));
  const statusEl = document.getElementById('nrc-status-sel');
  if (statusEl && statusEl.options.length <= 1) {
    [...statuses].sort().forEach(s => {
      const o = document.createElement('option');
      o.value = s; o.textContent = s;
      statusEl.appendChild(o);
    });
  }

  // Apply filters
  const filtered = base.filter(r => {
    const edd = parseDate(r.NRC_EDD);
    if (nrcFilters.from && edd < nrcFilters.from) return false;
    if (nrcFilters.to   && edd > nrcFilters.to)   return false;
    if (nrcFilters.status !== 'ALL' && r.NRC_Status !== nrcFilters.status) return false;
    return true;
  });

  // Enrich with computed fields
  const enriched = filtered.map(r => {
  const edd           = parseDate(r.NRC_EDD);
  const drift         = totalDrift(r);
  const lastShift     = r.EDD_Day_Difference || 0;
  const bodyCost      = bodyCostByOP[String(r.OP_).trim()] || 0;
  const bodyConfirmed = bodyConfirmedByOP[String(r.OP_).trim()] || false;
  const backlogRowId  = rowIdByOP[String(r.OP_).trim()] || null;
  return { ...r, edd, drift, lastShift, bodyCost, displayName: nrcName(r), bodyConfirmed, backlogRowId };
}).sort((a, b) => a.edd - b.edd);

  // ── KPIs ──
  const totalBodies    = enriched.length;
  const totalBodyCost  = enriched.reduce((s, r) => s + r.bodyCost, 0);
  const avgDrift       = enriched.length ? (enriched.reduce((s,r) => s+r.drift, 0) / enriched.length).toFixed(1) : 0;
  const delayed        = enriched.filter(r => r.drift > 0).length;
  const early          = enriched.filter(r => r.drift < 0).length;
  const onTime         = enriched.filter(r => r.drift === 0).length;

  const kpiStrip = document.getElementById('nrc-kpi-strip');
  if (kpiStrip) kpiStrip.innerHTML = `
    <div class="kpi-card inprog" style="flex:1;min-width:130px;">
      <div class="kpi-label">Expected Bodies</div>
      <div class="kpi-value" style="color:var(--kpi-inprog);">${totalBodies}</div>
      <div class="kpi-sub">upcoming deliveries</div>
    </div>
    <div class="kpi-card payment" style="flex:1;min-width:130px;">
      <div class="kpi-label">Total Body Cost</div>
      <div class="kpi-value" style="font-size:20px;color:var(--kpi-payment);">${fmtCurrency(totalBodyCost, true)}</div>
      <div class="kpi-sub">incoming value</div>
    </div>
    <div class="kpi-card ${parseFloat(avgDrift) > 0 ? 'payment' : 'onlot'}" style="flex:1;min-width:130px;">
      <div class="kpi-label">Avg Total Drift</div>
      <div class="kpi-value" style="font-size:22px;color:${parseFloat(avgDrift) > 7 ? '#dc2626' : parseFloat(avgDrift) > 0 ? '#d97706' : '#059669'};">${avgDrift > 0 ? '+' : ''}${avgDrift}d</div>
      <div class="kpi-sub">from original EDD</div>
    </div>
    <div class="kpi-card" style="flex:1;min-width:130px;border-top:3px solid #dc2626;">
      <div class="kpi-label">Delayed</div>
      <div class="kpi-value" style="color:#dc2626;">${delayed}</div>
      <div class="kpi-sub">EDD pushed later</div>
    </div>
    <div class="kpi-card" style="flex:1;min-width:130px;border-top:3px solid #059669;">
      <div class="kpi-label">On Time / Early</div>
      <div class="kpi-value" style="color:#059669;">${onTime + early}</div>
      <div class="kpi-sub">${early} pulled earlier</div>
    </div>`;

  // ── Charts ──
  renderNRCCharts(enriched);

  // ── Main table grouped by month ──
  const tableArea = document.getElementById('nrc-table-area');
  if (!tableArea) return;

  if (!enriched.length) {
    tableArea.innerHTML = '<p style="color:var(--text-muted);font-family:Barlow Condensed,sans-serif;letter-spacing:0.1em;padding:20px;text-align:center;">NO NRC ORDERS MATCH CURRENT FILTERS</p>';
    return;
  }

  // Group by month
  const monthGroups = {};
  enriched.forEach(r => {
    const key = r.edd.getFullYear() + '-' + String(r.edd.getMonth()+1).padStart(2,'0');
    if (!monthGroups[key]) monthGroups[key] = [];
    monthGroups[key].push(r);
  });

  const thStyle = "padding:6px 8px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#ffffff;border-bottom:1px solid rgba(255,255,255,0.2);white-space:nowrap;background:#2a4f7a;";

  function driftChip(days) {
    if (days === 0) return `<span style="font-family:'Barlow Condensed',sans-serif;font-size:11px;color:#059669;font-weight:600;">No change</span>`;
    const color = days > 0 ? '#dc2626' : '#059669';
    const sign  = days > 0 ? '+' : '';
    return `<span style="font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:${color};">${sign}${days}d</span>`;
  }

  function statusChip(status) {
    if (!status) return '—';
    const bg = status.toLowerCase().includes('production') ? 'rgba(37,99,235,0.12)' : 'rgba(107,114,128,0.12)';
    const color = status.toLowerCase().includes('production') ? '#2563eb' : '#6b7280';
    return `<span style="font-size:10px;padding:2px 7px;border-radius:2px;background:${bg};color:${color};font-family:'Barlow Condensed',sans-serif;letter-spacing:0.06em;white-space:nowrap;">${escHtml(status)}</span>`;
  }

  let tableHtml = '';
  Object.keys(monthGroups).sort().forEach(monthKey => {
    const rows      = monthGroups[monthKey];
    const d         = new Date(monthKey + '-01T00:00:00');
    const monthName = new Intl.DateTimeFormat('en-US', { month:'long', year:'numeric' }).format(d);
    const monthCost = rows.reduce((s, r) => s + r.bodyCost, 0);

    tableHtml += `
      <table style="width:100%;border-collapse:collapse;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,0.08);border-radius:4px;overflow:hidden;min-width:900px;">
        <thead>
          <tr style="background:var(--inservice-hdr);">
            <td colspan="11" style="padding:7px 10px;font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:#fff;">
              ${escHtml(monthName)}
              <span style="display:inline-block;margin-left:8px;background:rgba(255,255,255,0.2);border-radius:10px;padding:1px 8px;font-size:10px;">${rows.length}</span>
              <span style="float:right;font-size:11px;font-weight:400;opacity:0.85;">${fmtCurrency(monthCost, true)} body cost</span>
            </td>
          </tr>
          <tr>
            <th style="${thStyle}">OP #</th>
            <th style="${thStyle}">Customer</th>
            <th style="${thStyle}">Body Type</th>
            <th style="${thStyle}">NRC EDD</th>
            <th style="${thStyle}">Original EDD</th>
            <th style="${thStyle}">Total Drift</th>
            <th style="${thStyle}">Last Shift</th>
            <th style="${thStyle}">Last Update</th>
            <th style="${thStyle}">Supplier Status</th>
            <th style="${thStyle};text-align:center;">Confirmed</th>
            <th style="${thStyle};text-align:right;">Body Cost</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r, idx) => {
            const bg          = idx % 2 === 0 ? 'var(--inservice-odd)' : 'var(--inservice-even)';
            const origEDD     = parseDate(r.NRC_EDD_Original);
            const lastUpdate  = parseDate(r.EDD_Update_Received);
            const tdS         = 'padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;';
            return `<tr style="background:${bg};" onmouseover="this.style.filter='brightness(0.96)'" onmouseout="this.style.filter=''">
              <td style="${tdS}font-family:'Barlow Condensed',sans-serif;">${escHtml(String(r.OP_||''))}</td>
              <td style="${tdS}max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.displayName)}">${escHtml(r.displayName)}</td>
              <td style="${tdS}">${escHtml(r.Body_Type||'—')}</td>
              <td style="${tdS}font-weight:600;">${fmtDate(r.NRC_EDD)}</td>
              <td style="${tdS}color:var(--text-dim);">${origEDD ? fmtDate(origEDD) : '<span style="color:var(--text-muted);font-size:10px;">Unchanged</span>'}</td>
              <td style="${tdS}">${driftChip(r.drift)}</td>
              <td style="${tdS}">${r.lastShift !== 0 ? driftChip(r.lastShift) : '<span style="color:var(--text-muted);font-size:10px;">—</span>'}</td>
              <td style="${tdS}color:var(--text-dim);">${lastUpdate ? fmtDate(lastUpdate) : '—'}</td>
              <td style="${tdS}">${statusChip(r.NRC_Status)}</td>
              <td style="${tdS}text-align:center;">
                ${r.backlogRowId ? `
                <div class="toggle-wrap">
                  <label class="bool-toggle nrc-confirm-toggle" data-row="${r.backlogRowId}" data-col="Body_Confirmed">
                  <input type="checkbox" ${r.bodyConfirmed ? 'checked' : ''}/>
                  <span class="slider"></span>
                 </label>
                </div>` : '—'}
             </td>
            <td style="${tdS}text-align:right;font-family:'Barlow Condensed',sans-serif;">${r.bodyCost ? fmtCurrency(r.bodyCost) : '—'}</td>
            </tr>`;
          }).join('')}
        </tbody>
        <tfoot>
          <tr style="background:var(--inservice-hdr);">
            <td colspan="10" style="padding:6px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;">SUBTOTAL</td>
            <td style="padding:6px 10px;text-align:right;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;color:#fff;">${fmtCurrency(monthCost)}</td>
          </tr>
        </tfoot>
      </table>`;
  });

  tableArea.innerHTML = `<div style="overflow-x:auto;">${tableHtml}</div>`;
}

// Wire up Body_Confirmed toggles
document.querySelectorAll('.nrc-confirm-toggle input').forEach(inp => {
  inp.addEventListener('change', async e => {
    const wrap  = e.target.closest('.nrc-confirm-toggle');
    const rowId = parseInt(wrap.dataset.row);
    const val   = e.target.checked;
    wrap.classList.add('saving');
    try {
      await grist.docApi.applyUserActions([
        ['UpdateRecord', 'NEWS_Build_Backlog', rowId, { Body_Confirmed: val }]
      ]);
      wrap.classList.remove('saving');
      // Update local DB so it stays in sync without a full reload
      const rec = DB.backlog.find(b => b.id === rowId);
      if (rec) rec.Body_Confirmed = val;
      showToast('Saved');
    } catch(e) {
      wrap.classList.remove('saving');
      wrap.classList.add('error');
      setTimeout(() => wrap.classList.remove('error'), 2000);
      showToast('Save failed: ' + e.message, true);
      e.target.checked = !val;
    }
  });
});

function renderNRCCharts(enriched) {
  // ── Arrivals by month bar chart ──
  destroyChart('nrc-arrivals');
  const monthCounts = {}, monthCosts = {};
  enriched.forEach(r => {
    const key = new Intl.DateTimeFormat('en-US',{month:'short',year:'2-digit'}).format(r.edd);
    monthCounts[key] = (monthCounts[key] || 0) + 1;
    monthCosts[key]  = (monthCosts[key]  || 0) + r.bodyCost;
  });
  const monthKeys = Object.keys(monthCounts);
  const ctx1 = document.getElementById('chart-nrc-arrivals');
  if (ctx1) {
    charts['nrc-arrivals'] = new Chart(ctx1, {
      type: 'bar',
      data: {
        labels: monthKeys,
        datasets: [
          { label:'Bodies', data: monthKeys.map(k => monthCounts[k]), backgroundColor:'rgba(30,58,95,0.75)', borderRadius:2, yAxisID:'y' },
          { label:'Body Cost', data: monthKeys.map(k => monthCosts[k]), type:'line', borderColor:'#e8a020', borderWidth:2, pointRadius:2, fill:false, yAxisID:'y1' },
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:true,
        plugins:{ legend:{ labels:{ font:{family:'Barlow',size:10} }}},
        scales:{
          y:  { position:'left',  ticks:{ font:{family:'Barlow',size:9} } },
          y1: { position:'right', ticks:{ callback:v=>fmtCurrency(v,true), font:{family:'Barlow',size:9} }, grid:{ drawOnChartArea:false } }
        }
      }
    });
  }

  // ── Biggest last-change table (replaces status chart) ──
const ctx2container = document.getElementById('chart-nrc-status')?.parentElement;
if (ctx2container) {
  // Sort by absolute value of last shift, take top 5
  const top5 = [...enriched]
    .filter(r => r.lastShift !== 0)
    .sort((a, b) => Math.abs(b.lastShift) - Math.abs(a.lastShift))
    .slice(0, 5);

  const thS = "padding:6px 8px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#ffffff;border-bottom:1px solid rgba(255,255,255,0.2);white-space:nowrap;background:#2a4f7a;";

  ctx2container.innerHTML = `
    <div style="font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--text-muted);margin-bottom:10px;">Biggest Recent Changes</div>
    ${top5.length ? `
    <table style="width:100%;border-collapse:collapse;font-size:11.5px;border-radius:4px;overflow:hidden;">
      <thead>
        <tr>
          <td style="${thS}">OP #</td>
          <td style="${thS}">Customer</td>
          <td style="${thS}">Body Type</td>
          <td style="${thS}">NRC EDD</td>
          <td style="${thS};text-align:right;">Last Change</td>
        </tr>
      </thead>
      <tbody>
        ${top5.map((r, idx) => {
          const bg        = idx % 2 === 0 ? 'var(--inservice-odd)' : 'var(--inservice-even)';
          const shiftColor = r.lastShift > 0 ? '#dc2626' : '#059669';
          const sign       = r.lastShift > 0 ? '+' : '';
          const tdS        = 'padding:6px 8px;border-bottom:1px solid var(--border);white-space:nowrap;';
          return `<tr style="background:${bg};">
            <td style="${tdS}font-family:'Barlow Condensed',sans-serif;">${escHtml(String(r.OP_||''))}</td>
            <td style="${tdS}max-width:120px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.displayName)}">${escHtml(r.displayName)}</td>
            <td style="${tdS}">${escHtml(r.Body_Type||'—')}</td>
            <td style="${tdS}">${fmtDate(r.NRC_EDD)}</td>
            <td style="${tdS};text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:700;color:${shiftColor};">${sign}${r.lastShift}d</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>` : '<p style="color:var(--text-muted);font-size:11px;font-style:italic;">No recent changes</p>'}`;
}

  // ── EDD drift distribution bar chart ──
  destroyChart('nrc-drift');
  const driftBuckets = { 'Early (<0d)':0, 'No change':0, '1–7d':0, '8–14d':0, '15–30d':0, '>30d':0 };
  enriched.forEach(r => {
    if      (r.drift < 0)   driftBuckets['Early (<0d)']++;
    else if (r.drift === 0) driftBuckets['No change']++;
    else if (r.drift <= 7)  driftBuckets['1–7d']++;
    else if (r.drift <= 14) driftBuckets['8–14d']++;
    else if (r.drift <= 30) driftBuckets['15–30d']++;
    else                    driftBuckets['>30d']++;
  });
  const driftKeys   = Object.keys(driftBuckets);
  const driftColors = ['#059669','#6b7280','#fbbf24','#f59e0b','#ef4444','#dc2626'];
  const ctx3 = document.getElementById('chart-nrc-drift');
  if (ctx3) {
    charts['nrc-drift'] = new Chart(ctx3, {
      type:'bar',
      data:{
        labels: driftKeys,
        datasets:[{ label:'Orders', data: driftKeys.map(k=>driftBuckets[k]),
          backgroundColor: driftColors, borderRadius:3 }]
      },
      options:{
        responsive:true, maintainAspectRatio:true,
        plugins:{ legend:{ display:false } },
        scales:{
          y:{ ticks:{ font:{family:'Barlow',size:9} } },
          x:{ ticks:{ font:{family:'Barlow',size:9} } }
        }
      }
    });
  }
}

// ═══════════════════════════════════════════════════════════
// PAGE NAVIGATION
// ═══════════════════════════════════════════════════════════

const PAGE_RENDER = {
  operations: renderOperations,
  gantt:      renderGantt,
  sales:      renderSales,
  cashflow:   renderCashFlow,
  inventory:  renderInventory,
  nrc:        renderNRC,
};

function navigateTo(pageKey) {
  activePage = pageKey;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  const page = document.getElementById('page-' + pageKey);
  const tab  = document.querySelector(`.nav-tab[data-page="${pageKey}"]`);
  if (page) page.classList.add('active');
  if (tab)  tab.classList.add('active');
  setTimeout(() => {
    if (PAGE_RENDER[pageKey]) {
      try { PAGE_RENDER[pageKey](); }
      catch(e) { console.error('Page render error ['+pageKey+']:', e); }
    }
  }, 50);
}

// ═══════════════════════════════════════════════════════════
// EVENT LISTENERS
// ═══════════════════════════════════════════════════════════

document.querySelectorAll('.nav-tab').forEach(tab => {
  tab.addEventListener('click', () => navigateTo(tab.dataset.page));
});

document.getElementById('refresh-btn').addEventListener('click', async () => {
  await loadSecondaryTables();
  if (PAGE_RENDER[activePage]) PAGE_RENDER[activePage]();
});

['ops-filter-name','ops-filter-chassis','ops-filter-body'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener('change', e => {
    if (id==='ops-filter-name')    opsFilters.name    = e.target.value;
    if (id==='ops-filter-chassis') opsFilters.chassis = e.target.value;
    if (id==='ops-filter-body')    opsFilters.body    = e.target.value;
    renderBacklog();
  });
});

// Chart year filter (filters revenue chart + body class chart)
const chartYearEl = document.getElementById('chart-year-sel');
if (chartYearEl) {
  chartYearEl.addEventListener('change', e => {
    chartYear = e.target.value;
    renderMonthlyRevenueChart();
    renderBodyClassChart();
  });
}

// Monthly table month/year selectors (independent from chart year)
const salesMonthEl = document.getElementById('sales-month-sel');
const salesYearEl  = document.getElementById('sales-year-sel');
if (salesMonthEl) {
  salesMonthEl.value = String(selectedMonth);
  salesMonthEl.addEventListener('change', e => { selectedMonth = parseInt(e.target.value); renderMonthlySalesWIP(); });
}
if (salesYearEl) {
  salesYearEl.addEventListener('change', e => { selectedYear = parseInt(e.target.value); renderMonthlySalesWIP(); });
}

// Cash flow controls
document.querySelectorAll('#cf-period-toggle button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#cf-period-toggle button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    cfPeriod = btn.dataset.period;
    renderCashFlow();
  });
});
const cfHorizonEl = document.getElementById('cf-horizon');
if (cfHorizonEl) cfHorizonEl.addEventListener('change', e => { cfHorizon = parseInt(e.target.value); renderCashFlow(); });

// Gantt range
document.querySelectorAll('#gantt-range-toggle button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#gantt-range-toggle button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    ganttRange = btn.dataset.range;
    renderGantt();
  });
});

// Inventory controls
const invSnapEl = document.getElementById('inv-snap-date');
if (invSnapEl) {
  invSnapEl.value = toInputDate(invState.snapDate);
  invSnapEl.addEventListener('change', e => {
    invState.snapDate = new Date(e.target.value+'T00:00:00');
    renderInventory();
  });
}
['inv-filter-body','inv-filter-chassis','inv-filter-color'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener('change', e => {
    if (id==='inv-filter-body')    invState.body    = e.target.value;
    if (id==='inv-filter-chassis') invState.chassis = e.target.value;
    if (id==='inv-filter-color')   invState.color   = e.target.value;
    renderInventory();
  });
});
document.querySelectorAll('#inv-group-toggle button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#inv-group-toggle button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    invState.groupBy = btn.dataset.grp;
    renderInventory();
  });
});
const invProjModeEl = document.getElementById('inv-proj-mode');
if (invProjModeEl) {
  invProjModeEl.addEventListener('change', e => {
    invState.projMode = e.target.value;
    renderInventoryValueChart(invState.snapDate);
  });
}

// NRC filters
const nrcDateFrom  = document.getElementById('nrc-date-from');
const nrcDateTo    = document.getElementById('nrc-date-to');
const nrcStatusSel = document.getElementById('nrc-status-sel');

// Set default from date to today
if (nrcDateFrom) {
  nrcDateFrom.value = toInputDate(nrcFilters.from);
  nrcDateFrom.addEventListener('change', e => {
    nrcFilters.from = e.target.value ? new Date(e.target.value+'T00:00:00') : null;
    renderNRC();
  });
}
if (nrcDateTo) {
  nrcDateTo.addEventListener('change', e => {
    nrcFilters.to = e.target.value ? new Date(e.target.value+'T00:00:00') : null;
    renderNRC();
  });
}
if (nrcStatusSel) {
  nrcStatusSel.addEventListener('change', e => {
    nrcFilters.status = e.target.value;
    renderNRC();
  });
}

// ═══════════════════════════════════════════════════════════
// GRIST API
// ═══════════════════════════════════════════════════════════

grist.ready({
  requiredAccess: 'full',
  columns: [
    { name:'OP_',                                       title:'OP #',                   type:'Text'    },
    { name:'SO_',                                       title:'SO #',                   type:'Text'    },
    { name:'NAME',                                      title:'Customer',               type:'Choice'  },
    { name:'CHASSIS_MODEL',                             title:'Chassis / Model',        type:'Choice'  },
    { name:'CHASSIS_VIN_',                              title:'VIN #',                  type:'Text'    },
    { name:'Body_Class',                                title:'Body Class',             type:'Text'    },
    { name:'SERIAL_',                                   title:'Serial #',               type:'Text'    },
    { name:'Actual_Build_Start_Date',                   title:'Build Start',            type:'Date'    },
    { name:'Actual_Build_End_Date',                     title:'Build End (Actual)',      type:'Text'   },
    { name:'Paid_Date',                                 title:'Paid Date',              type:'Date'    },
    { name:'Date_of_Delivery',                          title:'Date of Delivery',       type:'Date'    },
    { name:'Projected_Build_End_Date',                  title:'Proj Build End',         type:'Date'    },
    { name:'Projected_Build_Start_Date',                title:'Proj Build Start',       type:'Date'    },
    { name:'Manual_Override_Projected_Build_Start_Date',title:'Manual Override Start',  type:'Date'    },
    { name:'ETA_WK_SALE_MONTH_FINAL_SALE_DATE',         title:'ETA / Final Sale Date',  type:'Date'    },
    { name:'Projected_Next_Month_Filter',               title:'Proj Next Month Filter', type:'Any'     },
    { name:'Projected_Date_of_Delivery',                title:'Projected Delivery',     type:'Any'     },
    { name:'Manual_Expected_Delivery_Date',             title:'Manual Expected Delivery',type:'Date'   },
    { name:'NRC_EDD',                                   title:'NRC EDD',                type:'Date'    },
    { name:'Body_Pickup_Date',                          title:'Body Pickup Date',       type:'Date'    },
    { name:'Build_Started_',                            title:'Build Started',          type:'Bool'    },
    { name:'Build_Ended_',                              title:'Build Ended',            type:'Bool'    },
    { name:'Delivered_',                                title:'Delivered',              type:'Bool'    },
    { name:'Paid_',                                     title:'Paid',                   type:'Bool'    },
    { name:'Toolbox',                                   title:'Toolbox',                type:'Bool'    },
    { name:'Body_Picked_Up',                            title:'Body Picked Up',         type:'Bool'    },
    { name:'Chassis_Delivered',                         title:'Chassis Delivered',      type:'Bool'    },
    { name:'TOTAL_RETAIL_SALE',                         title:'Total Retail Sale',      type:'Numeric' },
    { name:'Body_Cost',                                 title:'Body Cost',              type:'Numeric' },
    { name:'CHASSIS_COST',                              title:'Chassis Cost',           type:'Numeric' },
    { name:'TRUE_PROFIT',                               title:'True Profit',            type:'Numeric' },
    { name:'Total_Costs',                               title:'Total Costs',            type:'Numeric' },
    { name:'Color',                                     title:'Color',                  type:'Text'    },
    { name:'NOTES',                                     title:'Notes',                  type:'Text'    },
    { name:'Actual_Labor_Cost',                         title:'Actual Labor Cost',      type:'Numeric' },
    { name:'INTERNAL_BUILD_COST_PARTS_',                title:'Parts Cost',             type:'Numeric' },
    { name:'Chassis_EDD',                               title:'Chassis EDD',            type:'Date'    },
    { name:'Chassis_Delivery_Date',                     title:'Chassis Delivery Date',  type:'Date'    },
    { name:'Body_Confirmed',                            title:'Body Confirmed',         type:'Bool'    },
  ],
});

grist.onRecords(async function(records, mappings) {
  const mapped = records.map(r => {
    const m = grist.mapColumnNames(r, mappings);
    if (m) m.id = r.id;
    return m || { ...r.fields, id: r.id };
  });
  DB.backlog = mapped;
  if (!DB.pipeline.length) await loadSecondaryTables();
  document.getElementById('global-loading').style.display = 'none';
  navigateTo(activePage);
});
