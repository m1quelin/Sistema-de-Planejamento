import { showGlobalLoading, hideGlobalLoading, captureToClipboard } from "./ui.js";
import { listenPlanilhaMetadata } from "./admin.js";
let currentModalRows = [];
let sortState = { column: 'data', direction: 'desc' };

const MONTH_NAMES = [
  "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
  "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"
];
const MONTH_COLS = {
  JANEIRO:[2,3],FEVEREIRO:[5,6],MARÇO:[8,9],ABRIL:[11,12],
  MAIO:[14,15],JUNHO:[17,18],JULHO:[20,21],AGOSTO:[23,24],
  SETEMBRO:[26,27],OUTUBRO:[29,30],NOVEMBRO:[32,33],DEZEMBRO:[35,36]
};

let globalData = null;
let currentModalMonth = null;
let unsubPlanilha = null;

// ─── INIT ─────────────────────────────────────────────────────
export function initDashboard() {
  if (unsubPlanilha) unsubPlanilha();

  unsubPlanilha = listenPlanilhaMetadata(async (meta) => {
    const adminInfo = document.getElementById("adminCurrentInfo");
    if (adminInfo && meta) {
      const date = new Date(meta.updatedAt).toLocaleString("pt-BR");
      adminInfo.innerHTML = `<strong>${meta.filename}</strong><br><span style="color:var(--muted);font-size:12px">Atualizado em ${date} por ${meta.updatedBy}</span>`;
    } else if (adminInfo) {
      adminInfo.textContent = "Nenhuma planilha cadastrada.";
    }

    if (!meta || !meta.dataBase64) {
      renderEmptyState();
      return;
    }

    try {
      showGlobalLoading("Carregando planilha...");
      const binary = atob(meta.dataBase64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      parseAndRender(workbook);

      const badge = document.getElementById("savedBadge");
      if (badge) {
        badge.style.display = "flex";
        const txt = badge.querySelector("#savedBadgeText");
        if (txt) txt.textContent = `Sincronizado em ${new Date().toLocaleTimeString("pt-BR")}`;
      }
      hideGlobalLoading();
    } catch (err) {
      hideGlobalLoading();
      console.error("Erro ao carregar planilha:", err);
      renderEmptyState();
    }
  });

  bindModalEvents();
}

// ─── HELPERS ──────────────────────────────────────────────────
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === "number" && typeof XLSX !== "undefined")
    return XLSX.SSF.parse_date_code(val);
  return new Date(val);
}

function formatDate(d) {
  if (!d) return "";
  try { return (d instanceof Date ? d : new Date(d)).toLocaleDateString("pt-BR"); }
  catch { return ""; }
}

function toNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function fmt(n) {
  return n.toLocaleString("pt-BR", {
    style: "currency", currency: "BRL", minimumFractionDigits: 2
  });
}

function fmtShort(n) {
  if (n >= 1e6) return "R$ " + (n / 1e6).toFixed(2).replace(".", ",") + "M";
  if (n >= 1e3) return "R$ " + (n / 1e3).toFixed(1).replace(".", ",") + "K";
  return fmt(n);
}

function tagClass(cat) {
  if (!cat) return "tag-out";
  const c = cat.toLowerCase();
  if (c.includes("publicidade") || c.includes("propaganda")) return "tag-pub";
  if (c.includes("brinde")) return "tag-bri";
  if (c.includes("patroc")) return "tag-pat";
  if (c.includes("doa")) return "tag-doa";
  return "tag-out";
}

// ─── RENDER ───────────────────────────────────────────────────
function renderEmptyState() {
  document.getElementById("main-content").innerHTML = `
    <div class="empty-state">
      <div class="empty-icon">📊</div>
      <h2>Nenhuma planilha cadastrada</h2>
      <p>Aguarde o administrador subir a planilha do mês.<br>
      <span style="color:var(--green);font-size:12px;margin-top:6px;display:block">✓ Os dados aparecerão automaticamente para todos.</span></p>
    </div>`;
  document.getElementById("pdfBtn").disabled = true;
}

function parseAndRender(wb) {
  const mktSheet = wb.Sheets[wb.SheetNames[0]];
  const mktRaw = XLSX.utils.sheet_to_json(mktSheet, { header: 1, defval: null });
  const totalRowIdx = mktRaw.findIndex(r => r[0] === "TOTAL");
  const totalRow = totalRowIdx >= 0 ? mktRaw[totalRowIdx] : mktRaw[9] || [];
  const catRows = mktRaw.slice(2, totalRowIdx >= 0 ? totalRowIdx : 9).filter(r => r[0]);
  const orcadoAnual = toNum(totalRow[1]) || 2640000;

  const months = MONTH_NAMES.map(m => {
    const [oc, rc] = MONTH_COLS[m];
    const orcado = toNum(totalRow[oc]) || 0;
    const realizado = toNum(totalRow[rc]) || 0;
    const categories = catRows.map(r => ({
      name: r[0], realizado: toNum(r[rc]) || 0
    })).filter(c => c.realizado > 0);
    return { month: m, orcado, realizado, categories, hasData: realizado > 0 };
  });

  const lancSheet = wb.Sheets[wb.SheetNames[1]];
  const lancRaw = XLSX.utils.sheet_to_json(lancSheet, { header: 1, defval: null, cellDates: true });
  const lancamentos = lancRaw.slice(1).filter(r => r[1]).map(r => ({
    data: parseDate(r[1]),
    historico: r[2] || "",
    filial: r[3] || "",
    debito: toNum(r[4]) || 0,
    categoria: r[5] || ""
  }));

  globalData = { months, lancamentos, orcadoAnual };
  renderDashboard();
  document.getElementById("pdfBtn").disabled = false;
}

function renderDashboard() {
  const { months, orcadoAnual } = globalData;
  const realizados = months.filter(m => m.hasData);
  const totalRealizado = realizados.reduce((s, m) => s + m.realizado, 0);
  const totalOrcadoAte = realizados.reduce((s, m) => s + m.orcado, 0);
  const pctGeral = totalOrcadoAte > 0 ? (totalRealizado / totalOrcadoAte) * 100 : 0;
  const saldo = totalOrcadoAte - totalRealizado;
  const restante = orcadoAnual - totalRealizado;
  const pctRestante = orcadoAnual > 0 ? (restante / orcadoAnual) * 100 : 0;
  const pctGasto = orcadoAnual > 0 ? (totalRealizado / orcadoAnual) * 100 : 0;
  const kpiClass = pctRestante <= 20 ? "red" : pctRestante <= 40 ? "yellow" : "green";
  const barW = Math.min(pctGasto, 100).toFixed(1);
  const barColor = pctRestante <= 20 ? "var(--red)" : pctRestante <= 40 ? "var(--yellow)" : "var(--green)";

  document.getElementById("main-content").innerHTML = `
    <div class="summary-bar">
      <div class="kpi-card">
        <div class="kpi-label">Orçado Anual</div>
        <div class="kpi-value">${fmtShort(orcadoAnual)}</div>
        <div class="kpi-sub">Total previsto 2026</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Realizado Acumulado</div>
        <div class="kpi-value ${pctGeral > 100 ? "red" : "green"}">${fmtShort(totalRealizado)}</div>
        <div class="kpi-sub">${realizados.length} ${realizados.length === 1 ? "mês" : "meses"} realizados</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Orçamento Restante</div>
        <div class="kpi-value ${kpiClass}">${fmtShort(restante)}</div>
        <div class="progress-bar-wrap" style="margin-top:8px"><div class="progress-bar" style="width:${barW}%;background:${barColor}"></div></div>
        <div class="kpi-sub ${kpiClass}">${pctRestante.toFixed(1)}% restante do orçamento</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Saldo (orçado − realizado)</div>
        <div class="kpi-value ${saldo < 0 ? "red" : "green"}">${fmtShort(Math.abs(saldo))}</div>
        <div class="kpi-sub">${saldo < 0 ? "▲ Acima do orçado" : "▼ Abaixo do orçado"}</div>
      </div>
    </div>
    <div class="section-title">Meses — clique para ver lançamentos de NF</div>
    <div class="months-grid">${months.map(m => {
      const pct = m.orcado ? (m.realizado / m.orcado) * 100 : 0;
      const barColor = pct > 110 ? "var(--red)" : pct > 100 ? "var(--yellow)" : "var(--green)";
      const barW = Math.min(pct, 100).toFixed(1);
      if (!m.hasData) return `<div class="month-card no-data">
        <div class="month-name">${m.month.toLowerCase()}</div>
        <div class="month-val" style="color:var(--muted)">${fmtShort(m.orcado)}</div>
        <div class="month-val-label">orçado</div>
        <div class="badge-future">Sem lançamentos</div>
      </div>`;
      return `<div class="month-card" data-month="${m.month}">
        <div class="month-name">${m.month.toLowerCase()}</div>
        <div class="month-val" style="color:${barColor}">${fmtShort(m.realizado)}</div>
        <div class="month-val-label">realizado de ${fmtShort(m.orcado)}</div>
        <div class="progress-bar-wrap"><div class="progress-bar" style="width:${barW}%;background:${barColor}"></div></div>
        <div class="month-pct" style="color:${barColor}">${pct.toFixed(1)}%</div>
      </div>`;
    }).join("")}</div>
    <div class="section-title">Resumo por Categoria (acumulado)</div>
    ${renderCategoryTable()}
  `;

  document.querySelectorAll(".month-card[data-month]").forEach(card => {
    card.addEventListener("click", () => openModal(card.dataset.month));
  });
}

function renderCategoryTable() {
  const catMap = {};
  globalData.months.forEach(m => m.categories.forEach(c => {
    catMap[c.name] = (catMap[c.name] || 0) + c.realizado;
  }));
  const cats = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
  const total = cats.reduce((s, c) => s + c[1], 0);
  const rows = cats.map(([name, val]) => `<tr>
    <td><span class="tag ${tagClass(name)}">${name}</span></td>
    <td class="td-right">${fmt(val)}</td>
    <td class="td-right">${total > 0 ? (val / total * 100).toFixed(1) : "0.0"}%</td>
  </tr>`).join("");
  return `<div class="table-wrap">
    <div class="table-header"><span>Distribuição de Gastos Realizados</span><span style="font-size:12px;color:var(--muted)">Total: ${fmt(total)}</span></div>
    <table><thead><tr><th>Categoria</th><th style="text-align:right">Realizado (R$)</th><th style="text-align:right">Participação</th></tr></thead>
    <tbody>${rows}</tbody></table>
  </div>`;
}

function openModal(monthName) {
  const MONTH_NUM = MONTH_NAMES.indexOf(monthName) + 1;
  currentModalMonth = monthName;
  currentModalRows = globalData.lancamentos.filter(l => {
    if (!l.data) return false;
    const dt = l.data instanceof Date ? l.data : new Date(l.data);
    return dt.getMonth() + 1 === MONTH_NUM;
  });

  const total = currentModalRows.reduce((s, l) => s + l.debito, 0);
  document.getElementById("modalTitle").textContent =
    `Lançamentos — ${monthName.charAt(0) + monthName.slice(1).toLowerCase()}`;
  document.getElementById("modalSub").textContent =
    `${currentModalRows.length} NFs · Total: ${fmt(total)}`;

  // Reset sort to default
  sortState = { column: 'data', direction: 'desc' };
  applySortAndRender();
  document.getElementById("modalOverlay").classList.add("open");

  const search = document.getElementById("modalSearch");
  search.value = "";
  search.oninput = e => {
    const q = e.target.value.toLowerCase();
    const filtered = q ? currentModalRows.filter(l =>
      l.historico.toLowerCase().includes(q) ||
      l.categoria.toLowerCase().includes(q) ||
      String(l.filial).toLowerCase().includes(q)
    ) : [...currentModalRows];
    applySortAndRender(filtered);
  };

  bindSortHeaders();
}

function applySortAndRender(rows = currentModalRows) {
  const sorted = [...rows].sort((a, b) => {
    let valA, valB;
    if (sortState.column === 'data') {
      valA = a.data instanceof Date ? a.data.getTime() : new Date(a.data).getTime();
      valB = b.data instanceof Date ? b.data.getTime() : new Date(b.data).getTime();
    } else {
      valA = a.debito;
      valB = b.debito;
    }
    if (sortState.direction === 'asc') return valA - valB;
    return valB - valA;
  });
  renderModalTable(sorted);
  updateSortIndicators();
}

function updateSortIndicators() {
  document.querySelectorAll('.sort-header').forEach(th => {
    const col = th.dataset.sort;
    const ind = th.querySelector('.sort-indicator');
    if (!ind) return;
    if (col === sortState.column) {
      th.classList.add('active');
      ind.textContent = sortState.direction === 'asc' ? '▲' : '▼';
    } else {
      th.classList.remove('active');
      ind.textContent = '↕';
    }
  });
}

function bindSortHeaders() {
  document.querySelectorAll('.sort-header').forEach(th => {
    th.onclick = () => {
      const col = th.dataset.sort;
      if (sortState.column === col) {
        sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
      } else {
        sortState.column = col;
        sortState.direction = 'desc';
      }
      const search = document.getElementById('modalSearch');
      const q = search?.value.toLowerCase() || '';
      const rows = q ? currentModalRows.filter(l =>
        l.historico.toLowerCase().includes(q) ||
        l.categoria.toLowerCase().includes(q) ||
        String(l.filial).toLowerCase().includes(q)
      ) : [...currentModalRows];
      applySortAndRender(rows);
    };
  });
}

function renderModalTable(rows) {
  const tbody = document.getElementById("modalTableBody");
  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5">Nenhum lançamento encontrado.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(row => `<tr>
    <td>${formatDate(row.data)}</td>
    <td>${row.historico}</td>
    <td class="td-center">${row.filial}</td>
    <td style="text-align:right;font-weight:600">${fmt(row.debito)}</td>
    <td><span class="tag ${tagClass(row.categoria)}">${row.categoria || "—"}</span></td>
  </tr>`).join("");
}

function bindModalEvents() {
  document.getElementById("modalClose")?.addEventListener("click", () => {
    document.getElementById("modalOverlay").classList.remove("open");
  });
  document.getElementById("modalOverlay")?.addEventListener("click", e => {
    if (e.target === document.getElementById("modalOverlay"))
      document.getElementById("modalOverlay").classList.remove("open");
  });
  document.getElementById("modalPdfBtn")?.addEventListener("click", () => {
    if (currentModalMonth) captureToClipboard("Lançamentos — " + currentModalMonth);
  });
}

