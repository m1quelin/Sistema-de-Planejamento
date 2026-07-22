import { showGlobalLoading, hideGlobalLoading, captureToClipboard } from "./ui.js";
import { listenPlanilhaMetadata } from "./admin.js";
import { fetchFornecedores, listenFornecedores, matchFornecedor, lancamentoId, getCidadesConhecidas, addFornecedor, updateFornecedor, removeFornecedor } from "./fornecedores.js";
import { listenOverrides, getOverride, saveOverride, removeOverride } from "./overrides.js";

// ─── CONSTANTES ──────────────────────────────────────────────
const MONTH_NAMES = [
  "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
  "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"
];
const MONTH_COLS = {
  JANEIRO:[2,3], FEVEREIRO:[5,6], MARÇO:[8,9], ABRIL:[11,12],
  MAIO:[14,15], JUNHO:[17,18], JULHO:[20,21], AGOSTO:[23,24],
  SETEMBRO:[26,27], OUTUBRO:[29,30], NOVEMBRO:[32,33], DEZEMBRO:[35,36]
};

// ─── ESTADO ──────────────────────────────────────────────────
let globalData = null;
let currentModalMonth = null;
let currentModalRows = [];
let sortState = { column: 'data', direction: 'desc' };
let unsubPlanilha = null;

let fornecedoresData = [];
let overridesData = {};

// ─── INIT ────────────────────────────────────────────────────
export function initDashboard() {
  if (unsubPlanilha) unsubPlanilha();

  // Carrega fornecedores e escuta overrides
    listenFornecedores((data) => {
    fornecedoresData = data || [];
    // Se a aba de fornecedores estiver aberta, re-renderiza
    if (document.getElementById("fornTable")) {
      const container = document.getElementById("main-content");
      if (container) renderFornecedoresView(container);
    }
  });
  listenOverrides((data) => {
    overridesData = data || {};
  });

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

export function onFornecedoresUpdate(data) {
  fornecedoresData = data || [];
  console.log("[DEBUG] Fornecedores recebidos:", fornecedoresData.length, fornecedoresData);

  const container = document.getElementById("main-content");
  if (!container) return;

  if (document.getElementById("fornTable")) {
    renderFornecedoresView(container);
  } else if (document.getElementById("lancTable")) {
    renderLancamentosView(container);
  }
}

// ─── DEV MODE ────────────────────────────────────────────────
export function initDevDashboard() {
  const currentMonth = new Date().getMonth();

  const mockMonths = MONTH_NAMES.map((m, i) => ({
    month: m,
    orcado: 220000,
    realizado: i <= currentMonth ? Math.round(220000 * (0.7 + Math.random() * 0.5)) : 0,
    categories: i <= currentMonth ? [
      { name: 'Publicidade e Propaganda', realizado: Math.round(80000 + Math.random() * 40000) },
      { name: 'Brindes', realizado: Math.round(15000 + Math.random() * 10000) },
      { name: 'Patrocinios', realizado: Math.round(30000 + Math.random() * 20000) },
      { name: 'Doacoes', realizado: Math.round(10000 + Math.random() * 5000) },
      { name: 'Outros', realizado: Math.round(5000 + Math.random() * 8000) }
    ] : [],
    hasData: i <= currentMonth
  }));

  const mockLancamentos = [];
  for (let i = 0; i <= currentMonth; i++) {
    const numNfs = 8 + Math.floor(Math.random() * 12);
    for (let j = 0; j < numNfs; j++) {
      const cats = ['Publicidade e Propaganda', 'Brindes', 'Patrocinios', 'Doacoes', 'Outros'];
      mockLancamentos.push({
        data: new Date(2026, i, 1 + Math.floor(Math.random() * 27)),
        historico: `NF ${String(j + 1).padStart(4, '0')} - Fornecedor ${String.fromCharCode(65 + (j % 26))}`,
        filial: Math.random() > 0.5 ? '01' : '02',
        debito: Math.round(2000 + Math.random() * 18000),
        categoria: cats[Math.floor(Math.random() * cats.length)]
      });
    }
  }

  fornecedoresData = [
    { fornecedor: 'Fornecedor A', cidade: 'Presidente Prudente', razao_social: 'Fornecedor A Ltda', source: 'sheets' },
    { fornecedor: 'Fornecedor B', cidade: 'Osvaldo Cruz', razao_social: 'Fornecedor B EIRELI', source: 'sheets' }
  ];

  globalData = { months: mockMonths, lancamentos: mockLancamentos, orcadoAnual: 2640000 };
  renderDashboard();
  bindModalEvents();

  const badge = document.getElementById("savedBadge");
  if (badge) {
    badge.style.display = "flex";
    const txt = badge.querySelector("#savedBadgeText");
    if (txt) txt.textContent = "Modo Dev - dados ficticios";
  }
  document.getElementById("pdfBtn").disabled = false;

  const adminInfo = document.getElementById("adminCurrentInfo");
  if (adminInfo) {
    adminInfo.innerHTML = '<strong>orcamento_mock.xlsx</strong><br><span style="color:var(--muted);font-size:12px">Modo dev - upload desabilitado</span>';
  }
  const adminUploadBtn = document.getElementById("adminUploadBtn");
  if (adminUploadBtn) {
    adminUploadBtn.disabled = true;
    adminUploadBtn.textContent = "Indisponivel no modo dev";
  }
}

// ─── VIEW: DASHBOARD ─────────────────────────────────────────
export function renderDashboardView(container) {
  if (globalData) {
    renderDashboard();
  } else {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📊</div>
        <h2>Nenhuma planilha cadastrada</h2>
        <p>Aguarde o administrador subir a planilha do mes.<br>
        <span style="color:var(--green);font-size:12px;margin-top:6px;display:block">Os dados aparecerão automaticamente para todos.</span></p>
      </div>`;
    document.getElementById("pdfBtn").disabled = true;
  }
}

// ─── VIEW: LANCAMENTOS ───────────────────────────────────────
export function renderLancamentosView(container) {
  if (!globalData) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📋</div>
        <h2>Sem dados para exibir</h2>
        <p>Aguarde o administrador subir a planilha.</p>
      </div>`;
    return;
  }

  let filterMonth = "all";
  let filterCategory = "all";
  let filterCidade = "all";
  let searchTerm = "";
  let sortCol = "data";
  let sortDir = "desc";

  const allCategories = [...new Set(globalData.lancamentos.map(l => l.categoria).filter(Boolean))].sort();
  const allCidades = [...new Set(
    globalData.lancamentos.map(l => {
      const match = matchFornecedor(l.historico, fornecedoresData);
      return match?.cidade;
    }).filter(Boolean)
  )].sort();

  function getFiltered() {
    let rows = [...globalData.lancamentos];

    if (filterMonth !== "all") {
      const monthNum = MONTH_NAMES.indexOf(filterMonth) + 1;
      rows = rows.filter(l => {
        if (!l.data) return false;
        const dt = l.data instanceof Date ? l.data : new Date(l.data);
        return dt.getMonth() + 1 === monthNum;
      });
    }
    if (filterCategory !== "all") {
      rows = rows.filter(l => l.categoria === filterCategory);
    }
    if (filterCidade !== "all") {
      rows = rows.filter(l => {
        const match = matchFornecedor(l.historico, fornecedoresData);
        return match?.cidade === filterCidade;
      });
    }
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      rows = rows.filter(l =>
        l.historico.toLowerCase().includes(q) ||
        (l.categoria || "").toLowerCase().includes(q) ||
        String(l.filial).toLowerCase().includes(q)
      );
    }
    rows.sort((a, b) => {
      let valA, valB;
      if (sortCol === "data") {
        valA = a.data instanceof Date ? a.data.getTime() : new Date(a.data).getTime();
        valB = b.data instanceof Date ? b.data.getTime() : new Date(b.data).getTime();
      } else {
        valA = a.debito;
        valB = b.debito;
      }
      return sortDir === "asc" ? valA - valB : valB - valA;
    });
    return rows;
  }

  function renderTable() {
    const rows = getFiltered();
    const total = rows.reduce((s, l) => s + l.debito, 0);
    const tbody = document.getElementById("lancTableBody");
    if (!tbody) return;

    if (rows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--muted)">Nenhum lancamento encontrado.</td></tr>';
    } else {
      tbody.innerHTML = rows.map(row => {
        const lId = lancamentoId(row, "all");
        const override = getOverride(lId);
        let cidade = "";
        let isOverridden = false;

        if (override) {
          cidade = override.cidade || "";
          isOverridden = true;
        } else {
          const match = matchFornecedor(row.historico, fornecedoresData);
          if (match) cidade = match.cidade || "";
        }

        const cidadeClass = isOverridden ? "cidade-override" : (cidade ? "cidade" : "cidade-empty");
        const cidadeBadge = isOverridden
          ? '<span class="override-badge" title="Editado">✎</span>'
          : cidade
            ? '<span class="auto-badge" title=""></span>'
            : "";

        const monthName = (() => {
          if (!row.data) return "—";
          const dt = row.data instanceof Date ? row.data : new Date(row.data);
          return MONTH_NAMES[dt.getMonth()]?.toLowerCase() || "—";
        })();

        return `<tr data-lancamento-id="${lId}">
          <td>${formatDate(row.data)}</td>
          <td>${monthName}</td>
          <td>${row.historico}</td>
          <td class="td-center">${row.filial}</td>
          <td style="text-align:right;font-weight:600">${fmt(row.debito)}</td>
          <td><span class="tag ${tagClass(row.categoria)}">${row.categoria || "—"}</span></td>
          <td class="td-center">
            <div class="cidade-cell ${cidadeClass}" data-lancamento-id="${lId}">
              <span class="cidade-text">${cidade || "—"}</span>
              ${cidadeBadge}
            </div>
          </td>
        </tr>`;
      }).join("");

      tbody.querySelectorAll(".cidade-cell").forEach(cell => {
        cell.addEventListener("click", () => editCidade(cell));
      });
    }

    const totalEl = document.getElementById("lancTotal");
    if (totalEl) totalEl.textContent = `${rows.length} lancamentos · ${fmt(total)}`;

    document.querySelectorAll("#lancTable .sort-header").forEach(th => {
      const col = th.dataset.sort;
      const ind = th.querySelector(".sort-indicator");
      if (!ind) return;
      if (col === sortCol) {
        th.classList.add("active");
        ind.textContent = sortDir === "asc" ? "▲" : "▼";
      } else {
        th.classList.remove("active");
        ind.textContent = "↕";
      }
    });
  }

  container.innerHTML = `
    <div class="lancamentos-page">
      <div class="page-header">
        <h2>Lancamentos</h2>
        <span id="lancTotal" style="font-size:13px;color:var(--muted)"></span>
      </div>
      <div class="lanc-filters">
        <select id="filterMonth" class="lanc-filter-select">
          <option value="all">Todos os meses</option>
          ${MONTH_NAMES.map(m => `<option value="${m}">${m.charAt(0)+m.slice(1).toLowerCase()}</option>`).join("")}
        </select>
        <select id="filterCategory" class="lanc-filter-select">
          <option value="all">Todas as categorias</option>
          ${allCategories.map(c => `<option value="${c}">${c}</option>`).join("")}
        </select>
        <select id="filterCidade" class="lanc-filter-select">
          <option value="all">Todas as cidades</option>
          ${allCidades.map(c => `<option value="${c}">${c}</option>`).join("")}
        </select>
        <input type="text" id="lancSearch" style="margin-bottom:10px"; class="lanc-filter-input" placeholder="Buscar por historico, filial...">
      </div>
      <div class="table-wrap">
        <table id="lancTable" class="modal-table">
          <thead>
            <tr>
              <th class="sort-header" data-sort="data">Data <span class="sort-indicator">↕</span></th>
              <th>Mes</th>
              <th>Historico</th>
              <th>Filial</th>
              <th class="sort-header" data-sort="" style="text-align:right"> (R$) <span class="sort-indicator">↕</span></th>
              <th>Categoria</th>
              <th>Cidade</th>
            </tr>
          </thead>
          <tbody id="lancTableBody"></tbody>
        </table>
      </div>
    </div>
  `;

  document.getElementById("filterMonth")?.addEventListener("change", e => { filterMonth = e.target.value; renderTable(); });
  document.getElementById("filterCategory")?.addEventListener("change", e => { filterCategory = e.target.value; renderTable(); });
  document.getElementById("filterCidade")?.addEventListener("change", e => { filterCidade = e.target.value; renderTable(); });
  document.getElementById("lancSearch")?.addEventListener("input", e => { searchTerm = e.target.value; renderTable(); });

  document.querySelectorAll("#lancTable .sort-header").forEach(th => {
    th.onclick = () => {
      const col = th.dataset.sort;
      if (sortCol === col) { sortDir = sortDir === "asc" ? "desc" : "asc"; }
      else { sortCol = col; sortDir = "desc"; }
      renderTable();
    };
  });

  renderTable();
}

// ─── MÁSCARAS DE INPUT
function maskCNPJ(value) {
  const digits = value.replace(/\D/g, "").substring(0, 14);
  let masked = digits;
  if (digits.length > 2) masked = digits.slice(0, 2) + "." + digits.slice(2);
  if (digits.length > 5) masked = digits.slice(0, 2) + "." + digits.slice(2, 5) + "." + digits.slice(5);
  if (digits.length > 8) masked = digits.slice(0, 2) + "." + digits.slice(2, 5) + "." + digits.slice(5, 8) + "/" + digits.slice(8);
  if (digits.length > 12) masked = digits.slice(0, 2) + "." + digits.slice(2, 5) + "." + digits.slice(5, 8) + "/" + digits.slice(8, 12) + "-" + digits.slice(12);
  return masked;
}

function maskPhone(value) {
  const digits = value.replace(/\D/g, "").substring(0, 11);
  if (digits.length === 0) return "";
  if (digits.length <= 2) return `(${digits}`;
  if (digits.length <= 6) return `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
  if (digits.length <= 10) return `(${digits.slice(0, 2)}) ${digits.slice(2, 6)}-${digits.slice(6)}`;
  return `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
}

function hasLetters(value) {
  return /[a-zA-Z]/.test(value);
}

// ─── VIEW: FORNECEDORES
export function renderFornecedoresView(container) {
  let searchTerm = "";
  let editingId = null;

  function getFilteredList() {
    let list = [...fornecedoresData];
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      list = list.filter(f =>
        String(f.fornecedor || "").toLowerCase().includes(q) ||
        String(f.cnpj || "").toLowerCase().includes(q) ||
        String(f.cidade || "").toLowerCase().includes(q) ||
        String(f.razao_social || "").toLowerCase().includes(q) ||
        String(f.nome_fantasia || "").toLowerCase().includes(q) ||
        String(f.tipo_midia || "").toLowerCase().includes(q) ||
        String(f.contato || "").toLowerCase().includes(q)
      );
    }
    list.sort((a, b) => String(a.fornecedor || "").localeCompare(String(b.fornecedor || "")));
    return list;
  }

  function renderTable() {
    const list = getFilteredList();
    const tbody = document.getElementById("fornTableBody");
    if (!tbody) return;

    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--muted)">Nenhum fornecedor encontrado.</td></tr>';
    } else {
      tbody.innerHTML = list.map(f => {
        const editable = f.source === "firestore" && f.id;
        const actions = editable
          ? `<button class="forn-edit" data-id="${f.id}" title="Editar">✎</button>
             <button class="forn-delete" data-id="${f.id}" title="Excluir">🗑</button>`
          : '<span style="color:var(--muted);font-size:11px">—</span>';
        return `<tr>
          <td><strong>${f.fornecedor || "—"}</strong></td>
          <td style="font-size:12px">${f.cnpj ? maskCNPJ(f.cnpj) : "—"}</td>
          <td>${f.cidade || "—"}</td>
          <td style="font-size:12px">${f.razao_social || "—"}</td>
          <td>${f.nome_fantasia || "—"}</td>
          <td>${f.tipo_midia
          ? f.tipo_midia.split(",").map(t => `<span class="tag tag-out">${t.trim()}</span>`).join(" ")
          : "—"}</td>
          <td style="font-size:12px">${f.contato || "—"}</td>
          <td class="td-center">${actions}</td>
        </tr>`;
      }).join("");
    }

    // Atualiza contador
    const counter = document.getElementById("fornCount");
    if (counter) counter.textContent = `${list.length} cadastrados`;

    // Re-binda edit/delete nos novos botões
    bindRowActions();
  }

  function bindRowActions() {
    document.querySelectorAll(".forn-edit").forEach(btn => {
      btn.onclick = () => {
        const f = fornecedoresData.find(x => x.id === btn.dataset.id);
        if (!f) return;
        editingId = f.id;
        document.getElementById("fornFormTitle").textContent = "Editar fornecedor";
        document.getElementById("fornInputName").value = f.fornecedor || "";
        document.getElementById("fornInputCnpj").value = maskCNPJ(f.cnpj || "");
        document.getElementById("fornInputCidade").value = f.cidade || "";
        document.getElementById("fornInputRazao").value = f.razao_social || "";
        document.getElementById("fornInputFantasia").value = f.nome_fantasia || "";
        document.getElementById("fornInputMidia").value = f.tipo_midia || "";
        document.getElementById("fornInputContato").value = maskPhone(f.contato || "");
        document.getElementById("fornSubmitBtn").textContent = "Salvar";
        document.getElementById("fornCancelEdit").style.display = "inline-block";
        document.getElementById("fornInputName").focus();
      };
    });

    document.querySelectorAll(".forn-delete").forEach(btn => {
      btn.onclick = async () => {
        const f = fornecedoresData.find(x => x.id === btn.dataset.id);
        if (!f) return;
        if (!confirm(`Excluir "${f.fornecedor}"?`)) return;
        try {
          await removeFornecedor(f.id);
        } catch (err) {
          alert("Erro ao excluir: " + err.message);
        }
      };
    });
  }

  function renderPage() {
    const list = getFilteredList();
    const cidades = [...new Set(fornecedoresData.map(f => f.cidade).filter(Boolean))].sort();

    container.innerHTML = `
      <div class="fornecedores-page">
        <div class="page-header">
          <h2>Fornecedores</h2>
                    <span style="display: inline-block; margin-bottom: 12px; font-size: 13px; color: var(--muted)">
            ${list.length} cadastrados
          </span>
        </div>

        <div class="forn-form" id="fornForm">
          <h3 id="fornFormTitle">Adicionar fornecedor</h3>
          <div class="forn-form-grid">
            <div class="forn-field">
              <label>Nome do fornecedor *</label>
              <input type="text" id="fornInputName" placeholder="Ex: Gráfica Silva">
            </div>
            <div class="forn-field">
              <label>CNPJ</label>
              <input type="text" id="fornInputCnpj" placeholder="00.000.000/0000-00" maxlength="18">
              <span class="field-error" id="cnpjError"></span>
            </div>
            <div class="forn-field">
              <label>Cidade</label>
              <input type="text" id="fornInputCidade" placeholder="Ex: Presidente Prudente" list="fornCidadesOpt">
              <datalist id="fornCidadesOpt">
                ${cidades.map(c => `<option value="${c}">`).join("")}
              </datalist>
            </div>
            <div class="forn-field">
              <label>Razão Social</label>
              <input type="text" id="fornInputRazao" placeholder="Ex: Gráfica Silva Ltda">
            </div>
            <div class="forn-field">
              <label>Nome Fantasia</label>
              <input type="text" id="fornInputFantasia" placeholder="Ex: Silva Gráfica">
            </div>
            <div class="forn-field">
              <label>Tipo de Mídia / Produto</label>
              <input type="text" id="fornInputMidia" placeholder="Ex: Rádio, Outdoor, Gráfica">
            </div>
            <div class="forn-field">
              <label>Contato (Telefone)</label>
              <input type="text" id="fornInputContato" placeholder="(00) 00000-0000" maxlength="15">
            </div>
          </div>
          <div class="forn-form-actions">
            <button class="auth-btn" id="fornSubmitBtn" style="margin:0">Adicionar</button>
            <button class="forn-cancel-edit" id="fornCancelEdit" style="display:none">Cancelar</button>
          </div>
        </div>

        <div class="lanc-filters">
          <input type="text" style="margin-bottom:10px; id="fornSearch" class="lanc-filter-input" placeholder="Buscar fornecedor">
        </div>

        <div class="table-wrap">
          <table class="modal-table" id="fornTable">
            <thead>
              <tr>
                <th>Fornecedor</th>
                <th>CNPJ</th>
                <th>Cidade</th>
                <th>Razão Social</th>
                <th>Nome Fantasia</th>
                <th>Tipo de Mídia</th>
                <th>Contato</th>
                <th style="text-align:center">Ações</th>
              </tr>
            </thead>
            <tbody id="fornTableBody"></tbody>
          </table>
        </div>
      </div>
    `;

    // Busca — só atualiza o tbody, não destrói o input
    document.getElementById("fornSearch")?.addEventListener("input", e => {
      searchTerm = e.target.value;
      renderTable();
    });

    // Máscaras
    const cnpjInput = document.getElementById("fornInputCnpj");
    const cnpjError = document.getElementById("cnpjError");

    cnpjInput?.addEventListener("input", (e) => {
      const raw = e.target.value;
      if (hasLetters(raw)) {
        cnpjError.textContent = "Apenas números";
        cnpjError.classList.add("show");
        cnpjInput.classList.add("input-error");
        e.target.value = raw.replace(/[a-zA-Z]/g, "");
        setTimeout(() => {
          cnpjError.classList.remove("show");
          cnpjInput.classList.remove("input-error");
        }, 2000);
        return;
      }
      cnpjError.classList.remove("show");
      cnpjInput.classList.remove("input-error");
      e.target.value = maskCNPJ(raw);
    });

    cnpjInput?.addEventListener("paste", (e) => {
      e.preventDefault();
      const pasted = (e.clipboardData || window.clipboardData).getData("text");
      const digits = pasted.replace(/\D/g, "").substring(0, 14);
      e.target.value = maskCNPJ(digits);
    });

    const phoneInput = document.getElementById("fornInputContato");

    phoneInput?.addEventListener("input", (e) => {
      if (hasLetters(e.target.value)) {
        e.target.value = e.target.value.replace(/[a-zA-Z]/g, "");
        return;
      }
      e.target.value = maskPhone(e.target.value);
    });

    phoneInput?.addEventListener("paste", (e) => {
      e.preventDefault();
      const pasted = (e.clipboardData || window.clipboardData).getData("text");
      const digits = pasted.replace(/\D/g, "").substring(0, 11);
      e.target.value = maskPhone(digits);
    });

    // Submit
    document.getElementById("fornSubmitBtn")?.addEventListener("click", async () => {
      const data = {
        fornecedor: document.getElementById("fornInputName").value.trim(),
        cnpj: document.getElementById("fornInputCnpj").value.replace(/\D/g, ""),
        cidade: document.getElementById("fornInputCidade").value.trim(),
        razao_social: document.getElementById("fornInputRazao").value.trim(),
        nome_fantasia: document.getElementById("fornInputFantasia").value.trim(),
        tipo_midia: document.getElementById("fornInputMidia").value.trim(),
        contato: document.getElementById("fornInputContato").value.trim()
      };

      if (!data.fornecedor) { alert("Nome do fornecedor é obrigatório."); return; }

      const btn = document.getElementById("fornSubmitBtn");
      btn.classList.add("loading");
      btn.disabled = true;
      try {
        if (editingId) {
          await updateFornecedor(editingId, data);
          editingId = null;
          document.getElementById("fornFormTitle").textContent = "Adicionar fornecedor";
          document.getElementById("fornSubmitBtn").textContent = "Adicionar";
          document.getElementById("fornCancelEdit").style.display = "none";
          // Limpa form
          document.getElementById("fornInputName").value = "";
          document.getElementById("fornInputCnpj").value = "";
          document.getElementById("fornInputCidade").value = "";
          document.getElementById("fornInputRazao").value = "";
          document.getElementById("fornInputFantasia").value = "";
          document.getElementById("fornInputMidia").value = "";
          document.getElementById("fornInputContato").value = "";
        } else {
          await addFornecedor(data);
          // Limpa form
          document.getElementById("fornInputName").value = "";
          document.getElementById("fornInputCnpj").value = "";
          document.getElementById("fornInputCidade").value = "";
          document.getElementById("fornInputRazao").value = "";
          document.getElementById("fornInputFantasia").value = "";
          document.getElementById("fornInputMidia").value = "";
          document.getElementById("fornInputContato").value = "";
        }
      } catch (err) {
        alert("Erro ao salvar: " + err.message);
      } finally {
        btn.classList.remove("loading");
        btn.disabled = false;
      }
    });

    // Cancel edit
    document.getElementById("fornCancelEdit")?.addEventListener("click", () => {
      editingId = null;
      document.getElementById("fornFormTitle").textContent = "Adicionar fornecedor";
      document.getElementById("fornSubmitBtn").textContent = "Adicionar";
      document.getElementById("fornCancelEdit").style.display = "none";
      document.getElementById("fornInputName").value = "";
      document.getElementById("fornInputCnpj").value = "";
      document.getElementById("fornInputCidade").value = "";
      document.getElementById("fornInputRazao").value = "";
      document.getElementById("fornInputFantasia").value = "";
      document.getElementById("fornInputMidia").value = "";
      document.getElementById("fornInputContato").value = "";
    });

    // Renderiza a tabela pela primeira vez
    renderTable();
  }

  renderPage();
}

// ─── RENDER DASHBOARD ────────────────────────────────────────
function renderEmptyState() {
  document.getElementById("main-content").innerHTML = `
    <div class="empty-state">
      <div class="empty-icon">📊</div>
      <h2>Nenhuma planilha cadastrada</h2>
      <p>Aguarde o administrador subir a planilha do mes.<br>
      <span style="color:var(--green);font-size:12px;margin-top:6px;display:block">Os dados aparecerão automaticamente para todos.</span></p>
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
        <div class="kpi-label">Orcado Anual</div>
        <div class="kpi-value">${fmtShort(orcadoAnual)}</div>
        <div class="kpi-sub">Total previsto 2026</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Realizado Acumulado</div>
        <div class="kpi-value ${pctGeral > 100 ? "red" : "green"}">${fmtShort(totalRealizado)}</div>
        <div class="kpi-sub">${realizados.length} ${realizados.length === 1 ? "mes" : "meses"} realizados</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Orcamento Restante</div>
        <div class="kpi-value ${kpiClass}">${fmtShort(restante)}</div>
        <div class="progress-bar-wrap" style="margin-top:8px"><div class="progress-bar" style="width:${barW}%;background:${barColor}"></div></div>
        <div class="kpi-sub ${kpiClass}">${pctRestante.toFixed(1)}% restante do orcamento</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Saldo (orcado - realizado)</div>
        <div class="kpi-value ${saldo < 0 ? "red" : "green"}">${fmtShort(Math.abs(saldo))}</div>
        <div class="kpi-sub">${saldo < 0 ? "▲ Acima do orcado" : "▼ Abaixo do orcado"}</div>
      </div>
    </div>
    <div class="section-title">Meses - clique para ver lancamentos de NF</div>
    <div class="months-grid">${months.map(m => {
      const pct = m.orcado ? (m.realizado / m.orcado) * 100 : 0;
      const bc = pct > 110 ? "var(--red)" : pct > 100 ? "var(--yellow)" : "var(--green)";
      const bw = Math.min(pct, 100).toFixed(1);
      if (!m.hasData) return `<div class="month-card no-data">
        <div class="month-name">${m.month.toLowerCase()}</div>
        <div class="month-val" style="color:var(--muted)">${fmtShort(m.orcado)}</div>
        <div class="month-val-label">orcado</div>
        <div class="badge-future">Sem lancamentos</div>
      </div>`;
      return `<div class="month-card" data-month="${m.month}">
        <div class="month-name">${m.month.toLowerCase()}</div>
        <div class="month-val" style="color:${bc}">${fmtShort(m.realizado)}</div>
        <div class="month-val-label">realizado de ${fmtShort(m.orcado)}</div>
        <div class="progress-bar-wrap"><div class="progress-bar" style="width:${bw}%;background:${bc}"></div></div>
        <div class="month-pct" style="color:${bc}">${pct.toFixed(1)}%</div>
      </div>`;
    }).join("")}</div>
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
    <div class="table-header"><span>Distribuicao de Gastos Realizados</span><span style="font-size:12px;color:var(--muted)">Total: ${fmt(total)}</span></div>
    <table><thead><tr><th>Categoria</th><th style="text-align:right">Realizado (R$)</th><th style="text-align:right">Participacao</th></tr></thead>
    <tbody>${rows}</tbody></table>
  </div>`;
}

// ─── MODAL ───────────────────────────────────────────────────
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
    `Lancamentos — ${monthName.charAt(0) + monthName.slice(1).toLowerCase()}`;
  document.getElementById("modalSub").textContent =
    `${currentModalRows.length} NFs · Total: ${fmt(total)}`;

  sortState = { column: 'data', direction: 'desc' };
  applySortAndRender();
  document.getElementById("modalOverlay").classList.add("open");

  const search = document.getElementById("modalSearch");
  search.value = "";
  search.oninput = e => {
    const q = e.target.value.toLowerCase();
    const filtered = q ? currentModalRows.filter(l =>
      l.historico.toLowerCase().includes(q) ||
      (l.categoria || "").toLowerCase().includes(q) ||
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
        (l.categoria || "").toLowerCase().includes(q) ||
        String(l.filial).toLowerCase().includes(q)
      ) : [...currentModalRows];
      applySortAndRender(rows);
    };
  });
}

function renderModalTable(rows) {
  const tbody = document.getElementById("modalTableBody");
  if (!tbody) return;
  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6">Nenhum lancamento encontrado.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(row => {
    const lId = lancamentoId(row, currentModalMonth);
    const override = getOverride(lId);
    let cidade = "";
    let isOverridden = false;

    if (override) {
      cidade = override.cidade || "";
      isOverridden = true;
    } else if (fornecedoresData.length > 0) {
      const match = matchFornecedor(row.historico, fornecedoresData);
      if (match) cidade = match.cidade || "";
    }

    const cidadeClass = isOverridden ? "cidade-override" : (cidade ? "cidade-auto" : "cidade-empty");
    const cidadeBadge = isOverridden
      ? '<span class="override-badge" title="Editado pelo usuario">✎</span>'
      : cidade
        ? '<span class="auto-badge" title=""></span>'
        : "";

    return `<tr data-lancamento-id="${lId}">
      <td>${formatDate(row.data)}</td>
      <td>${row.historico}</td>
      <td class="td-center">${row.filial}</td>
      <td style="text-align:right;font-weight:600">${fmt(row.debito)}</td>
      <td><span class="tag ${tagClass(row.categoria)}">${row.categoria || "—"}</span></td>
      <td class="td-center">
        <div class="cidade-cell ${cidadeClass}" data-lancamento-id="${lId}">
          <span class="cidade-text">${cidade || "—"}</span>
          ${cidadeBadge}
        </div>
      </td>
    </tr>`;
  }).join("");

  tbody.querySelectorAll(".cidade-cell").forEach(cell => {
    cell.addEventListener("click", () => editCidade(cell));
  });
}

function editCidade(cell) {
  const lId = cell.dataset.lancamentoId;
  const textEl = cell.querySelector(".cidade-text");
  const currentText = textEl.textContent === "—" ? "" : textEl.textContent;

  const cidadesConhecidas = getCidadesConhecidas(fornecedoresData);

  cell.innerHTML = `
    <div class="cidade-edit-wrap">
      <input type="text" class="cidade-input" value="${currentText}"
        list="cidadesListModal" placeholder="Digite a cidade...">
      <datalist id="cidadesListModal">
        ${cidadesConhecidas.map(c => `<option value="${c}">`).join("")}
      </datalist>
      <button class="cidade-save" title="Salvar">✓</button>
      <button class="cidade-cancel" title="Cancelar">✕</button>
    </div>
  `;

  const input = cell.querySelector(".cidade-input");
  input.focus();
  input.select();

  cell.querySelector(".cidade-save").onclick = async () => {
    const newCidade = input.value.trim();
    if (newCidade) {
      await saveOverride(lId, { cidade: newCidade });
    } else {
      await removeOverride(lId);
    }
  };

  cell.querySelector(".cidade-cancel").onclick = () => {
    const search = document.getElementById("modalSearch");
    const q = search?.value.toLowerCase() || '';
    const rows = q ? currentModalRows.filter(l =>
      l.historico.toLowerCase().includes(q) ||
      (l.categoria || "").toLowerCase().includes(q) ||
      String(l.filial).toLowerCase().includes(q)
    ) : [...currentModalRows];
    applySortAndRender(rows);
  };

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") cell.querySelector(".cidade-save").click();
    if (e.key === "Escape") cell.querySelector(".cidade-cancel").click();
  });
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
    if (currentModalMonth) captureToClipboard("Lancamentos — " + currentModalMonth);
  });
}

// ─── HELPERS ─────────────────────────────────────────────────
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