import { renderDashboardView, renderLancamentosView, renderFornecedoresView, renderCategoriasView } from "./dashboard.js";

const views = {
  dashboard: { label: "Dashboard", render: renderDashboardView },
  lancamentos: { label: "Lançamentos", render: renderLancamentosView },
  fornecedores: { label: "Fornecedores", render: renderFornecedoresView },
  categorias: { label: "Categorias", render: renderCategoriasView },
};

let currentView = "dashboard";

export function navigateTo(viewName) {
  if (!views[viewName]) return;

  currentView = viewName;
  const main = document.getElementById("main-content");
  if (!main) return;

  main.innerHTML = "";
  views[viewName].render(main);

  document.querySelectorAll(".sidebar-item[data-page]").forEach(item => {
    item.classList.toggle("active", item.dataset.page === viewName);
  });
}

export function getCurrentView() {
  return currentView;
}