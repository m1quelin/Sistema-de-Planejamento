import { initTheme, setupThemeToggle, captureToClipboard } from "./ui.js";
import { loginWithEmail, registerWithEmail, logoutUser, onAuthChanged, getCurrentUser } from "./auth.js";
import { initDashboard } from "./dashboard.js";
import { uploadPlanilha } from "./admin.js";

initTheme();

const authView = document.getElementById("auth-view");
const dashboardView = document.getElementById("dashboard-view");
const globalLoading = document.getElementById("globalLoading");

// ─── TOGGLE LOGIN/REGISTRO ────────────────────────────────────
window.switchToRegister = function () {
  document.getElementById("loginView").style.display = "none";
  document.getElementById("registerView").style.display = "block";
  document.getElementById("tabLogin")?.classList.remove("active");
  document.getElementById("tabRegister")?.classList.add("active");
  clearErrors();
};

window.switchToLogin = function () {
  document.getElementById("registerView").style.display = "none";
  document.getElementById("loginView").style.display = "block";
  document.getElementById("tabRegister")?.classList.remove("active");
  document.getElementById("tabLogin")?.classList.add("active");
  clearErrors();
};

function clearErrors() {
  ["loginError", "registerError"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = ""; el.classList.remove("show"); }
  });
}

// ─── VALIDAÇÃO DE DOMÍNIO ─────────────────────────────────────
function setupDomainValidation(inputId, hintId) {
  const input = document.getElementById(inputId);
  const hint = document.getElementById(hintId);
  if (!input || !hint) return;

  input.addEventListener("input", () => {
    const email = input.value.trim();
    if (!email) {
      input.classList.remove("valid", "invalid");
      hint.classList.remove("show");
      return;
    }
    if (email.toLowerCase().endsWith("@oestesaude.com.br")) {
      input.classList.remove("invalid");
      input.classList.add("valid");
      hint.textContent = "Domínio permitido";
      hint.className = "domain-hint success show";
    } else {
      input.classList.remove("valid");
      input.classList.add("invalid");
      hint.textContent = "Domínio não permitido";
      hint.className = "domain-hint error show";
    }
  });
}

setupDomainValidation("loginEmail", "loginDomainHint");
setupDomainValidation("regEmail", "regDomainHint");

// ─── FORMS ────────────────────────────────────────────────────
document.getElementById("loginForm")?.addEventListener("submit", (e) => {
  e.preventDefault();
  const email = document.getElementById("loginEmail").value.trim();
  const password = document.getElementById("loginPassword").value;
  loginWithEmail(email, password, {
    btn: document.getElementById("loginBtn"),
    errorEl: document.getElementById("loginError")
  }).catch(() => {});
});

document.getElementById("registerForm")?.addEventListener("submit", (e) => {
  e.preventDefault();
  const name = document.getElementById("regName").value.trim();
  const email = document.getElementById("regEmail").value.trim();
  const password = document.getElementById("regPassword").value;
  registerWithEmail(email, password, name, {
    btn: document.getElementById("registerBtn"),
    errorEl: document.getElementById("registerError")
  }).catch(() => {});
});

document.getElementById("logoutBtn")?.addEventListener("click", logoutUser);

// ─── ADMIN PANEL ──────────────────────────────────────────────
const adminBtn = document.getElementById("adminBtn");
const adminOverlay = document.getElementById("adminOverlay");

adminBtn?.addEventListener("click", () => adminOverlay?.classList.add("open"));
document.getElementById("adminClose")?.addEventListener("click", () => {
  adminOverlay?.classList.remove("open");
});
adminOverlay?.addEventListener("click", e => {
  if (e.target === adminOverlay) adminOverlay.classList.remove("open");
});

const adminFileInput = document.getElementById("adminFileInput");
const adminFileLabel = document.getElementById("adminFileLabel");
const adminUploadBtn = document.getElementById("adminUploadBtn");
const adminError = document.getElementById("adminError");

adminFileInput?.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) {
    adminFileLabel.textContent = file.name;
    adminUploadBtn.disabled = false;
  }
});

adminUploadBtn?.addEventListener("click", async () => {
  const file = adminFileInput?.files[0];
  if (!file) return;
  adminUploadBtn.classList.add("loading");
  adminUploadBtn.disabled = true;
  try {
    await uploadPlanilha(file, getCurrentUser());
    adminError?.classList.remove("show");
    adminUploadBtn.textContent = "Enviado!";
    setTimeout(() => {
      adminOverlay?.classList.remove("open");
      adminUploadBtn.textContent = "Enviar e atualizar";
      adminUploadBtn.classList.remove("loading");
      adminUploadBtn.disabled = true;
      adminFileInput.value = "";
      adminFileLabel.textContent = "Selecionar arquivo XLSX";
    }, 1200);
  } catch (err) {
    adminError.textContent = err.message;
    adminError.classList.add("show");
    adminUploadBtn.classList.remove("loading");
    adminUploadBtn.disabled = false;
  }
});

// ─── CAPTURA DE TELA ──────────────────────────────────────────
document.getElementById("pdfBtn")?.addEventListener("click", () => {
  captureToClipboard("Dashboard MKT");
});

// ─── SIDEBAR 
const sidebarAdmin = document.getElementById("sidebarAdmin");
const sidebarLogout = document.getElementById("sidebarLogout");

// Painel Admin pela sidebar
sidebarAdmin?.addEventListener("click", () => {
  document.getElementById("adminOverlay")?.classList.add("open");
});

// Logout pela sidebar
sidebarLogout?.addEventListener("click", () => {
  logoutUser();
});

// Navegação ativa
document.querySelectorAll(".sidebar-item[data-page]").forEach(item => {
  item.addEventListener("click", (e) => {
    e.preventDefault();
    document.querySelectorAll(".sidebar-item").forEach(i => i.classList.remove("active"));
    item.classList.add("active");
    // Por enquanto só Dashboard tem conteúdo real
    // Os outros são placeholders pra futuras páginas
  });
});

// ─── AUTH STATE ───────────────────────────────────────────────
globalLoading?.classList.add("show");

onAuthChanged(user => {
  globalLoading?.classList.remove("show");

  if (user) {
    authView?.classList.add("hidden");
    dashboardView?.classList.add("active");

    const userInfo = document.getElementById("userInfo");
    const userName = document.getElementById("userName");
    if (userName) userName.textContent = user.displayName || user.email;
    if (userInfo) userInfo.style.display = "flex";
    adminBtn?.classList.remove("hidden");

    initDashboard();

    if (!document.getElementById("themeToggle")) {
      const actions = document.querySelector(".header-actions");
      if (actions) {
        const btn = document.createElement("button");
        btn.id = "themeToggle";
        btn.className = "theme-btn";
        btn.title = "Alternar Tema";
        btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>`;
        actions.insertBefore(btn, actions.firstChild);
        setupThemeToggle();
      }
    }
  } else {
    authView?.classList.remove("hidden");
    dashboardView?.classList.remove("active");
    const userInfo = document.getElementById("userInfo");
    if (userInfo) userInfo.style.display = "none";
    adminBtn?.classList.add("hidden");
  }
});

