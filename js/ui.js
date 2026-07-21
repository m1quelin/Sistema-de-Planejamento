// ─── TEMA ─────────────────────────────────────────────────────
export function initTheme() {
  const saved = localStorage.getItem("dashboard_theme");
  if (saved === "light") document.body.classList.add("light-mode");
}

export function setupThemeToggle() {
  const btn = document.getElementById("themeToggle");
  if (!btn) return;
  btn.addEventListener("click", () => {
    document.body.classList.toggle("light-mode");
    localStorage.setItem("dashboard_theme",
      document.body.classList.contains("light-mode") ? "light" : "dark");
  });
}

// ─── LOADING GLOBAL ───────────────────────────────────────────
export function showGlobalLoading(msg = "Carregando...") {
  const el = document.getElementById("globalLoading");
  if (!el) return;
  el.querySelector("p").textContent = msg;
  el.classList.add("show");
}

export function hideGlobalLoading() {
  const el = document.getElementById("globalLoading");
  if (!el) return;
  el.classList.remove("show");
}

// ─── CAPTURA DE TELA ──────────────────────────────────────────
export async function captureToClipboard(label) {
  const loading = document.getElementById("pdfLoading");
  const msg = document.getElementById("pdfLoadingMsg");
  if (!loading || !msg) return;
  try {
    msg.textContent = `Capturando ${label || "tela"}...`;
    loading.classList.add("show");
    await new Promise(r => setTimeout(r, 50));
    const target = document.getElementById("main-content");
    const canvas = await html2canvas(target, {
      scale: window.devicePixelRatio || 1,
      useCORS: true,
      backgroundColor: document.body.classList.contains("light-mode")
        ? "#f5f6fa" : "#0f1117"
    });
    const blob = await new Promise((res, rej) => {
      canvas.toBlob(b => b ? res(b) : rej(new Error("Falha")), "image/png");
    });
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    msg.textContent = "Imagem copiada! Agora é só colar (Ctrl+V)";
    await new Promise(r => setTimeout(r, 2000));
  } catch (err) {
    alert("Erro ao capturar: " + err.message);
  } finally {
    loading.classList.remove("show");
  }
}