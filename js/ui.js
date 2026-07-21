// ─── TEMA ───────────────────────────────────────────────────────
export function initTheme() {
  const saved = localStorage.getItem('dashboard_theme');
  if (saved === 'light') document.body.classList.add('light-mode');
}

export function setupThemeToggle() {
  const btn = document.getElementById('themeToggle');
  if (!btn) return;
  btn.addEventListener('click', () => {
    document.body.classList.toggle('light-mode');
    localStorage.setItem('dashboard_theme',
      document.body.classList.contains('light-mode') ? 'light' : 'dark');
  });
}

// ─── LOADING GLOBAL ─────────────────────────────────────────────
export function showGlobalLoading(msg = "Carregando...") {
  const el = document.getElementById('globalLoading');
  if (!el) return;
  el.querySelector('p').textContent = msg;
  el.classList.add('show');
}

export function hideGlobalLoading() {
  const el = document.getElementById('globalLoading');
  if (!el) return;
  el.classList.remove('show');
}

// ─── CAPTURA DE TELA ───────────────────────────────────────────
export async function captureToClipboard(label) {
  const loading = document.getElementById('pdfLoading');
  const msg = document.getElementById('pdfLoadingMsg');
  if (!loading || !msg) return;
  try {
    msg.textContent = `Capturando ${label || 'tela'}...`;
    loading.classList.add('show');
    await new Promise(r => setTimeout(r, 50));
    const target = document.getElementById('main-content');
    const canvas = await html2canvas(target, {
      scale: window.devicePixelRatio || 1,
      useCORS: true,
      backgroundColor: document.body.classList.contains('light-mode') ? '#f5f6fa' : '#0f1117'
    });
    const blob = await new Promise((res, rej) => {
      canvas.toBlob(b => b ? res(b) : rej(new Error('Falha')), 'image/png');
    });
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
    msg.textContent = 'Imagem copiada! Agora é só colar (Ctrl+V)';
    await new Promise(r => setTimeout(r, 2000));
  } catch (err) {
    console.error(err);
    alert('Erro ao capturar: ' + err.message);
  } finally {
    loading.classList.remove('show');
  }
}

// ─── HELPERS ───────────────────────────────────────────────────
export function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number' && typeof XLSX !== 'undefined') return XLSX.SSF.parse_date_code(val);
  return new Date(val);
}

export function formatDate(d) {
  if (!d) return '';
  try { return (d instanceof Date ? d : new Date(d)).toLocaleDateString('pt-BR'); }
  catch { return ''; }
}

export function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

export function fmt(n) {
  return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 2 });
}

export function fmtShort(n) {
  if (n >= 1e6) return 'R$ ' + (n/1e6).toFixed(2).replace('.',',') + 'M';
  if (n >= 1e3) return 'R$ ' + (n/1e3).toFixed(1).replace('.',',') + 'K';
  return fmt(n);
}

export function tagClass(cat) {
  if (!cat) return 'tag-out';
  const c = cat.toLowerCase();
  if (c.includes('publicidade') || c.includes('propaganda')) return 'tag-pub';
  if (c.includes('brinde')) return 'tag-bri';
  if (c.includes('patroc')) return 'tag-pat';
  if (c.includes('doa')) return 'tag-doa';
  return 'tag-out';
}