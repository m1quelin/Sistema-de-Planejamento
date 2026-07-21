import { loginWithEmail, registerWithEmail } from './auth.js';

// Estado da UI
let currentView = 'login'; // 'login' | 'register'

export function initAuthUI() {
  bindToggles();
  bindForms();
  bindRealtimeValidation();
}

// ─── TOGGLE ENTRE VIEWS COM ANIMAÇÃO ─────────────────────────
function bindToggles() {
  const toRegisterBtn = document.getElementById('toRegister');
  const toLoginBtn = document.getElementById('toLogin');

  toRegisterBtn?.addEventListener('click', () => switchView('register'));
  toLoginBtn?.addEventListener('click', () => switchView('login'));
}

function switchView(target) {
  if (currentView === target) return;

  const loginView = document.getElementById('loginView');
  const registerView = document.getElementById('registerView');
  const authBox = document.querySelector('.auth-box');

  // Animação de saída
  authBox.style.transform = 'scale(0.98)';
  authBox.style.opacity = '0.7';

  setTimeout(() => {
    if (target === 'register') {
      loginView.style.display = 'none';
      registerView.style.display = 'block';
      document.getElementById('tabLogin')?.classList.remove('active');
      document.getElementById('tabRegister')?.classList.add('active');
    } else {
      registerView.style.display = 'none';
      loginView.style.display = 'block';
      document.getElementById('tabRegister')?.classList.remove('active');
      document.getElementById('tabLogin')?.classList.add('active');
    }

    // Limpa erros ao trocar
    clearErrors();

    // Animação de entrada
    setTimeout(() => {
      authBox.style.transform = 'scale(1)';
      authBox.style.opacity = '1';
    }, 30);

    // Foco no primeiro campo
    const firstInput = target === 'register'
      ? document.getElementById('regName')
      : document.getElementById('loginEmail');
    firstInput?.focus();

    currentView = target;
  }, 150);
}

// ─── FORMULÁRIOS ─────────────────────────────────────────────
function bindForms() {
  const loginForm = document.getElementById('loginForm');
  const registerForm = document.getElementById('registerForm');

  loginForm?.addEventListener('submit', (e) => {
    e.preventDefault();
    const email = document.getElementById('loginEmail').value.trim();
    const password = document.getElementById('loginPassword').value;
    loginWithEmail(email, password, {
      btn: document.getElementById('loginBtn'),
      errorEl: document.getElementById('loginError')
    }).catch(() => {});
  });

  registerForm?.addEventListener('submit', (e) => {
    e.preventDefault();
    const name = document.getElementById('regName').value.trim();
    const email = document.getElementById('regEmail').value.trim();
    const password = document.getElementById('regPassword').value;
    registerWithEmail(email, password, name, {
      btn: document.getElementById('registerBtn'),
      errorEl: document.getElementById('registerError')
    }).catch(() => {});
  });
}

// ─── VALIDAÇÃO EM TEMPO REAL DO DOMÍNIO ──────────────────────
function bindRealtimeValidation() {
  const loginEmail = document.getElementById('loginEmail');
  const regEmail = document.getElementById('regEmail');

  loginEmail?.addEventListener('input', () => validateDomain(loginEmail, 'loginDomainHint'));
  loginEmail?.addEventListener('blur', () => validateDomain(loginEmail, 'loginDomainHint'));

  regEmail?.addEventListener('input', () => validateDomain(regEmail, 'regDomainHint'));
  regEmail?.addEventListener('blur', () => validateDomain(regEmail, 'regDomainHint'));
}

function validateDomain(input, hintId) {
  const hint = document.getElementById(hintId);
  if (!input || !hint) return;

  const email = input.value.trim();
  if (!email) {
    input.classList.remove('invalid', 'valid');
    hint.classList.remove('show', 'success');
    return;
  }

  const isValid = email.toLowerCase().endsWith('@oestesaude.com.br');

  if (isValid) {
    input.classList.remove('invalid');
    input.classList.add('valid');
    hint.textContent = '✓ Domínio autorizado';
    hint.className = 'domain-hint success show';
  } else {
    input.classList.remove('valid');
    input.classList.add('invalid');
    hint.textContent = '✗ Apenas @oestesaude.com.br';
    hint.className = 'domain-hint error show';
  }
}

function clearErrors() {
  ['loginError', 'registerError'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.textContent = '';
      el.classList.remove('show');
    }
  });
}