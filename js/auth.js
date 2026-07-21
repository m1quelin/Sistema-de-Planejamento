import { auth, ALLOWED_DOMAIN } from "./firebase-config.js";
import {
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  updateProfile
} from "https://www.gstatic.com/firebasejs/10.12.4/firebase-auth.js";

function isAllowedEmail(email) {
  return email.toLowerCase().endsWith("@" + ALLOWED_DOMAIN);
}

function setLoading(btn, isLoading) {
  if (!btn) return;
  if (isLoading) {
    btn.dataset.originalText = btn.textContent;
    btn.textContent = "";
    btn.classList.add("loading");
    btn.disabled = true;
  } else {
    btn.textContent = btn.dataset.originalText || "";
    btn.classList.remove("loading");
    btn.disabled = false;
  }
}

function showError(el, msg) {
  if (!el) return;
  el.textContent = msg;
  el.classList.add("show");
  console.error("[AUTH]", msg);
}

function hideError(el) {
  if (!el) return;
  el.classList.remove("show");
}

export function loginWithEmail(email, password, opts = {}) {
  const { btn, errorEl } = opts;
  hideError(errorEl);
  if (!isAllowedEmail(email)) {
    showError(errorEl, "Domínio não permitido");
    return Promise.reject(new Error("Domínio não autorizado"));
  }
  setLoading(btn, true);
  return signInWithEmailAndPassword(auth, email, password)
    .catch(err => {
      let msg = "Erro ao entrar. Tente novamente.";
      if (err.code === "auth/invalid-credential") msg = "E-mail ou senha incorretos.";
      if (err.code === "auth/user-not-found") msg = "Usuário não encontrado.";
      if (err.code === "auth/too-many-requests") msg = "Muitas tentativas. Aguarde.";
      showError(errorEl, msg);
      throw err;
    })
    .finally(() => setLoading(btn, false));
}

export function registerWithEmail(email, password, name, opts = {}) {
  const { btn, errorEl } = opts;
  hideError(errorEl);
  if (!isAllowedEmail(email)) {
    showError(errorEl, "Domínio não permitido");
    return Promise.reject(new Error("Domínio não autorizado"));
  }
  setLoading(btn, true);
  return createUserWithEmailAndPassword(auth, email, password)
    .then(cred => {
      return updateProfile(cred.user, { displayName: name })
        .then(() => cred.user)
        .catch(() => cred.user);
    })
    .catch(err => {
      let msg = "Erro ao cadastrar. Tente novamente.";
      if (err.code === "auth/email-already-in-use") msg = "Este e-mail já está cadastrado.";
      if (err.code === "auth/weak-password") msg = "A senha deve ter pelo menos 6 caracteres.";
      if (err.code === "auth/invalid-email") msg = "E-mail inválido.";
      if (err.code === "auth/operation-not-allowed") msg = "Cadastro por e-mail não está habilitado no Firebase.";
      showError(errorEl, msg);
      throw err;
    })
    .finally(() => setLoading(btn, false));
}

export function logoutUser() {
  return signOut(auth);
}

export function onAuthChanged(callback) {
  return onAuthStateChanged(auth, callback);
}

export function getCurrentUser() {
  return auth.currentUser;
}