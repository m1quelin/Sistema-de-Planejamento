import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-app.js";
import {
  getAuth,
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  updateProfile
} from "https://www.gstatic.com/firebasejs/10.12.4/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyCnGOaJ1zDCM-m0tvlohq8OqziLGnIyCqM",
  authDomain: "login-sistema-mkt.firebaseapp.com",
  projectId: "login-sistema-mkt",
  storageBucket: "login-sistema-mkt.firebasestorage.app",
  messagingSenderId: "136742172158",
  appId: "1:136742172158:web:951e957a247bdf810afcaf",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const ALLOWED_DOMAIN = "oestesaude.com.br";

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
  console.error("[AUTH ERROR]", msg);
}

function hideError(el) {
  if (!el) return;
  el.classList.remove("show");
}

export function loginWithEmail(email, password, opts = {}) {
  const { btn, errorEl } = opts;
  hideError(errorEl);
  if (!isAllowedEmail(email)) {
    showError(errorEl, `Apenas e-mails @${ALLOWED_DOMAIN} são permitidos.`);
    return Promise.reject(new Error("Domínio não autorizado"));
  }
  setLoading(btn, true);
  console.log("[LOGIN] Tentando login com:", email);
  return signInWithEmailAndPassword(auth, email, password)
    .then((cred) => {
      console.log("[LOGIN] Sucesso:", cred.user.uid);
      return cred;
    })
    .catch(err => {
      console.error("[LOGIN] Erro:", err.code, err.message);
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
    showError(errorEl, `Apenas e-mails @${ALLOWED_DOMAIN} podem se cadastrar.`);
    return Promise.reject(new Error("Domínio não autorizado"));
  }
  setLoading(btn, true);
  console.log("[REGISTER] Tentando cadastro:", email);
  
  return createUserWithEmailAndPassword(auth, email, password)
    .then(cred => {
      console.log("[REGISTER] Usuário criado no Firebase:", cred.user.uid);
      return updateProfile(cred.user, { displayName: name })
        .then(() => {
          console.log("[REGISTER] Perfil atualizado com nome:", name);
          return cred.user;
        })
        .catch(profileErr => {
          console.warn("[REGISTER] Falha ao atualizar perfil:", profileErr);
          return cred.user;
        });
    })
    .catch(err => {
      console.error("[REGISTER] Erro:", err.code, err.message);
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