import { db } from "./firebase-config.js";
import {
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy
} from "https://www.gstatic.com/firebasejs/10.12.4/firebase-firestore.js";

const FORN_COL = collection(db, "fornecedores");

let firestoreCache = [];
let unsubForn = null;

// ─── LISTEN (tempo real)
export function listenFornecedores(callback) {
  if (unsubForn) unsubForn();
  const q = query(FORN_COL, orderBy("fornecedor", "asc"));
  unsubForn = onSnapshot(q, (snap) => {
    firestoreCache = [];
    snap.forEach(d => firestoreCache.push({ id: d.id, ...d.data(), source: "firestore" }));
    if (callback) callback(firestoreCache);
  });
}

// ─── FETCH (pra compat com o dashboard.js que chama fetchFornecedores)
export async function fetchFornecedores() {
  return firestoreCache;
}

export function stopListenFornecedores() {
  if (unsubForn) { unsubForn(); unsubForn = null; }
}

// ─── CRUD
export async function addFornecedor(data) {
  if (!data.fornecedor) throw new Error("Nome do fornecedor é obrigatório.");
  const id = data.fornecedor.toLowerCase().replace(/[^a-z0-9]/g, "_").substring(0, 50);
  await setDoc(doc(db, "fornecedores", id), {
    fornecedor: data.fornecedor.trim(),
    cnpj: data.cnpj?.trim() || "",
    cidade: data.cidade?.trim() || "",
    razao_social: data.razao_social?.trim() || "",
    nome_fantasia: data.nome_fantasia?.trim() || "",
    tipo_midia: data.tipo_midia?.trim() || "",
    contato: data.contato?.trim() || "",
    createdAt: new Date().toISOString(),
    createdBy: "user"
  });
  return id;
}

export async function updateFornecedor(id, data) {
  await setDoc(doc(db, "fornecedores", id), {
    ...data,
    updatedAt: new Date().toISOString()
  }, { merge: true });
}

export async function removeFornecedor(id) {
  await deleteDoc(doc(db, "fornecedores", id));
}

// ─── MATCH
export function matchFornecedor(historico, fornecedores) {
  if (!historico || !fornecedores) return null;
  const h = historico.toLowerCase();

  let match = fornecedores.find(f =>
    String(f.fornecedor || "").toLowerCase() === h
  );

  if (!match) {
    match = fornecedores.find(f => {
      const fn = String(f.fornecedor || "").toLowerCase();
      return fn && (h.includes(fn) || fn.includes(h));
    });
  }

  return match || null;
}

// ─── ID DE LANÇAMENTO (pra overrides)
export function lancamentoId(lanc, monthName) {
  const data = lanc.data instanceof Date
    ? lanc.data.toISOString().split("T")[0]
    : String(lanc.data || "");
  const historico = String(lanc.historico || "").trim();
  const debito = String(lanc.debito || "");
  return `${monthName}_${data}_${historico}_${debito}`.replace(/\s+/g, "_");
}

// ─── CIDADES CONHECIDAS (autocomplete)
export function getCidadesConhecidas(fornecedores) {
  return [...new Set(
    fornecedores.map(f => f.cidade).filter(Boolean)
  )].sort();
}

// ─── IS FIRESTORE
export function isFirestoreFornecedor(forn) {
  return forn?.source === "firestore" && forn?.id;
}