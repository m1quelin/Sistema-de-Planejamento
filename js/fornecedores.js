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

  const normalize = (str) =>
    String(str || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[_\-\.;,\/\()]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const h = normalize(historico);

  const STOPWORDS = new Set([
    "e", "de", "da", "do", "das", "dos", "me", "lt", "ltda",
    "sa", "cia", "nf", "ltd", "mei", "sl", "sc"
  ]);

  const hTokens = h.split(" ").filter(t => t.length > 1);
  const hTokenSet = new Set(hTokens);

  let bestMatch = null;
  let bestScore = 0;

  for (const f of fornecedores) {
    const rs = normalize(f.razao_social);
    if (!rs || rs.length < 3) continue;

    // 1. Substring direto (casa mais específico)
    if (h.includes(rs) || rs.includes(h)) {
      const score = rs.length;
      if (score > bestScore) { bestScore = score; bestMatch = f; }
      continue;
    }

    // 2. Token match — lida com histórico truncado
    const rsTokens = rs.split(" ").filter(t => t.length > 1 && !STOPWORDS.has(t));
    if (rsTokens.length === 0) continue;

    let matches = 0;
    for (const t of rsTokens) {
      if (hTokenSet.has(t)) matches++;
    }

    // Precisa de no mínimo 60% dos tokens da razão social + pelo menos 2 acertos
    const pct = matches / rsTokens.length;
    if (pct >= 0.6 && matches >= 2) {
      const score = pct * rs.length;
      if (score > bestScore) { bestScore = score; bestMatch = f; }
    }
  }

  return bestMatch;
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