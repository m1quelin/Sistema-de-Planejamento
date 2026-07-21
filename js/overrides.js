import { db } from "./firebase-config.js";
import {
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot
} from "https://www.gstatic.com/firebasejs/10.12.4/firebase-firestore.js";

const OVERRIDES_COL = collection(db, "overrides");

let overridesCache = {};
let unsubOverrides = null;

// Escuta mudanças em tempo real
export function listenOverrides(callback) {
  if (unsubOverrides) unsubOverrides();
  unsubOverrides = onSnapshot(OVERRIDES_COL, (snap) => {
    overridesCache = {};
    snap.forEach(doc => {
      overridesCache[doc.id] = doc.data();
    });
    if (callback) callback(overridesCache);
  });
}

// Busca override de um lançamento específico
export function getOverride(lancamentoId) {
  return overridesCache[lancamentoId] || null;
}

// Salva override
export async function saveOverride(id, data) {
  await setDoc(doc(db, "overrides", id), {
    ...data,
    updatedAt: new Date().toISOString()
  });
}

// Remove override (volta pro auto-preenchido)
export async function removeOverride(id) {
  await deleteDoc(doc(db, "overrides", id));
}

export function stopListenOverrides() {
  if (unsubOverrides) {
    unsubOverrides();
    unsubOverrides = null;
  }
}