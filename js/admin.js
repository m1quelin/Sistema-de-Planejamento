import { db } from "./firebase-config.js";
import { doc, setDoc, onSnapshot } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-firestore.js";

const PLANILHA_DOC = doc(db, "config", "planilha");

export function listenPlanilhaMetadata(callback) {
  return onSnapshot(PLANILHA_DOC, (snap) => {
    callback(snap.exists() ? snap.data() : null);
  });
}

export async function uploadPlanilha(file, user) {
  if (!file) throw new Error("Nenhum arquivo selecionado.");
  if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
    throw new Error("Apenas arquivos .xlsx ou .xls são permitidos.");
  }

  // Lê como base64
  const base64 = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = () => reject(new Error("Erro ao ler arquivo"));
    reader.readAsDataURL(file);
  });

  // Salva no Firestore
  await setDoc(PLANILHA_DOC, {
    dataBase64: base64,
    filename: file.name,
    updatedAt: new Date().toISOString(),
    updatedBy: user?.displayName || user?.email || "anon",
    updatedByUid: user?.uid || "anon",
    sizeBytes: file.size
  });

  return true;
}