import { db } from "./firebase-config.js";
import { doc, setDoc, onSnapshot } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-firestore.js";

const PLANILHA_DOC = doc(db, "config", "planilha");

export function listenPlanilhaMetadata(callback) {
  return onSnapshot(PLANILHA_DOC, (snap) => {
    if (snap.exists()) callback(snap.data());
    else callback(null);
  });
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

export async function uploadPlanilha(file, user) {
  if (!file) throw new Error("Nenhum arquivo selecionado.");

  const base64 = await fileToBase64(file);

  await setDoc(PLANILHA_DOC, {
    base64: base64,
    filename: file.name,
    updatedAt: new Date().toISOString(),
    updatedBy: user.displayName || user.email,
    updatedByUid: user.uid
  });

  return true;
}