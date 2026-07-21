import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.12.4/firebase-firestore.js";

export const firebaseConfig = {
  apiKey: "AIzaSyCnGOaJ1zDCM-m0tvlohq8OqziLGnIyCqM",
  authDomain: "login-sistema-mkt.firebaseapp.com",
  projectId: "login-sistema-mkt",
  storageBucket: "login-sistema-mkt.firebasestorage.app",
  messagingSenderId: "136742172158",
  appId: "1:136742172158:web:951e957a247bdf810afcaf",
};

export const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
export const ALLOWED_DOMAIN = "oestesaude.com.br";