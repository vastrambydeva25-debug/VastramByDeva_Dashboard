// ============================================
// 🔥 Shared Firebase Configuration
// ============================================
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyAY_oTerSMIF8FYoKyyFIokX-jadXqEtb0",
  authDomain: "vastrambydeva-fa343.firebaseapp.com",
  projectId: "vastrambydeva-fa343",
  storageBucket: "vastrambydeva-fa343.firebasestorage.app",
  messagingSenderId: "988690580826",
  appId: "1:988690580826:web:b447539adcca2032d89f63"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth(app);

export { app, db, auth };
