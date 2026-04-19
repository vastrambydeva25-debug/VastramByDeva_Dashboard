// ============================================
// 🔒 Auth Guard — import on every protected page
// ============================================
import { auth } from "./firebase-config.js";
import { onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

/**
 * Checks if user is logged in.
 * Redirects to index.html if not.
 * Returns a promise that resolves with the user object.
 */
export function requireAuth() {
  return new Promise((resolve) => {
    onAuthStateChanged(auth, (user) => {
      if (!user) {
        window.location.href = "index.html";
      } else {
        // Show user info in nav if element exists
        const el = document.getElementById("userEmail");
        if (el) el.textContent = user.email;
        resolve(user);
      }
    });
  });
}

export async function logout() {
  await signOut(auth);
  window.location.href = "index.html";
}
