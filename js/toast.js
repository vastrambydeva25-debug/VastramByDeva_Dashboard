// ============================================
// 🔔 Toast Notification Utility
// ============================================

/**
 * Show a toast notification
 * @param {string} message
 * @param {"success"|"error"|"info"} type
 * @param {number} duration ms
 */
export function showToast(message, type = "success", duration = 3000) {
  // Remove existing toast
  const existing = document.querySelector(".toast");
  if (existing) existing.remove();

  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateY(20px)";
    toast.style.transition = ".3s ease";
    setTimeout(() => toast.remove(), 300);
  }, duration);
}
