// ============================================
// 🧭 Shared Navigation Component
// ============================================
import { logout } from "./auth-guard.js";

export function renderNav(activePage) {
  const nav = document.createElement("nav");
  nav.className = "app-nav";
  nav.innerHTML = `
    <a href="home.html" class="nav-brand" style="text-decoration:none;">
      <img src="assets/Vastram_logo.png" alt="Vastram" class="nav-logo">
      <div>
        <div class="nav-title">Vastram By Deva</div>
        <div class="nav-subtitle" id="userEmail"></div>
      </div>
    </a>
    <div class="nav-links">
      <a href="home.html" class="${activePage === 'home' ? 'active' : ''}">
        <span class="nav-icon">🏠</span> Home
      </a>
      <a href="dashboard.html" class="${activePage === 'dashboard' ? 'active' : ''}">
        <span class="nav-icon">📊</span> Dashboard
      </a>
      <a href="orders.html" class="${activePage === 'orders' ? 'active' : ''}">
        <span class="nav-icon">➕</span> New Order
      </a>
      <a href="reports.html" class="${activePage === 'reports' ? 'active' : ''}">
        <span class="nav-icon">📋</span> Reports
      </a>
      <a href="customers.html" class="${activePage === 'customers' ? 'active' : ''}">
        <span class="nav-icon">👥</span> Customers
      </a>
      <a href="vendors.html" class="${activePage === 'vendors' ? 'active' : ''}">
        <span class="nav-icon">🏪</span> Vendors
      </a>
    </div>
    <div class="nav-actions">
      <button class="nav-logout" id="logoutBtn">Logout</button>
    </div>
  `;

  document.body.prepend(nav);

  document.getElementById("logoutBtn").addEventListener("click", logout);
}
