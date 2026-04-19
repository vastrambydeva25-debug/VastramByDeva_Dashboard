// ============================================
// ⚡ Data Cache — prevents re-fetching on every page
// ============================================
import { db } from "./firebase-config.js";
import { collection, getDocs } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const CACHE_KEY = "vd_orders_cache";
const CACHE_TIME_KEY = "vd_orders_cache_time";
const CACHE_DURATION = 2 * 60 * 1000; // 2 minutes

let _ordersInMemory = null;

/**
 * Get all orders — uses memory > sessionStorage > Firestore
 * Call forceRefresh() after saving/updating/deleting an order
 */
export async function getOrders() {
  // 1. Memory cache (instant)
  if (_ordersInMemory) return _ordersInMemory;

  // 2. Session storage cache (fast)
  const cached = sessionStorage.getItem(CACHE_KEY);
  const cachedTime = sessionStorage.getItem(CACHE_TIME_KEY);
  if (cached && cachedTime && (Date.now() - Number(cachedTime)) < CACHE_DURATION) {
    _ordersInMemory = JSON.parse(cached);
    return _ordersInMemory;
  }

  // 3. Firestore (network)
  return await forceRefresh();
}

/**
 * Force re-fetch from Firestore and update cache
 */
export async function forceRefresh() {
  const snapshot = await getDocs(collection(db, "orders"));
  const orders = [];
  snapshot.forEach(d => {
    orders.push({ ...d.data(), _id: d.id });
  });
  orders.sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  _ordersInMemory = orders;
  try {
    sessionStorage.setItem(CACHE_KEY, JSON.stringify(orders));
    sessionStorage.setItem(CACHE_TIME_KEY, String(Date.now()));
  } catch (e) {
    // sessionStorage full — that's okay, memory cache still works
  }
  return orders;
}

/**
 * Invalidate cache — call after create/update/delete
 */
export function invalidateCache() {
  _ordersInMemory = null;
  sessionStorage.removeItem(CACHE_KEY);
  sessionStorage.removeItem(CACHE_TIME_KEY);
}
