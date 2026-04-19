// ============================================
// 🆔 Order ID Generator (Firestore-based counter)
// ============================================
import { db } from "./firebase-config.js";
import { doc, getDoc, setDoc, runTransaction } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

/**
 * Generates a unique Order ID using Firestore atomic counter.
 * Format: VD_CustomerName_0001_ddmmyy
 */
export async function generateOrderId(customerName, orderDate) {
  const [yyyy, mm, dd] = orderDate.split("-");
  const dateCode = `${dd}${mm}${yyyy.slice(-2)}`;

  // Clean customer name
  let name = customerName.trim().replace(/\s+/g, "").replace(/[^a-zA-Z0-9]/g, "");
  if (!name) name = "Order";

  // Atomic counter in Firestore
  const counterRef = doc(db, "counters", "orderCounter");

  const newSeq = await runTransaction(db, async (transaction) => {
    const counterDoc = await transaction.get(counterRef);
    const current = counterDoc.exists() ? counterDoc.data().value : 0;
    const next = current + 1;
    transaction.set(counterRef, { value: next });
    return next;
  });

  const seq = String(newSeq).padStart(4, "0");
  return `VD_${name}_${seq}_${dateCode}`;
}
