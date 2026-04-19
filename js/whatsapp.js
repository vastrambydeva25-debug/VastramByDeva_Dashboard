// ============================================
// 📱 WhatsApp Message Templates
// ============================================

/**
 * Open WhatsApp with a pre-filled message
 */
function openWhatsApp(phone, message) {
  // Clean phone number
  let cleaned = phone.replace(/[\s\-\(\)]/g, "");
  if (cleaned.startsWith("0")) cleaned = "91" + cleaned.slice(1);
  if (!cleaned.startsWith("+") && !cleaned.startsWith("91")) cleaned = "91" + cleaned;
  cleaned = cleaned.replace("+", "");

  const url = `https://wa.me/${cleaned}?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

/**
 * Send order confirmation
 */
export function sendOrderConfirmation(order) {
  const message = `🙏 *Vastram By Deva*

Hello *${order.customer}*! 🌸

Your order has been confirmed! ✅

📋 *Order Details:*
🆔 Order ID: ${order.orderId}
👗 Item: ${order.item || "Saree"}
💰 Amount: ₹${order.custAmount}
📅 Date: ${order.date}

We will update you once your order is shipped! 🚚

Thank you for choosing Vastram By Deva! 🙏✨
_Elegant • Traditional • Premium_`;

  openWhatsApp(order.contact, message);
}

/**
 * Send shipping/tracking update
 */
export function sendTrackingMessage(order) {
  const message = `🚚 *Shipping Update — Vastram By Deva*

Hello *${order.customer}*! 🌸

Your order has been shipped! 📦

🆔 Order ID: ${order.orderId}
🚛 Courier: ${order.courier || "N/A"}
📍 Tracking: ${order.tracking || "N/A"}

📦 *Shipping To:*
${order.addressLine1 || ""}
${order.addressLine2 || ""}
${order.city || ""} ${order.pincode || ""}
${order.state || ""}

You can track your order with the tracking number above.

Thank you! 🙏✨
_Vastram By Deva_`;

  openWhatsApp(order.contact, message);
}

/**
 * Send delivery confirmation
 */
export function sendDeliveryMessage(order) {
  const message = `✅ *Delivered — Vastram By Deva*

Hello *${order.customer}*! 🌸

Your order *${order.orderId}* has been delivered! 🎉

We hope you love your purchase! 💖
If you have any questions, feel free to message us.

⭐ We'd love your feedback!

Thank you for choosing Vastram By Deva! 🙏✨`;

  openWhatsApp(order.contact, message);
}

/**
 * Send payment reminder
 */
export function sendPaymentReminder(order) {
  const due = order.custAmount - (order.paidAmount || 0);
  const message = `💰 *Payment Reminder — Vastram By Deva*

Hello *${order.customer}*! 🌸

This is a gentle reminder regarding your order:

🆔 Order ID: ${order.orderId}
💰 Total Amount: ₹${order.custAmount}
⏳ Pending: ₹${due}

Please complete the payment at your earliest convenience.

Thank you! 🙏
_Vastram By Deva_`;

  openWhatsApp(order.contact, message);
}
