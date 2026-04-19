// ============================================
// 📊 Export Utilities — Colorful Excel & PDF
// ============================================

// ---- Helper: Load external script ----
function loadScript(url) {
  return new Promise((resolve, reject) => {
    // Check if already loaded
    const existing = document.querySelector(`script[src="${url}"]`);
    if (existing) { resolve(); return; }
    const s = document.createElement("script");
    s.src = url;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Failed to load: " + url));
    document.head.appendChild(s);
  });
}

// ---- Compute totals helper ----
function computeTotals(orders) {
  const totalRevenue = orders.reduce((s, o) => s + (Number(o.custAmount) || 0), 0);
  const totalVendor = orders.reduce((s, o) => s + (Number(o.vendorAmount) || 0), 0);
  const totalExpense = orders.reduce((s, o) => s + (Number(o.expenseAmount) || 0), 0);
  const totalProfit = orders.reduce((s, o) => s + (Number(o.profit) || 0), 0);
  return { totalRevenue, totalVendor, totalExpense, totalProfit };
}

// ============================================
// 📗 COLORFUL EXCEL EXPORT (ExcelJS)
// ============================================

export async function exportToExcel(orders, filename = "Vastram_Orders") {
  try {
    await loadScript("https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js");
  } catch {
    alert("Could not load Excel library. Check your internet connection.");
    return;
  }

  const ExcelJS = window.ExcelJS;
  const wb = new ExcelJS.Workbook();
  wb.creator = "Vastram By Deva";
  wb.created = new Date();

  const { totalRevenue, totalVendor, totalExpense, totalProfit } = computeTotals(orders);

  // Brand colors
  const MAROON = "FF6E2C1A";
  const GOLD = "FFC8A24D";
  const CREAM = "FFF9F3EC";
  const WHITE = "FFFFFFFF";
  const GREEN_BG = "FFF0FDF4";
  const GREEN_TXT = "FF16A34A";
  const RED_BG = "FFFEF2F2";
  const RED_TXT = "FFDC2626";
  const AMBER_BG = "FFFFFBEB";
  const AMBER_TXT = "FFB45309";
  const BLUE_BG = "FFEFF6FF";
  const BLUE_TXT = "FF1D4ED8";
  const DARK = "FF333333";

  const headerFont = { bold: true, color: { argb: WHITE }, size: 11, name: "Calibri" };
  const headerFill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  const headerBorder = {
    top: { style: "thin", color: { argb: GOLD } },
    bottom: { style: "thin", color: { argb: GOLD } },
    left: { style: "thin", color: { argb: GOLD } },
    right: { style: "thin", color: { argb: GOLD } }
  };
  const thinBorder = {
    top: { style: "thin", color: { argb: "FFE5E5E5" } },
    bottom: { style: "thin", color: { argb: "FFE5E5E5" } },
    left: { style: "thin", color: { argb: "FFE5E5E5" } },
    right: { style: "thin", color: { argb: "FFE5E5E5" } }
  };

  // ========== SHEET 1: ORDERS ==========
  const ws1 = wb.addWorksheet("Orders", {
    properties: { tabColor: { argb: MAROON } },
    views: [{ state: "frozen", ySplit: 2 }]
  });

  // Title row
  ws1.mergeCells("A1:T1");
  const titleCell = ws1.getCell("A1");
  titleCell.value = "VASTRAM BY DEVA — ORDER REPORT";
  titleCell.font = { bold: true, size: 16, color: { argb: WHITE }, name: "Calibri" };
  titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws1.getRow(1).height = 36;

  // Date subtitle
  ws1.mergeCells("A2:T2");
  const dateCell = ws1.getCell("A2");
  dateCell.value = "Generated: " + new Date().toLocaleDateString("en-IN", { weekday: "long", year: "numeric", month: "long", day: "numeric" }) + "  |  Total Orders: " + orders.length;
  dateCell.font = { italic: true, size: 10, color: { argb: GOLD }, name: "Calibri" };
  dateCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  dateCell.alignment = { horizontal: "center", vertical: "middle" };
  ws1.getRow(2).height = 22;

  // Headers (row 3)
  const headers = [
    "Order ID", "Date", "Category", "Item", "Customer", "Contact",
    "City", "State", "Vendor", "Courier", "Tracking",
    "Amount (₹)", "Vendor Amt (₹)", "Expense (₹)", "Profit (₹)", "Profit %",
    "Payment Mode", "Delivery", "Vendor Pay", "Notes"
  ];

  const headerRow = ws1.addRow(headers);
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.border = headerBorder;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  // Column widths
  const colWidths = [30, 13, 11, 24, 20, 15, 13, 13, 18, 13, 20, 15, 15, 13, 13, 11, 14, 13, 13, 24];
  colWidths.forEach((w, i) => { ws1.getColumn(i + 1).width = w; });

  // Data rows
  orders.forEach((o, idx) => {
    const profit = Number(o.profit) || 0;
    const margin = o.custAmount ? Math.round((profit / o.custAmount) * 100) : 0;
    const delivery = o.deliveryStatus || "Pending";
    const vendorPay = o.vendorStatus || "Pending";

    const row = ws1.addRow([
      o.orderId || "-", o.date || "-", o.category || "-", o.item || "-",
      o.customer || "-", o.contact || "-", o.city || "-", o.state || "-",
      o.vendor || "-", o.courier || "-", o.tracking || "-",
      Number(o.custAmount) || 0, Number(o.vendorAmount) || 0,
      Number(o.expenseAmount) || 0, profit, margin + "%",
      o.custMode || "-", delivery, vendorPay, o.notes || ""
    ]);

    // Alternate row colors
    const rowFill = idx % 2 === 0
      ? { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } }
      : { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };

    row.eachCell((cell, colNum) => {
      cell.border = thinBorder;
      cell.font = { size: 10, name: "Calibri", color: { argb: DARK } };
      cell.fill = rowFill;
      cell.alignment = { vertical: "middle" };
    });

    // Amount columns - right aligned, number format
    [12, 13, 14, 15].forEach(col => {
      const cell = row.getCell(col);
      cell.numFmt = "#,##0";
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.font = { size: 10, name: "Calibri", bold: col === 12 || col === 15, color: { argb: DARK } };
    });

    // Profit coloring
    const profitCell = row.getCell(15);
    if (profit > 0) {
      profitCell.font = { size: 10, name: "Calibri", bold: true, color: { argb: GREEN_TXT } };
      profitCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_BG } };
    } else if (profit < 0) {
      profitCell.font = { size: 10, name: "Calibri", bold: true, color: { argb: RED_TXT } };
      profitCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED_BG } };
    }

    // Margin coloring
    const marginCell = row.getCell(16);
    marginCell.alignment = { horizontal: "center", vertical: "middle" };
    if (margin >= 30) {
      marginCell.font = { size: 10, bold: true, color: { argb: GREEN_TXT } };
      marginCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_BG } };
    } else if (margin >= 15) {
      marginCell.font = { size: 10, bold: true, color: { argb: AMBER_TXT } };
      marginCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMBER_BG } };
    } else {
      marginCell.font = { size: 10, bold: true, color: { argb: RED_TXT } };
      marginCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED_BG } };
    }

    // Delivery status coloring
    const delCell = row.getCell(18);
    delCell.alignment = { horizontal: "center", vertical: "middle" };
    if (delivery === "Delivered") {
      delCell.font = { bold: true, size: 10, color: { argb: GREEN_TXT } };
      delCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_BG } };
    } else if (delivery === "Shipped") {
      delCell.font = { bold: true, size: 10, color: { argb: BLUE_TXT } };
      delCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_BG } };
    } else if (delivery === "Returned") {
      delCell.font = { bold: true, size: 10, color: { argb: RED_TXT } };
      delCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED_BG } };
    } else {
      delCell.font = { bold: true, size: 10, color: { argb: AMBER_TXT } };
      delCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMBER_BG } };
    }

    // Vendor payment coloring
    const vpCell = row.getCell(19);
    vpCell.alignment = { horizontal: "center", vertical: "middle" };
    if (vendorPay === "Paid") {
      vpCell.font = { bold: true, size: 10, color: { argb: GREEN_TXT } };
      vpCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_BG } };
    } else {
      vpCell.font = { bold: true, size: 10, color: { argb: AMBER_TXT } };
      vpCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMBER_BG } };
    }
  });

  // ---- TOTAL ROW ----
  ws1.addRow([]); // blank row
  const totalRow = ws1.addRow([
    "TOTAL", orders.length + " orders", "", "", "", "", "", "", "", "", "",
    totalRevenue, totalVendor, totalExpense, totalProfit,
    totalRevenue ? Math.round((totalProfit / totalRevenue) * 100) + "%" : "0%",
    "", "", "", ""
  ]);
  totalRow.height = 28;
  totalRow.eachCell((cell) => {
    cell.font = { bold: true, size: 12, color: { argb: WHITE }, name: "Calibri" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
    cell.border = headerBorder;
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  [12, 13, 14, 15].forEach(col => {
    totalRow.getCell(col).numFmt = "#,##0";
    totalRow.getCell(col).alignment = { horizontal: "right", vertical: "middle" };
  });

  // ========== SHEET 2: SUMMARY DASHBOARD ==========
  const ws2 = wb.addWorksheet("Summary", {
    properties: { tabColor: { argb: GOLD } }
  });
  ws2.getColumn(1).width = 28;
  ws2.getColumn(2).width = 22;
  ws2.getColumn(3).width = 5;
  ws2.getColumn(4).width = 28;
  ws2.getColumn(5).width = 22;

  // Title
  ws2.mergeCells("A1:E1");
  const s2Title = ws2.getCell("A1");
  s2Title.value = "VASTRAM BY DEVA — BUSINESS SUMMARY";
  s2Title.font = { bold: true, size: 16, color: { argb: WHITE } };
  s2Title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  s2Title.alignment = { horizontal: "center", vertical: "middle" };
  ws2.getRow(1).height = 36;

  ws2.mergeCells("A2:E2");
  const s2Date = ws2.getCell("A2");
  s2Date.value = new Date().toLocaleDateString("en-IN", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  s2Date.font = { italic: true, size: 10, color: { argb: GOLD } };
  s2Date.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  s2Date.alignment = { horizontal: "center", vertical: "middle" };
  ws2.getRow(2).height = 22;

  // Helper to add styled summary section
  function addSummarySection(ws, startRow, title, items, accentColor) {
    // Section header
    ws.mergeCells(`A${startRow}:B${startRow}`);
    const hdr = ws.getCell(`A${startRow}`);
    hdr.value = title;
    hdr.font = { bold: true, size: 13, color: { argb: WHITE } };
    hdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentColor } };
    hdr.alignment = { vertical: "middle" };
    ws.getRow(startRow).height = 28;

    items.forEach((item, i) => {
      const r = startRow + 1 + i;
      const labelCell = ws.getCell(`A${r}`);
      const valCell = ws.getCell(`B${r}`);
      labelCell.value = item.label;
      valCell.value = item.value;
      labelCell.font = { size: 11, color: { argb: DARK } };
      valCell.font = { size: 11, bold: true, color: { argb: accentColor } };
      valCell.alignment = { horizontal: "right" };
      const bg = i % 2 === 0 ? WHITE : CREAM;
      labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      labelCell.border = thinBorder;
      valCell.border = thinBorder;
      ws.getRow(r).height = 22;
    });

    return startRow + 1 + items.length + 1; // next section start
  }

  let row = 4;
  row = addSummarySection(ws2, row, "💰 FINANCIAL OVERVIEW", [
    { label: "Total Orders", value: orders.length },
    { label: "Total Revenue", value: "₹" + totalRevenue.toLocaleString("en-IN") },
    { label: "Total Vendor Cost", value: "₹" + totalVendor.toLocaleString("en-IN") },
    { label: "Total Expenses", value: "₹" + totalExpense.toLocaleString("en-IN") },
    { label: "Net Profit", value: "₹" + totalProfit.toLocaleString("en-IN") },
    { label: "Profit Margin", value: totalRevenue ? Math.round((totalProfit / totalRevenue) * 100) + "%" : "0%" }
  ], MAROON);

  // Right side - Delivery status
  const delItems = [
    { label: "Pending", value: orders.filter(o => (o.deliveryStatus || "Pending") === "Pending").length },
    { label: "Shipped", value: orders.filter(o => o.deliveryStatus === "Shipped").length },
    { label: "Delivered", value: orders.filter(o => o.deliveryStatus === "Delivered").length },
    { label: "Returned", value: orders.filter(o => o.deliveryStatus === "Returned").length }
  ];

  row = addSummarySection(ws2, row, "🚚 DELIVERY STATUS", delItems, "FF1D4ED8");

  const vpItems = [
    { label: "Vendor Paid", value: orders.filter(o => o.vendorStatus === "Paid").length },
    { label: "Vendor Pending", value: orders.filter(o => o.vendorStatus !== "Paid").length }
  ];
  row = addSummarySection(ws2, row, "💳 VENDOR PAYMENTS", vpItems, GOLD.replace("FF", "FF"));

  // ========== SHEET 3: CATEGORY BREAKDOWN ==========
  const ws3 = wb.addWorksheet("By Category", {
    properties: { tabColor: { argb: GREEN_TXT } }
  });

  ws3.mergeCells("A1:E1");
  const s3Title = ws3.getCell("A1");
  s3Title.value = "CATEGORY BREAKDOWN";
  s3Title.font = { bold: true, size: 14, color: { argb: WHITE } };
  s3Title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  s3Title.alignment = { horizontal: "center", vertical: "middle" };
  ws3.getRow(1).height = 32;

  const catHeaders = ["Category", "Orders", "Revenue (₹)", "Profit (₹)", "Margin %"];
  const catHdrRow = ws3.addRow(catHeaders);
  catHdrRow.height = 24;
  catHdrRow.eachCell(cell => {
    cell.font = headerFont;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GOLD } };
    cell.font = { bold: true, size: 11, color: { argb: MAROON } };
    cell.border = headerBorder;
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  ws3.getColumn(1).width = 18;
  ws3.getColumn(2).width = 12;
  ws3.getColumn(3).width = 18;
  ws3.getColumn(4).width = 18;
  ws3.getColumn(5).width = 14;

  // Category colors for each row
  const catColors = ["FFC8A24D", "FF6E2C1A", "FF16A34A", "FF1D4ED8", "FFDC2626", "FFB45309"];
  const catMap = {};
  orders.forEach(o => {
    const cat = o.category || "Other";
    if (!catMap[cat]) catMap[cat] = { count: 0, revenue: 0, profit: 0 };
    catMap[cat].count++;
    catMap[cat].revenue += Number(o.custAmount) || 0;
    catMap[cat].profit += Number(o.profit) || 0;
  });

  Object.entries(catMap).forEach(([cat, v], idx) => {
    const margin = v.revenue ? Math.round((v.profit / v.revenue) * 100) : 0;
    const r = ws3.addRow([cat, v.count, v.revenue, v.profit, margin + "%"]);
    r.height = 24;
    const accent = catColors[idx % catColors.length];
    r.eachCell((cell, colNum) => {
      cell.border = thinBorder;
      cell.font = { size: 11, name: "Calibri", color: { argb: DARK } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? WHITE : CREAM } };
      cell.alignment = { vertical: "middle" };
    });
    // Category name bold with accent color
    r.getCell(1).font = { bold: true, size: 11, color: { argb: accent } };
    // Revenue & profit
    r.getCell(3).numFmt = "#,##0";
    r.getCell(3).alignment = { horizontal: "right" };
    r.getCell(4).numFmt = "#,##0";
    r.getCell(4).alignment = { horizontal: "right" };
    r.getCell(4).font = { bold: true, size: 11, color: { argb: v.profit >= 0 ? GREEN_TXT : RED_TXT } };
    r.getCell(5).alignment = { horizontal: "center" };
    r.getCell(5).font = { bold: true, size: 11, color: { argb: margin >= 20 ? GREEN_TXT : AMBER_TXT } };
  });

  // ========== SHEET 4: CUSTOMER BREAKDOWN ==========
  const ws4 = wb.addWorksheet("By Customer", {
    properties: { tabColor: { argb: "FF8B5CF6" } }
  });

  ws4.mergeCells("A1:E1");
  const s4Title = ws4.getCell("A1");
  s4Title.value = "TOP CUSTOMERS BY REVENUE";
  s4Title.font = { bold: true, size: 14, color: { argb: WHITE } };
  s4Title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: MAROON } };
  s4Title.alignment = { horizontal: "center", vertical: "middle" };
  ws4.getRow(1).height = 32;

  const custHeaders = ["Customer", "Orders", "Revenue (₹)", "Profit (₹)", "Avg Order (₹)"];
  const custHdrRow = ws4.addRow(custHeaders);
  custHdrRow.height = 24;
  custHdrRow.eachCell(cell => {
    cell.font = { bold: true, size: 11, color: { argb: MAROON } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GOLD } };
    cell.border = headerBorder;
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  ws4.getColumn(1).width = 24;
  ws4.getColumn(2).width = 12;
  ws4.getColumn(3).width = 18;
  ws4.getColumn(4).width = 18;
  ws4.getColumn(5).width = 16;

  const custMap = {};
  orders.forEach(o => {
    const name = (o.customer || "Unknown").trim();
    if (!custMap[name]) custMap[name] = { count: 0, revenue: 0, profit: 0 };
    custMap[name].count++;
    custMap[name].revenue += Number(o.custAmount) || 0;
    custMap[name].profit += Number(o.profit) || 0;
  });

  Object.entries(custMap)
    .sort((a, b) => b[1].revenue - a[1].revenue)
    .forEach(([name, v], idx) => {
      const r = ws4.addRow([name, v.count, v.revenue, v.profit, Math.round(v.revenue / v.count)]);
      r.height = 22;
      r.eachCell((cell) => {
        cell.border = thinBorder;
        cell.font = { size: 10, color: { argb: DARK } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? WHITE : CREAM } };
        cell.alignment = { vertical: "middle" };
      });
      // Top 3 customers get gold highlight
      if (idx < 3) {
        r.getCell(1).font = { bold: true, size: 11, color: { argb: MAROON } };
        r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF7ED" } };
      }
      r.getCell(3).numFmt = "#,##0";
      r.getCell(3).alignment = { horizontal: "right" };
      r.getCell(3).font = { bold: true, size: 10, color: { argb: DARK } };
      r.getCell(4).numFmt = "#,##0";
      r.getCell(4).alignment = { horizontal: "right" };
      r.getCell(4).font = { bold: true, size: 10, color: { argb: v.profit >= 0 ? GREEN_TXT : RED_TXT } };
      r.getCell(5).numFmt = "#,##0";
      r.getCell(5).alignment = { horizontal: "right" };
    });

  // ========== DOWNLOAD ==========
  const dateStr = new Date().toISOString().split("T")[0];
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${filename}_${dateStr}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}


// ============================================
// 📕 PDF EXPORT (jsPDF + autoTable)
// ============================================

export async function exportToPDF(orders, filename = "Vastram_Orders") {
  try {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
    // Wait a tick for jspdf to register
    await new Promise(r => setTimeout(r, 100));
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js");
    await new Promise(r => setTimeout(r, 100));
  } catch (err) {
    console.error("PDF lib load error:", err);
    alert("Could not load PDF library. Check your internet connection.");
    return;
  }

  const JsPDF = window.jspdf?.jsPDF;
  if (!JsPDF) {
    alert("PDF library did not initialize properly. Please try again.");
    return;
  }

  const pdf = new JsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const pageW = 297;

  // Brand colors
  const MAROON = [110, 44, 26];
  const GOLD = [200, 162, 77];
  const CREAM = [249, 243, 236];
  const WHITE = [255, 255, 255];
  const GREEN = [22, 163, 74];
  const RED = [220, 38, 38];
  const AMBER = [245, 158, 11];
  const BLUE = [59, 130, 246];
  const DARK = [55, 55, 55];

  const { totalRevenue, totalVendor, totalExpense, totalProfit } = computeTotals(orders);

  // =============================================
  // PAGE 1: COVER / SUMMARY
  // =============================================

  // Full maroon header
  pdf.setFillColor(...MAROON);
  pdf.rect(0, 0, pageW, 38, "F");

  // Gold accent line
  pdf.setFillColor(...GOLD);
  pdf.rect(0, 38, pageW, 2, "F");

  // Title
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(24);
  pdf.setTextColor(...WHITE);
  pdf.text("Vastram By Deva", 20, 18);

  pdf.setFontSize(11);
  pdf.setFont("helvetica", "normal");
  pdf.setTextColor(...GOLD);
  pdf.text("Elegant  •  Traditional  •  Premium  •  Royal", 20, 28);

  // Date & count on right
  const dateStr = new Date().toLocaleDateString("en-IN", {
    weekday: "long", year: "numeric", month: "long", day: "numeric"
  });
  pdf.setTextColor(...WHITE);
  pdf.setFontSize(9);
  pdf.text(dateStr, pageW - 20, 18, { align: "right" });
  pdf.setFontSize(10);
  pdf.text(`Total Orders: ${orders.length}`, pageW - 20, 28, { align: "right" });

  // ---- 6 Summary Cards ----
  const pendingDel = orders.filter(o => (o.deliveryStatus || "Pending") === "Pending").length;
  const shipped = orders.filter(o => o.deliveryStatus === "Shipped").length;
  const delivered = orders.filter(o => o.deliveryStatus === "Delivered").length;
  const vendorDue = orders.filter(o => o.vendorStatus !== "Paid").length;

  const cards = [
    { label: "Total Revenue", value: "₹" + totalRevenue.toLocaleString("en-IN"), bg: [255, 251, 235], accent: GOLD },
    { label: "Total Expenses", value: "₹" + (totalVendor + totalExpense).toLocaleString("en-IN"), bg: [254, 242, 242], accent: RED },
    { label: "Net Profit", value: "₹" + totalProfit.toLocaleString("en-IN"), bg: [240, 253, 244], accent: GREEN },
    { label: "Profit Margin", value: totalRevenue ? Math.round((totalProfit / totalRevenue) * 100) + "%" : "0%", bg: [249, 243, 236], accent: MAROON },
    { label: "Pending / Shipped", value: `${pendingDel} / ${shipped}`, bg: [255, 251, 235], accent: AMBER },
    { label: "Delivered", value: String(delivered), bg: [240, 253, 244], accent: GREEN }
  ];

  const cardW = 40;
  const cardH = 28;
  const cardGap = 6;
  const totalCardW = cards.length * cardW + (cards.length - 1) * cardGap;
  let cx = (pageW - totalCardW) / 2;
  const cy = 48;

  cards.forEach(card => {
    // Card background
    pdf.setFillColor(...card.bg);
    pdf.roundedRect(cx, cy, cardW, cardH, 3, 3, "F");

    // Left accent bar
    pdf.setFillColor(...card.accent);
    pdf.rect(cx, cy + 3, 2.5, cardH - 6, "F");

    // Label
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(7);
    pdf.setTextColor(120, 120, 120);
    pdf.text(card.label, cx + cardW / 2, cy + 10, { align: "center" });

    // Value
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(14);
    pdf.setTextColor(...card.accent);
    pdf.text(card.value, cx + cardW / 2, cy + 22, { align: "center" });

    cx += cardW + cardGap;
  });

  // ---- Section Title ----
  pdf.setFillColor(...MAROON);
  pdf.roundedRect(20, 84, 4, 10, 1, 1, "F");
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(13);
  pdf.setTextColor(...MAROON);
  pdf.text("Order Details", 28, 92);

  // ---- ORDERS TABLE ----
  const tableHeaders = [
    "Order ID", "Date", "Customer", "Item", "Amount (₹)",
    "Vendor (₹)", "Profit (₹)", "Margin", "Delivery", "Vendor Pay"
  ];

  const tableData = orders.map(o => {
    const profit = Number(o.profit) || 0;
    const margin = o.custAmount ? Math.round((profit / o.custAmount) * 100) : 0;
    return [
      o.orderId || "-",
      o.date || "-",
      o.customer || "-",
      (o.item || "-").substring(0, 22),
      (Number(o.custAmount) || 0).toLocaleString("en-IN"),
      (Number(o.vendorAmount) || 0).toLocaleString("en-IN"),
      profit.toLocaleString("en-IN"),
      margin + "%",
      o.deliveryStatus || "Pending",
      o.vendorStatus || "Pending"
    ];
  });

  // Total row
  tableData.push([
    "TOTAL", orders.length + " orders", "", "",
    totalRevenue.toLocaleString("en-IN"),
    totalVendor.toLocaleString("en-IN"),
    totalProfit.toLocaleString("en-IN"),
    totalRevenue ? Math.round((totalProfit / totalRevenue) * 100) + "%" : "0%",
    "", ""
  ]);

  const totalRowIdx = tableData.length - 1;

  pdf.autoTable({
    startY: 98,
    head: [tableHeaders],
    body: tableData,
    theme: "grid",
    margin: { left: 14, right: 14 },
    styles: {
      fontSize: 7.5,
      cellPadding: 2.5,
      lineColor: [210, 210, 210],
      lineWidth: 0.15,
      textColor: DARK,
      font: "helvetica"
    },
    headStyles: {
      fillColor: MAROON,
      textColor: WHITE,
      fontStyle: "bold",
      fontSize: 7.5,
      halign: "center"
    },
    alternateRowStyles: {
      fillColor: [252, 250, 247]
    },
    columnStyles: {
      0: { cellWidth: 36, fontStyle: "bold", fontSize: 6.5 },
      1: { cellWidth: 18, halign: "center" },
      2: { cellWidth: 30 },
      3: { cellWidth: 34 },
      4: { cellWidth: 22, halign: "right", fontStyle: "bold" },
      5: { cellWidth: 22, halign: "right" },
      6: { cellWidth: 22, halign: "right", fontStyle: "bold" },
      7: { cellWidth: 16, halign: "center" },
      8: { cellWidth: 22, halign: "center" },
      9: { cellWidth: 22, halign: "center" }
    },
    didParseCell: function (data) {
      if (data.section !== "body") return;

      // Profit column coloring
      if (data.column.index === 6) {
        const raw = String(data.cell.raw || "");
        if (raw.startsWith("-") || raw === "0") {
          data.cell.styles.textColor = RED;
        } else {
          data.cell.styles.textColor = GREEN;
        }
      }

      // Delivery status coloring with background pill
      if (data.column.index === 8) {
        const val = data.cell.raw;
        if (val === "Delivered") {
          data.cell.styles.textColor = GREEN;
          data.cell.styles.fillColor = [240, 253, 244];
          data.cell.styles.fontStyle = "bold";
        } else if (val === "Shipped") {
          data.cell.styles.textColor = BLUE;
          data.cell.styles.fillColor = [239, 246, 255];
          data.cell.styles.fontStyle = "bold";
        } else if (val === "Pending") {
          data.cell.styles.textColor = AMBER;
          data.cell.styles.fillColor = [255, 251, 235];
        } else if (val === "Returned") {
          data.cell.styles.textColor = RED;
          data.cell.styles.fillColor = [254, 242, 242];
          data.cell.styles.fontStyle = "bold";
        }
      }

      // Vendor payment coloring
      if (data.column.index === 9) {
        const val = data.cell.raw;
        if (val === "Paid") {
          data.cell.styles.textColor = GREEN;
          data.cell.styles.fillColor = [240, 253, 244];
          data.cell.styles.fontStyle = "bold";
        } else {
          data.cell.styles.textColor = AMBER;
          data.cell.styles.fillColor = [255, 251, 235];
        }
      }

      // Margin coloring
      if (data.column.index === 7) {
        const pct = parseInt(data.cell.raw) || 0;
        if (pct >= 30) data.cell.styles.textColor = GREEN;
        else if (pct >= 15) data.cell.styles.textColor = AMBER;
        else data.cell.styles.textColor = RED;
        data.cell.styles.fontStyle = "bold";
      }

      // Total row - bold maroon with cream background
      if (data.row.index === totalRowIdx) {
        data.cell.styles.fillColor = CREAM;
        data.cell.styles.fontStyle = "bold";
        data.cell.styles.textColor = MAROON;
        data.cell.styles.fontSize = 8;
      }
    },

    // Page footer
    didDrawPage: function () {
      const pageH = pdf.internal.pageSize.height;
      // Footer line
      pdf.setDrawColor(...GOLD);
      pdf.setLineWidth(0.5);
      pdf.line(20, pageH - 14, pageW - 20, pageH - 14);
      // Footer text
      pdf.setFontSize(7);
      pdf.setTextColor(140, 140, 140);
      pdf.text("Vastram By Deva — Business Report", 20, pageH - 9);
      pdf.text(
        `Page ${pdf.internal.getCurrentPageInfo().pageNumber}`,
        pageW - 20, pageH - 9, { align: "right" }
      );
    }
  });

  // ---- Download ----
  const fileDateStr = new Date().toISOString().split("T")[0];
  pdf.save(`${filename}_${fileDateStr}.pdf`);
}


// ============================================
// 🏷️ Shipping Label
// ============================================
export function printShippingLabel(orderId) {
  window.open(`label.html?id=${encodeURIComponent(orderId)}`, "_blank");
}
