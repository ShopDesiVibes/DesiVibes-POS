// ╔══════════════════════════════════════════════════════════════════╗
// ║  🪻 DESI VIBES — Inventory & POS System                        ║
// ║  Google Apps Script · Single File Deploy                       ║
// ║  Paste entire file into Apps Script editor and deploy as:      ║
// ║    Execute as: Me  |  Access: Anyone                           ║
// ╚══════════════════════════════════════════════════════════════════╝

const CFG = {
  SHEET_ID : '1xJzsAL2vA6iQpoyQndavBLrcKsLXDXfubwCPU8GK4qY',
  EMAILS   : ['tupi.desivibes@gmail.com', 'tupi.desai@gmail.com'],
  PREFIX   : 'DV',
  TZ       : Session.getScriptTimeZone()
};

// ═══════════════════════════════════════════════════════════════════
//  SERVE APP
// ═══════════════════════════════════════════════════════════════════
function doGet() {
  return HtmlService.createHtmlOutput(getHTML())
    .setTitle('🪻 Desi Vibes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ═══════════════════════════════════════════════════════════════════
function ss_() { return SpreadsheetApp.openById(CFG.SHEET_ID); }

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#E8651A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function suppliersSheet_() {
  return sheet_('Suppliers', ['Supplier ID','Name','Added Date']);
}
function shipmentsSheet_() {
  return sheet_('Shipments', [
    'Shipment ID','Date','Supplier','Collection','Total Units',
    'Product INR','Shipping INR','Total Paid USD',
    'Product USD','Shipping USD','Effective Rate INR/USD',
    'Payment Method','Notes'
  ]);
}
function itemsSheet_() {
  return sheet_('Items', [
    'Item ID','Shipment ID','Date Received','Category','SKU',
    'Cost USD','Shipping Per Unit','Total Cost',
    'Suggested Retail','Actual Retail','Status','Notes'
  ]);
}
function salesSheet_() {
  return sheet_('Sales', [
    'Sale ID','Sale Date','Item ID','Category',
    'Customer Name','Customer Phone','Sale Price','Cost',
    'Gross Profit','Payment Method','Payment Status',
    'Paid Date','Paid Method','Notes'
  ]);
}

function fmtDate_(d) {
  try { return Utilities.formatDate(new Date(d), CFG.TZ, 'MMM d, yyyy'); } catch(e) { return ''; }
}

// ═══════════════════════════════════════════════════════════════════
//  SUPPLIERS
// ═══════════════════════════════════════════════════════════════════
function getSuppliers() {
  const data = suppliersSheet_().getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({ id: r[0], name: r[1] }));
}

function addSupplier(name) {
  if (!name || !name.trim()) return { ok: false, error: 'Name required' };
  const sh = suppliersSheet_();
  const id = 'SUP-' + String(sh.getLastRow()).padStart(3,'0');
  sh.appendRow([id, name.trim(), new Date().toISOString()]);
  return { ok: true, id, name: name.trim() };
}

// ═══════════════════════════════════════════════════════════════════
//  SHIPMENTS
// ═══════════════════════════════════════════════════════════════════
function logShipment(d) {
  const sh = shipmentsSheet_();
  const id = 'SHP-' + String(sh.getLastRow()).padStart(4,'0');
  const pINR = parseFloat(d.productINR) || 0;
  const sINR = parseFloat(d.shippingINR) || 0;
  const tINR = pINR + sINR;
  const tUSD = parseFloat(d.totalUSD) || 0;
  const rate = (tINR > 0 && tUSD > 0) ? Math.round(tINR / tUSD * 100) / 100 : 0;
  const pUSD = tINR > 0 ? Math.round(tUSD * (pINR / tINR) * 100) / 100 : 0;
  const shUSD = Math.round((tUSD - pUSD) * 100) / 100;

  sh.appendRow([
    id, d.date || new Date().toISOString(), d.supplier, d.collection,
    parseInt(d.totalUnits) || 0,
    pINR, sINR, tUSD, pUSD, shUSD, rate,
    d.paymentMethod, d.notes || ''
  ]);
  return { ok: true, id };
}

function getShipments() {
  const data = shipmentsSheet_().getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({
    id: r[0],
    label: r[0] + ' — ' + r[2] + ' (' + fmtDate_(r[1]) + ')',
    supplier: r[2], collection: r[3],
    totalUnits: r[4], totalUSD: r[7],
    productUSD: r[8], shippingUSD: r[9]
  }));
}

// ═══════════════════════════════════════════════════════════════════
//  ITEMS
// ═══════════════════════════════════════════════════════════════════
function getNextItemId_() {
  const sh = itemsSheet_();
  const num = Math.max(1, sh.getLastRow());
  return CFG.PREFIX + '-' + String(num).padStart(3, '0');
}

function receiveItem(d) {
  const sh = itemsSheet_();
  const itemId = getNextItemId_();
  const cost   = parseFloat(d.costUSD) || 0;
  const ship   = parseFloat(d.shippingPerUnit) || 0;
  const total  = Math.round((cost + ship) * 100) / 100;
  const sugg   = Math.round(total * 2 * 100) / 100;
  const retail = parseFloat(d.retailPrice) || sugg;

  sh.appendRow([
    itemId, d.shipmentId, new Date().toISOString(),
    d.category, d.sku || '', cost, ship, total, sugg, retail,
    'Available', d.notes || ''
  ]);
  return { ok: true, itemId, nextId: getNextItemId_() };
}

function getShipmentTally(shipmentId) {
  const shipData = shipmentsSheet_().getDataRange().getValues();
  const shpRow = shipData.slice(1).find(r => r[0] === shipmentId);
  if (!shpRow) return null;

  const itemData = itemsSheet_().getDataRange().getValues();
  const items = itemData.slice(1).filter(r => r[1] === shipmentId);
  const costTotal = items.reduce((s, r) => s + (parseFloat(r[5]) || 0), 0);
  const shipTotal = items.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const totalUnits = parseInt(shpRow[4]) || 0;
  const perUnit = totalUnits > 0 ? Math.round(shpRow[9] / totalUnits * 100) / 100 : 0;

  return {
    shipmentId, supplier: shpRow[2], collection: shpRow[3],
    totalUnits, totalProductUSD: shpRow[8], totalShippingUSD: shpRow[9],
    perUnitShipping: perUnit,
    itemsTagged: items.length,
    costTotal: Math.round(costTotal * 100) / 100,
    shipTotal: Math.round(shipTotal * 100) / 100,
    nextItemId: getNextItemId_()
  };
}

// ═══════════════════════════════════════════════════════════════════
//  SALES
// ═══════════════════════════════════════════════════════════════════
function getAvailableItems() {
  const data = itemsSheet_().getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(r => r[10] === 'Available').map(r => ({
    itemId: r[0], category: r[3], sku: r[4],
    cost: r[7], retail: r[9]
  }));
}

function logSale(d) {
  const itemSh = itemsSheet_();
  const itemData = itemSh.getDataRange().getValues();
  const iRow = itemData.findIndex(r => r[0] === d.itemId);
  if (iRow < 1) return { ok: false, error: 'Item not found or already sold' };
  if (itemData[iRow][10] !== 'Available') return { ok: false, error: 'Item not available' };

  const cost     = parseFloat(itemData[iRow][7]) || 0;
  const category = itemData[iRow][3];
  itemSh.getRange(iRow + 1, 11).setValue('Sold');

  const saleSh   = salesSheet_();
  const saleId   = 'SAL-' + String(saleSh.getLastRow()).padStart(4, '0');
  const salePrice= parseFloat(d.salePrice) || 0;
  const profit   = Math.round((salePrice - cost) * 100) / 100;
  const pending  = d.paymentMethod === 'Pending';

  saleSh.appendRow([
    saleId, new Date().toISOString(), d.itemId, category,
    d.customerName || '', d.customerPhone || '',
    salePrice, cost, profit,
    d.paymentMethod, pending ? 'Pending' : 'Paid',
    pending ? '' : new Date().toISOString(),
    pending ? '' : d.paymentMethod,
    d.notes || ''
  ]);
  return { ok: true, saleId };
}

// ═══════════════════════════════════════════════════════════════════
//  PENDING PAYMENTS
// ═══════════════════════════════════════════════════════════════════
function getPendingPayments() {
  const data = salesSheet_().getDataRange().getValues();
  if (data.length <= 1) return [];
  const now = Date.now();
  return data.slice(1).filter(r => r[10] === 'Pending').map(r => ({
    saleId: r[0],
    saleDate: fmtDate_(r[1]),
    itemId: r[2], category: r[3],
    customerName: r[4] || 'No name',
    customerPhone: r[5] || '',
    salePrice: parseFloat(r[6]) || 0,
    daysAgo: Math.floor((now - new Date(r[1]).getTime()) / 86400000)
  }));
}

function markAsPaid(saleId, paymentMethod) {
  const sh = salesSheet_();
  const data = sh.getDataRange().getValues();
  const row = data.findIndex(r => r[0] === saleId);
  if (row < 1) return { ok: false, error: 'Sale not found' };
  sh.getRange(row + 1, 11).setValue('Paid');
  sh.getRange(row + 1, 12).setValue(new Date().toISOString());
  sh.getRange(row + 1, 13).setValue(paymentMethod);
  return { ok: true };
}

function sendPendingReminder() {
  const pending = getPendingPayments();
  if (!pending.length) return { ok: true, msg: 'No pending payments to remind about' };
  const total = pending.reduce((s, p) => s + p.salePrice, 0).toFixed(2);
  const list  = pending.map(p =>
    `• ${p.itemId} (${p.category}) — ${p.customerName} — $${p.salePrice.toFixed(2)} — ${p.daysAgo} day(s) ago${p.customerPhone ? ' — ' + p.customerPhone : ''}`
  ).join('\n');
  const subject = `🪻 Desi Vibes: ${pending.length} Pending Payment${pending.length > 1 ? 's' : ''} — $${total} owed`;
  const body = `Desi Vibes Pending Payment Reminder\n${'='.repeat(40)}\n\n${pending.length} unpaid sale(s) totaling $${total}:\n\n${list}\n\n${'─'.repeat(40)}\nLog in to the Desi Vibes app to mark as paid.`;
  CFG.EMAILS.forEach(email => {
    try { MailApp.sendEmail(email, subject, body); } catch (e) { Logger.log('Email error: ' + e); }
  });
  return { ok: true, count: pending.length, total };
}

// ═══════════════════════════════════════════════════════════════════
//  DASHBOARD
// ═══════════════════════════════════════════════════════════════════
function getDashboard() {
  const items = itemsSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const sales = salesSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const now   = Date.now();

  const avail    = items.filter(r => r[10] === 'Available');
  const sold     = items.filter(r => r[10] === 'Sold');
  const pending  = sales.filter(r => r[10] === 'Pending');
  const paidSale = sales.filter(r => r[10] === 'Paid');

  const revenue    = paidSale.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const cogs       = paidSale.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);
  const retailVal  = avail.reduce((s, r) => s + (parseFloat(r[9]) || 0), 0);
  const pendingAmt = pending.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const aging30    = avail.filter(r => (now - new Date(r[2]).getTime()) / 86400000 >= 30).length;

  return {
    available: avail.length,
    sold: sold.length,
    revenue: +revenue.toFixed(2),
    profit: +(revenue - cogs).toFixed(2),
    margin: revenue > 0 ? +((revenue - cogs) / revenue * 100).toFixed(1) : 0,
    retailValue: +retailVal.toFixed(2),
    pendingCount: pending.length,
    pendingAmt: +pendingAmt.toFixed(2),
    aging30,
    totalShipments: shipmentsSheet_().getLastRow() - 1
  };
}

// ═══════════════════════════════════════════════════════════════════
//  INVENTORY AGE
// ═══════════════════════════════════════════════════════════════════
function getInventoryAge() {
  const now  = Date.now();
  const data = itemsSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  return data.filter(r => r[10] === 'Available').map(r => {
    const days  = Math.floor((now - new Date(r[2]).getTime()) / 86400000);
    const retail= parseFloat(r[9]) || 0;
    const cost  = parseFloat(r[7]) || 0;
    let level = 'good', suggest = '✅ Fresh stock — no action needed', mdPrice = retail;
    if (days >= 60) {
      mdPrice = +(retail * 0.75).toFixed(2);
      const m = mdPrice > 0 ? Math.round((mdPrice - cost) / mdPrice * 100) : 0;
      suggest = `🔴 Suggest 25% off → $${mdPrice} (${m}% margin) · Move this item!`;
      level = 'critical';
    } else if (days >= 45) {
      mdPrice = +(retail * 0.85).toFixed(2);
      const m = mdPrice > 0 ? Math.round((mdPrice - cost) / mdPrice * 100) : 0;
      suggest = `⚠️ Suggest 15% off → $${mdPrice} (${m}% margin)`;
      level = 'warn';
    } else if (days >= 30) {
      mdPrice = +(retail * 0.90).toFixed(2);
      const m = mdPrice > 0 ? Math.round((mdPrice - cost) / mdPrice * 100) : 0;
      suggest = `⚠️ Suggest 10% off → $${mdPrice} (${m}% margin)`;
      level = 'warn';
    }
    return { itemId: r[0], shipmentId: r[1], category: r[3], sku: r[4],
             cost, retail, days, suggest, level, mdPrice };
  }).sort((a, b) => b.days - a.days);
}

// ═══════════════════════════════════════════════════════════════════
//  REPORTS
// ═══════════════════════════════════════════════════════════════════
function getReports() {
  const shipData = shipmentsSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const itemData = itemsSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const saleData = salesSheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const now = Date.now();

  const paidSales = saleData.filter(r => r[10] === 'Paid');
  const revenue   = paidSales.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const cogs      = paidSales.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);
  const invested  = shipData.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);
  const avail     = itemData.filter(r => r[10] === 'Available');
  const retailOH  = avail.reduce((s, r) => s + (parseFloat(r[9]) || 0), 0);
  const costOH    = avail.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);

  const fresh   = avail.filter(r => (now - new Date(r[2]).getTime()) / 86400000 < 30).length;
  const aging30 = avail.filter(r => { const d=(now-new Date(r[2]).getTime())/86400000; return d>=30&&d<60; }).length;
  const aging60 = avail.filter(r => (now - new Date(r[2]).getTime()) / 86400000 >= 60).length;

  const byShipment = shipData.map(shp => {
    const shpId  = shp[0];
    const shpItems = itemData.filter(r => r[1] === shpId);
    const soldItems= shpItems.filter(r => r[10] === 'Sold');
    const availIt  = shpItems.filter(r => r[10] === 'Available');
    const shpSales = saleData.filter(r => shpItems.some(i => i[0] === r[2]) && r[10] === 'Paid');
    const shpRev   = shpSales.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
    const shpCOGS  = shpSales.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);
    let oldestDays = 0;
    if (availIt.length) oldestDays = Math.max(...availIt.map(r => Math.floor((now - new Date(r[2]).getTime()) / 86400000)));
    const sellThru = shpItems.length > 0 ? Math.round(soldItems.length / shpItems.length * 100) : 0;
    return {
      id: shpId, date: fmtDate_(shp[1]), supplier: shp[2], collection: shp[3],
      totalUnits: shp[4], totalPaid: shp[7],
      received: shpItems.length, sold: soldItems.length, available: availIt.length,
      revenue: +shpRev.toFixed(2), cogs: +shpCOGS.toFixed(2),
      profit: +(shpRev - shpCOGS).toFixed(2),
      sellThru, oldestDays
    };
  }).reverse();

  return {
    summary: {
      invested: +invested.toFixed(2), revenue: +revenue.toFixed(2),
      profit: +(revenue - cogs).toFixed(2),
      margin: revenue > 0 ? +((revenue - cogs) / revenue * 100).toFixed(1) : 0,
      costOH: +costOH.toFixed(2), retailOH: +retailOH.toFixed(2)
    },
    health: { fresh, aging30, aging60, total: avail.length },
    byShipment
  };
}

// ═══════════════════════════════════════════════════════════════════
//  INITIALIZE SHEETS (run once from Apps Script editor)
// ═══════════════════════════════════════════════════════════════════
function initializeSheets() {
  suppliersSheet_(); shipmentsSheet_(); itemsSheet_(); salesSheet_();
  return 'All sheets initialized ✅';
}

// ═══════════════════════════════════════════════════════════════════
//  HTML — Full Single-Page App (Mobile-First, Fixed for GAS iframe)
// ═══════════════════════════════════════════════════════════════════
function getHTML() {
return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no,viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<title>🪻 Desi Vibes</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --o:#E8651A;--ol:#F5803E;--ob:#FFF4EE;
  --dk:#1C1917;--md:#44403C;--gy:#78716C;
  --br:#E7E5E4;--bg:#FAF9F7;--wh:#FFFFFF;
  --gn:#16A34A;--gnb:#F0FDF4;
  --rd:#DC2626;--rdb:#FEF2F2;
  --bl:#2563EB;
  --nav-h:60px;
  --header-h:56px;
}
*{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{
  font-family:'Inter',system-ui,-apple-system,sans-serif;
  background:var(--bg);color:var(--dk);
  height:100%;width:100%;
  overflow:hidden;
  font-size:14px;
  -webkit-text-size-adjust:100%;
}
#app{
  display:flex;flex-direction:column;
  width:100%;
  height:100%;
  /* JS will set exact height */
}
#app-header{
  background:var(--o);color:#fff;
  padding:10px 16px 8px;
  flex-shrink:0;height:var(--header-h);
  display:flex;align-items:center;justify-content:space-between;
}
#app-header h1{font-size:18px;font-weight:800;letter-spacing:-0.02em;line-height:1.2}
#app-header .sub{font-size:11px;opacity:.8;margin-top:1px}
#pending-badge{background:#fff;color:var(--o);font-size:11px;font-weight:800;
  border-radius:20px;padding:3px 10px;display:none;white-space:nowrap}
#screen-body{
  flex:1;
  overflow-y:auto;
  overflow-x:hidden;
  padding:12px 12px 8px;
  -webkit-overflow-scrolling:touch;
  overscroll-behavior:contain;
}
#bottom-nav{
  display:flex;
  background:var(--wh);
  border-top:2px solid var(--br);
  flex-shrink:0;
  height:var(--nav-h);
  overflow-x:auto;
  overflow-y:hidden;
  scrollbar-width:none;
  -webkit-overflow-scrolling:touch;
  /* safe area padding for iPhone home bar */
  padding-bottom:env(safe-area-inset-bottom,0px);
}
#bottom-nav::-webkit-scrollbar{display:none}
.nav-item{
  flex:0 0 auto;
  min-width:56px;
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  padding:6px 8px 4px;
  font-size:9px;color:var(--gy);font-weight:700;
  gap:3px;cursor:pointer;position:relative;
  white-space:nowrap;letter-spacing:.03em;
  user-select:none;-webkit-user-select:none;
  -webkit-tap-highlight-color:transparent;
}
.nav-item.active{color:var(--o)}
.nav-item.active .ni{transform:scale(1.15)}
.nav-item .ni{font-size:22px;line-height:1;transition:transform .15s}
.nav-dot{position:absolute;top:4px;right:8px;width:8px;height:8px;
  background:var(--rd);border-radius:50%;display:none;
  border:2px solid var(--wh)}
.screen{display:none;flex-direction:column;gap:10px}
.screen.active{display:flex}
.card{background:var(--wh);border:1px solid var(--br);border-radius:14px;padding:14px}
.card-title{font-size:12px;font-weight:700;margin-bottom:10px;
  display:flex;align-items:center;gap:5px;color:var(--md);letter-spacing:.03em;text-transform:uppercase}
.stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.stat-card{background:var(--wh);border:1px solid var(--br);border-radius:12px;padding:12px 14px}
.stat-val{font-size:24px;font-weight:800;color:var(--o);
  font-variant-numeric:tabular-nums;letter-spacing:-0.03em;line-height:1.1}
.stat-val.gn{color:var(--gn)}.stat-val.sm{font-size:18px}
.stat-label{font-size:10px;color:var(--gy);font-weight:500;margin-top:3px}
.fgrp{display:flex;flex-direction:column;gap:5px;margin-bottom:10px}
.fgrp:last-child{margin-bottom:0}
label{font-size:10px;font-weight:700;color:var(--md);letter-spacing:.06em;text-transform:uppercase}
.inp{
  background:var(--bg);border:1.5px solid var(--br);border-radius:10px;
  padding:12px 13px;font-size:16px;color:var(--dk);
  font-family:'Inter',system-ui,sans-serif;
  width:100%;outline:none;transition:border-color .15s;
  -webkit-appearance:none;appearance:none;
  min-height:48px;
}
.inp:focus{border-color:var(--o);background:var(--wh)}
select.inp{
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%2378716C' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 13px center;padding-right:36px;
}
.inp2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.hint{font-size:11px;color:var(--gy);margin-top:2px;line-height:1.4}
.btn{
  border-radius:12px;border:none;
  padding:14px;font-size:15px;font-weight:700;
  font-family:'Inter',system-ui,sans-serif;
  cursor:pointer;display:flex;align-items:center;justify-content:center;
  gap:6px;width:100%;
  transition:opacity .15s, transform .1s;
  min-height:52px;-webkit-appearance:none;
}
.btn:active{opacity:.8;transform:scale(0.98)}
.btn.primary{background:var(--o);color:#fff}
.btn.green{background:var(--gn);color:#fff}
.btn.red{background:var(--rd);color:#fff}
.btn.ghost{background:var(--bg);color:var(--md);border:1.5px solid var(--br)}
.btn.sm{padding:11px;font-size:13px;min-height:44px}
.alert{border-radius:10px;padding:11px 13px;font-size:12px;font-weight:500;
  display:flex;align-items:flex-start;gap:7px;line-height:1.5}
.alert.warn{background:#FFF7ED;color:#9A3412;border:1px solid #FDBA74}
.alert.info{background:#EFF6FF;color:#1E40AF;border:1px solid #BFDBFE}
.alert.ok{background:var(--gnb);color:#14532D;border:1px solid #86EFAC}
.tabs{display:flex;background:var(--bg);border:1px solid var(--br);
  border-radius:10px;padding:3px;gap:2px}
.tab{flex:1;padding:9px 4px;border-radius:8px;border:none;background:transparent;
  font-size:12px;font-weight:700;font-family:'Inter',system-ui,sans-serif;
  color:var(--gy);cursor:pointer;text-align:center;
  transition:all .15s;-webkit-tap-highlight-color:transparent}
.tab.active{background:var(--wh);color:var(--o);box-shadow:0 1px 4px rgba(0,0,0,.12)}
.preview-box{background:var(--ob);border:1px solid #FDBA74;border-radius:10px;padding:12px 14px;font-size:13px}
.prow{display:flex;justify-content:space-between;padding:3px 0}
.prow.ptot{border-top:1px solid #FDBA74;padding-top:9px;margin-top:6px;
  font-weight:800;font-size:16px;color:var(--o)}
.tally{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:12px;padding:12px 14px}
.tally-title{font-size:11px;font-weight:700;color:#1E40AF;margin-bottom:10px}
.trow{display:flex;justify-content:space-between;font-size:12px;margin-bottom:5px}
.tval{font-weight:700;font-variant-numeric:tabular-nums}
.prog{background:var(--br);border-radius:4px;height:6px;overflow:hidden;margin-top:3px;margin-bottom:6px}
.prog-fill{height:100%;border-radius:4px;transition:width .4s ease}
.prog-fill.o{background:var(--o)}.prog-fill.g{background:var(--gn)}.prog-fill.b{background:var(--bl)}
.item-found{background:var(--gnb);border:1.5px solid #86EFAC;border-radius:12px;padding:12px}
.item-found .iname{font-size:16px;font-weight:700}
.item-found .itag{font-size:12px;color:var(--gy);margin-top:3px}
.drow{display:flex;justify-content:space-between;font-size:13px;padding:3px 0}
.drow .dk{color:var(--gy)}.drow .dv{font-weight:700}
.pc{background:#FFF7ED;border:1.5px solid #FED7AA;border-radius:12px;
  padding:13px;display:flex;flex-direction:column;gap:7px}
.pc-id{font-size:14px;font-weight:800;color:var(--o)}
.pc-name{font-size:12px;color:var(--gy)}
.pc-amt{font-size:26px;font-weight:800;font-variant-numeric:tabular-nums;letter-spacing:-0.03em}
.pc-foot{display:flex;align-items:center;justify-content:space-between;gap:8px}
.badge{border-radius:20px;padding:3px 10px;font-size:11px;font-weight:700}
.badge.rd{background:var(--rd);color:#fff}
.badge.or{background:var(--o);color:#fff}
.badge.gn{background:var(--gn);color:#fff}
.ii{background:var(--wh);border:1px solid var(--br);border-radius:12px;
  padding:12px 14px;display:flex;flex-direction:column;gap:5px}
.ii-top{display:flex;justify-content:space-between;align-items:flex-start}
.ii-id{font-size:14px;font-weight:800;color:var(--o)}
.ii-name{font-size:12px;color:var(--gy)}
.ii-prices{display:flex;gap:12px;font-size:12px;font-variant-numeric:tabular-nums}
.ii-prices .cur{color:var(--gy)}.ii-prices .xout{text-decoration:line-through;color:var(--rd)}
.age-badge{border-radius:8px;padding:4px 9px;font-size:11px;font-weight:700;white-space:nowrap}
.age-badge.good{background:var(--gnb);color:var(--gn)}
.age-badge.warn{background:#FFF7ED;color:#9A3412}
.age-badge.critical{background:var(--rdb);color:var(--rd)}
.sug{font-size:12px;padding:6px 10px;border-radius:8px;font-weight:500;line-height:1.4}
.sug.good{background:var(--gnb);color:var(--gn)}
.sug.warn{background:#FFF7ED;color:#C2410C}
.sug.critical{background:var(--rdb);color:var(--rd)}
.sr{background:var(--wh);border:1px solid var(--br);border-radius:12px;
  padding:12px 14px;display:flex;flex-direction:column;gap:7px}
.sr-top{display:flex;justify-content:space-between;align-items:flex-start}
.sr-id{font-size:14px;font-weight:800;color:var(--o)}
.sr-sub{font-size:12px;color:var(--gy)}
.coll-tag{font-size:11px;font-weight:700;background:var(--ob);color:var(--o);
  border-radius:20px;padding:3px 10px;white-space:nowrap}
.sr-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:5px}
.srs{background:var(--bg);border-radius:8px;padding:7px;text-align:center}
.srs .sv{font-size:16px;font-weight:800;font-variant-numeric:tabular-nums}
.srs .sk{font-size:10px;color:var(--gy);margin-top:2px}
.miniprog{background:var(--br);border-radius:3px;height:5px;overflow:hidden}
.minifill{height:100%;border-radius:3px}
.metric-row{display:flex;justify-content:space-between;align-items:center;
  font-size:13px;padding:7px 0;border-bottom:1px solid var(--br)}
.metric-row:last-child{border-bottom:none}
.mk{color:var(--gy)}.mv{font-weight:700;font-variant-numeric:tabular-nums}
.mv.gn{color:var(--gn)}.mv.rd{color:var(--rd)}.mv.or{color:var(--o)}.mv.bl{color:var(--bl)}
/* Modal */
#modal-overlay{
  position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:100;
  display:none;align-items:flex-end;justify-content:center;
}
#modal-overlay.open{display:flex}
#modal-sheet{
  background:var(--wh);border-radius:22px 22px 0 0;
  padding:18px 16px;
  width:100%;max-width:480px;
  display:flex;flex-direction:column;gap:11px;
  padding-bottom:calc(18px + env(safe-area-inset-bottom,0px));
  max-height:85vh;overflow-y:auto;
}
.modal-handle{width:40px;height:4px;background:var(--br);border-radius:2px;margin:0 auto 4px}
.modal-title{font-size:17px;font-weight:800;text-align:center}
.modal-amt{text-align:center;font-size:34px;font-weight:800;color:var(--o);
  font-variant-numeric:tabular-nums;letter-spacing:-0.03em}
.modal-sub{text-align:center;font-size:12px;color:var(--gy);margin-top:-6px}
.pay-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.pay-btn{
  padding:13px 8px;border-radius:12px;border:1.5px solid var(--br);
  background:var(--bg);font-size:13px;font-weight:700;
  font-family:'Inter',system-ui,sans-serif;
  cursor:pointer;display:flex;flex-direction:column;align-items:center;
  gap:4px;transition:all .15s;min-height:66px;
  -webkit-tap-highlight-color:transparent;
}
.pay-btn.sel{border-color:var(--o);background:var(--ob);color:var(--o)}
.pay-btn .pi{font-size:22px}
/* Toast */
#toast{
  position:fixed;top:calc(var(--header-h) + 10px);left:50%;
  transform:translateX(-50%);
  background:var(--dk);color:#fff;
  padding:11px 20px;border-radius:30px;
  font-size:14px;font-weight:700;z-index:200;
  opacity:0;transition:opacity .25s;
  pointer-events:none;white-space:nowrap;max-width:90vw;
  box-shadow:0 4px 20px rgba(0,0,0,.3);
}
#toast.show{opacity:1}
#toast.ok{background:var(--gn)}
#toast.err{background:var(--rd)}
.divider{border:none;border-top:1px solid var(--br);margin:10px 0}
.section-label{font-size:11px;font-weight:700;color:var(--md);
  letter-spacing:.06em;text-transform:uppercase;padding:2px 0}
.loading{text-align:center;padding:30px 0;color:var(--gy);font-size:14px}
</style>
</head>
<body>
<div id="app">
  <div id="app-header">
    <div>
      <h1>🪻 Desi Vibes</h1>
      <div class="sub" id="screen-title">Dashboard</div>
    </div>
    <div id="pending-badge">0 Pending</div>
  </div>
  <div id="screen-body">
    <!-- DASHBOARD -->
    <div class="screen active" id="scr-dashboard">
      <div class="stat-grid">
        <div class="stat-card"><div class="stat-val" id="d-avail">—</div><div class="stat-label">Items Available</div></div>
        <div class="stat-card"><div class="stat-val" id="d-sold">—</div><div class="stat-label">Items Sold</div></div>
        <div class="stat-card"><div class="stat-val gn sm" id="d-rev">—</div><div class="stat-label">Revenue</div></div>
        <div class="stat-card"><div class="stat-val gn sm" id="d-profit">—</div><div class="stat-label">Gross Profit</div></div>
        <div class="stat-card" style="grid-column:1/-1">
          <div class="stat-val sm" id="d-retail">—</div>
          <div class="stat-label">Inventory on Hand (@ Retail)</div>
        </div>
      </div>
      <div class="alert warn" id="d-pending-alert" style="display:none"></div>
      <div class="card">
        <div class="card-title">📈 Quick Stats</div>
        <div class="metric-row"><span class="mk">Profit Margin</span><span class="mv gn" id="d-margin">—</span></div>
        <div class="metric-row"><span class="mk">Total Shipments</span><span class="mv" id="d-shipments">—</span></div>
        <div class="metric-row"><span class="mk">Total Items Received</span><span class="mv" id="d-total">—</span></div>
        <div class="metric-row"><span class="mk">Items Aging 30+ days</span><span class="mv" id="d-aging">—</span></div>
      </div>
      <button class="btn red sm" id="btn-view-pending" style="display:none" onclick="showScreen('pending')">⏳ View Pending Payments →</button>
      <button class="btn ghost sm" onclick="showScreen('inventory')">📋 Check Inventory Age →</button>
    </div>
    <!-- SHIPMENT -->
    <div class="screen" id="scr-shipment">
      <div class="card">
        <div class="fgrp">
          <label>Supplier</label>
          <select class="inp" id="shp-supplier" onchange="onSupplierChange()">
            <option value="">Loading suppliers...</option>
          </select>
        </div>
        <div class="fgrp" id="shp-new-supplier-grp" style="display:none">
          <label>New Supplier Name</label>
          <input class="inp" id="shp-new-supplier" placeholder="Type supplier name here..." type="text">
          <div class="hint">Will be saved to supplier list automatically</div>
        </div>
        <hr class="divider">
        <div class="fgrp">
          <label>Collection / Category</label>
          <select class="inp" id="shp-collection" onchange="onCollectionChange('shp-collection','shp-custom-collection')">
            <option value="Navratri">🪷 Navratri</option>
            <option value="Diwali">🪔 Diwali</option>
            <option value="New Year">🎉 New Year</option>
            <option value="Everyday Wear">👗 Everyday Wear</option>
            <option value="__other__">✏️ Other (type below)</option>
          </select>
        </div>
        <div class="fgrp" id="shp-custom-coll-grp" style="display:none">
          <label>Custom Collection Name</label>
          <input class="inp" id="shp-custom-collection" placeholder="e.g. Summer Collection..." type="text">
        </div>
        <div class="inp2">
          <div class="fgrp"><label>Date</label><input class="inp" id="shp-date" type="date"></div>
          <div class="fgrp"><label>Total Units</label><input class="inp" id="shp-units" type="number" min="1" placeholder="20" oninput="calcShipmentUSD()"></div>
        </div>
        <div class="hint" style="margin-top:-6px;margin-bottom:10px">← Units auto-calc per-item shipping on Receive screen</div>
        <hr class="divider">
        <div class="inp2">
          <div class="fgrp"><label>Product Cost (INR)</label><input class="inp" id="shp-pinr" type="number" placeholder="₹ 42000" oninput="calcShipmentUSD()"></div>
          <div class="fgrp"><label>Shipping Cost (INR)</label><input class="inp" id="shp-sinr" type="number" placeholder="₹ 12600" oninput="calcShipmentUSD()"></div>
        </div>
        <div class="fgrp">
          <label>Total I Actually Paid (USD)</label>
          <input class="inp" id="shp-usd" type="number" step="0.01" placeholder="$ 650.00" oninput="calcShipmentUSD()">
          <div class="hint">All-in: include Xoom / bank transfer fees</div>
        </div>
        <div class="preview-box" id="shp-preview" style="display:none">
          <div class="prow"><span style="color:var(--md)">Effective Rate</span><span id="pr-rate" style="font-weight:700"></span></div>
          <div class="prow"><span style="color:var(--md)">Product (USD)</span><span id="pr-product"></span></div>
          <div class="prow"><span style="color:var(--md)">Shipping (USD)</span><span id="pr-shipping"></span></div>
          <div class="prow ptot"><span>Total Paid</span><span id="pr-total"></span></div>
        </div>
        <div class="fgrp" style="margin-top:10px">
          <label>Payment Method</label>
          <select class="inp" id="shp-payment">
            <option>Xoom</option><option>Indian Bank Account</option>
            <option>Cash</option><option>Credit Card</option><option>Pending</option>
          </select>
        </div>
        <div class="fgrp"><label>Notes (Optional)</label><input class="inp" id="shp-notes" type="text" placeholder="Any notes about this shipment..."></div>
        <button class="btn primary" style="margin-top:6px" onclick="submitShipment()">📤 Log Shipment</button>
      </div>
    </div>
    <!-- RECEIVE -->
    <div class="screen" id="scr-receive">
      <div class="fgrp">
        <label>Select Shipment</label>
        <select class="inp" id="rcv-shipment" onchange="onShipmentSelect()">
          <option value="">— Choose a shipment —</option>
        </select>
      </div>
      <div id="rcv-tally" style="display:none">
        <div class="tally" id="rcv-tally-box"></div>
        <div class="card">
          <div class="card-title" id="rcv-item-title">➕ Add Item</div>
          <div class="fgrp">
            <label>Category</label>
            <select class="inp" id="rcv-category" onchange="onCollectionChange('rcv-category','rcv-custom-category')">
              <option value="Navratri">🪷 Navratri</option>
              <option value="Diwali">🪔 Diwali</option>
              <option value="New Year">🎉 New Year</option>
              <option value="Everyday Wear">👗 Everyday Wear</option>
              <option value="__other__">✏️ Other (type below)</option>
            </select>
          </div>
          <div class="fgrp" id="rcv-custom-cat-grp" style="display:none">
            <label>Custom Category</label>
            <input class="inp" id="rcv-custom-category" placeholder="e.g. Chaniya Choli..." type="text">
          </div>
          <div class="fgrp"><label>SKU Code (Optional)</label><input class="inp" id="rcv-sku" type="text" placeholder="e.g. ANK-RED-L"></div>
          <div class="inp2">
            <div class="fgrp"><label>Item Cost (USD)</label><input class="inp" id="rcv-cost" type="number" step="0.01" placeholder="$24.85" oninput="calcRetail()"></div>
            <div class="fgrp"><label>Shipping / Unit</label><input class="inp" id="rcv-ship" type="number" step="0.01" placeholder="$7.47" oninput="calcRetail()"><div class="hint" id="rcv-ship-hint">Auto from shipment</div></div>
          </div>
          <div class="preview-box" id="rcv-cost-preview" style="display:none">
            <div class="prow"><span style="color:var(--md)">Total Cost</span><span id="rcp-cost" style="font-weight:700"></span></div>
            <div class="prow ptot"><span style="color:var(--gn)">Suggested Retail (50% margin)</span><span id="rcp-retail" style="color:var(--gn)"></span></div>
          </div>
          <div class="fgrp"><label>Actual Retail Price</label><input class="inp" id="rcv-retail" type="number" step="0.01" placeholder="$64.64"><div class="hint">Edit if you want to price differently</div></div>
          <button class="btn primary" id="rcv-submit-btn" onclick="submitReceiveItem()">➕ Add Item</button>
        </div>
      </div>
      <div id="rcv-empty" class="alert info">Select a shipment above to start receiving items.</div>
    </div>
    <!-- SALE -->
    <div class="screen" id="scr-sale">
      <div class="card">
        <div class="fgrp">
          <label>Search by Item ID</label>
          <div style="display:flex;gap:8px">
            <input class="inp" id="sale-search" type="text" placeholder="DV-021"
              style="text-transform:uppercase;flex:1" onkeyup="if(event.key==='Enter')searchItem()">
            <button class="btn ghost sm" style="width:auto;padding:0 16px;min-height:48px" onclick="searchItem()">Search</button>
          </div>
        </div>
        <div class="fgrp"><label>Or Pick from Available Items</label>
          <select class="inp" id="sale-picker" onchange="onItemPick()">
            <option value="">— Choose an item —</option>
          </select>
        </div>
        <div id="sale-item-found" style="display:none;margin-bottom:8px">
          <div class="item-found">
            <div class="iname" id="si-name"></div>
            <div class="itag" id="si-tag"></div>
            <div style="margin-top:8px;display:flex;flex-direction:column;gap:4px">
              <div class="drow"><span class="dk">Your Cost:</span><span class="dv" id="si-cost"></span></div>
              <div class="drow"><span class="dk">Listed Retail:</span><span class="dv" id="si-retail"></span></div>
            </div>
          </div>
        </div>
        <hr class="divider">
        <div class="fgrp"><label>Sale Price ($)</label>
          <input class="inp" id="sale-price" type="number" step="0.01" placeholder="Enter sale price" oninput="calcSaleMargin()">
          <div class="hint" id="sale-margin-hint"></div>
        </div>
        <div class="inp2">
          <div class="fgrp"><label>Customer Name</label><input class="inp" id="sale-name" type="text" placeholder="Priya Shah"></div>
          <div class="fgrp"><label>Phone (Optional)</label><input class="inp" id="sale-phone" type="tel" placeholder="864-555-0192"></div>
        </div>
        <div class="fgrp">
          <label>Payment Method</label>
          <select class="inp" id="sale-payment">
            <option>Cash</option><option>Credit Card</option><option>Zelle</option>
            <option>Venmo</option><option>PayPal</option><option>CashApp</option>
            <option>Pending</option><option value="__other__">Other...</option>
          </select>
        </div>
        <div class="fgrp" id="sale-other-grp" style="display:none">
          <label>Specify Payment Method</label>
          <input class="inp" id="sale-other" type="text" placeholder="e.g. Bank Transfer...">
        </div>
        <button class="btn green" onclick="submitSale()">✅ Complete Sale</button>
      </div>
    </div>
    <!-- PENDING -->
    <div class="screen" id="scr-pending">
      <div id="pend-summary" class="alert warn" style="display:none"></div>
      <div id="pend-list"></div>
      <button class="btn ghost sm" onclick="sendReminder()">📧 Email Reminder to Myself + Tupi</button>
      <div id="pend-empty" class="alert ok" style="display:none">✅ No pending payments — all clear!</div>
    </div>
    <!-- INVENTORY -->
    <div class="screen" id="scr-inventory">
      <div id="inv-filter" class="tabs">
        <div class="tab active" onclick="filterInv('all',this)">All Items</div>
        <div class="tab" onclick="filterInv('30',this)">⚠️ 30+ days</div>
        <div class="tab" onclick="filterInv('60',this)">🔴 60+ days</div>
      </div>
      <div id="inv-list"></div>
      <div id="inv-empty" class="loading" style="display:none">No items in this filter.</div>
    </div>
    <!-- REPORTS -->
    <div class="screen" id="scr-reports">
      <div class="card">
        <div class="card-title">💼 Overall Summary</div>
        <div class="metric-row"><span class="mk">Total Invested (USD)</span><span class="mv or" id="rp-invested">—</span></div>
        <div class="metric-row"><span class="mk">Total Revenue</span><span class="mv gn" id="rp-rev">—</span></div>
        <div class="metric-row"><span class="mk">Gross Profit</span><span class="mv gn" id="rp-profit">—</span></div>
        <div class="metric-row"><span class="mk">Avg Profit Margin</span><span class="mv gn" id="rp-margin">—</span></div>
        <div class="metric-row"><span class="mk">Remaining (cost)</span><span class="mv" id="rp-costoh">—</span></div>
        <div class="metric-row"><span class="mk">Remaining (@ retail)</span><span class="mv or" id="rp-retailoh">—</span></div>
      </div>
      <div class="card"><div class="card-title">📦 Inventory Health</div><div id="rp-health"></div></div>
      <div class="section-label">📋 By Shipment</div>
      <div id="rp-shipments"><div class="loading">Loading...</div></div>
    </div>
  </div><!-- end screen-body -->

  <div id="bottom-nav">
    <div class="nav-item active" onclick="showScreen('dashboard')"><span class="ni">📊</span>Dashboard</div>
    <div class="nav-item" onclick="showScreen('shipment')"><span class="ni">📤</span>Shipment</div>
    <div class="nav-item" onclick="showScreen('receive')"><span class="ni">📦</span>Receive</div>
    <div class="nav-item" onclick="showScreen('sale')"><span class="ni">💳</span>Sale</div>
    <div class="nav-item" onclick="showScreen('pending')"><span class="ni">⏳</span>Pending<div class="nav-dot" id="nav-dot-pending"></div></div>
    <div class="nav-item" onclick="showScreen('inventory')"><span class="ni">📋</span>Inventory</div>
    <div class="nav-item" onclick="showScreen('reports')"><span class="ni">📈</span>Reports</div>
  </div>
</div><!-- end app -->

<!-- Mark-as-Paid Modal -->
<div id="modal-overlay" onclick="closeModal(event)">
  <div id="modal-sheet">
    <div class="modal-handle"></div>
    <div class="modal-title" id="modal-title">How did they pay?</div>
    <div class="modal-amt" id="modal-amt"></div>
    <div class="modal-sub" id="modal-sub"></div>
    <div class="pay-grid">
      <button class="pay-btn" onclick="selectPayMethod(this,'Cash')"><span class="pi">💵</span>Cash</button>
      <button class="pay-btn" onclick="selectPayMethod(this,'Credit Card')"><span class="pi">💳</span>Credit Card</button>
      <button class="pay-btn" onclick="selectPayMethod(this,'Zelle')"><span class="pi">🟢</span>Zelle</button>
      <button class="pay-btn" onclick="selectPayMethod(this,'Venmo')"><span class="pi">🔵</span>Venmo</button>
      <button class="pay-btn" onclick="selectPayMethod(this,'PayPal')"><span class="pi">🅿️</span>PayPal</button>
      <button class="pay-btn" onclick="selectPayMethod(this,'CashApp')"><span class="pi">💚</span>CashApp</button>
    </div>
    <div class="fgrp"><label>Or Type Method</label>
      <input class="inp" id="modal-other" type="text" placeholder="e.g. Check, Bank Wire...">
    </div>
    <button class="btn green" id="modal-confirm-btn" onclick="confirmMarkAsPaid()">✅ Confirm Payment</button>
    <button class="btn ghost sm" onclick="closeModalDirect()">Cancel</button>
  </div>
</div>
<div id="toast"></div>

<script>
// ─── State ───────────────────────────────────────────────────────
let currentScreen='dashboard',invData=[],modalSaleId='',selectedPayMethod='',currentItemId='',currentItemCost=0;
const SCREEN_TITLES={dashboard:'Dashboard',shipment:'Log Shipment',receive:'Receive Items',sale:'Log Sale',pending:'Pending Payments',inventory:'Inventory Age',reports:'Reports'};

// ─── Fix height for GAS iframe / Safari ──────────────────────────
function fixHeight(){
  const h=window.innerHeight||document.documentElement.clientHeight||screen.height;
  document.getElementById('app').style.height=h+'px';
}
window.addEventListener('resize',fixHeight);
fixHeight();

// ─── Navigation ──────────────────────────────────────────────────
function showScreen(name){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('scr-'+name).classList.add('active');
  const idx=['dashboard','shipment','receive','sale','pending','inventory','reports'].indexOf(name);
  if(idx>=0)document.querySelectorAll('.nav-item')[idx].classList.add('active');
  document.getElementById('screen-title').textContent=SCREEN_TITLES[name]||name;
  document.getElementById('screen-body').scrollTop=0;
  currentScreen=name;
  const loaders={dashboard:loadDashboard,shipment:loadShipmentForm,receive:loadReceiveForm,sale:loadSaleForm,pending:loadPending,inventory:loadInventory,reports:loadReports};
  if(loaders[name])loaders[name]();
}

// ─── Helpers ─────────────────────────────────────────────────────
function toast(msg,type=''){const t=document.getElementById('toast');t.textContent=msg;t.className='show '+(type||'');setTimeout(()=>t.className='',2800);}
const $=id=>document.getElementById(id);
function fmt(n){return'$'+(parseFloat(n)||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});}

// ─── Dashboard ───────────────────────────────────────────────────
function loadDashboard(){
  google.script.run
    .withSuccessHandler(d=>{
      $('d-avail').textContent=d.available;
      $('d-sold').textContent=d.sold;
      $('d-rev').textContent=fmt(d.revenue);
      $('d-profit').textContent=fmt(d.profit);
      $('d-retail').textContent=fmt(d.retailValue);
      $('d-margin').textContent=d.margin+'%';
      $('d-shipments').textContent=d.totalShipments;
      $('d-total').textContent=d.available+d.sold;
      const ag=$('d-aging');ag.textContent=d.aging30+' items';ag.className='mv '+(d.aging30>0?'rd':'gn');
      const pa=$('d-pending-alert'),pb=$('btn-view-pending'),badge=$('pending-badge'),dot=$('nav-dot-pending');
      if(d.pendingCount>0){
        pa.style.display='flex';
        pa.textContent='⚠️ '+d.pendingCount+' sale'+(d.pendingCount>1?'s have':' has')+' unpaid balances · '+fmt(d.pendingAmt);
        pb.style.display='flex';badge.style.display='block';
        badge.textContent=d.pendingCount+' Pending';dot.style.display='block';
      } else {
        pa.style.display='none';pb.style.display='none';
        badge.style.display='none';dot.style.display='none';
      }
    })
    .withFailureHandler(()=>toast('Error loading dashboard','err'))
    .getDashboard();
}

// ─── Shipment Form ───────────────────────────────────────────────
function loadShipmentForm(){
  const today=new Date().toISOString().split('T')[0];$('shp-date').value=today;
  google.script.run
    .withSuccessHandler(suppliers=>{
      const sel=$('shp-supplier');
      sel.innerHTML='<option value="">— Select Supplier —</option>';
      suppliers.forEach(s=>{const opt=document.createElement('option');opt.value=s.name;opt.textContent=s.name;sel.appendChild(opt);});
      const newOpt=document.createElement('option');newOpt.value='__new__';newOpt.textContent='➕ Add New Supplier...';sel.appendChild(newOpt);
    })
    .withFailureHandler(()=>{})
    .getSuppliers();
}
function onSupplierChange(){$('shp-new-supplier-grp').style.display=$('shp-supplier').value==='__new__'?'block':'none';}
function onCollectionChange(selId,inputId){
  const sel=$(selId);
  const grpId=selId==='shp-collection'?'shp-custom-coll-grp':'rcv-custom-cat-grp';
  $(grpId).style.display=sel.value==='__other__'?'block':'none';
}
function calcShipmentUSD(){
  const pINR=parseFloat($('shp-pinr').value)||0,sINR=parseFloat($('shp-sinr').value)||0;
  const tINR=pINR+sINR,tUSD=parseFloat($('shp-usd').value)||0,box=$('shp-preview');
  if(tUSD>0&&tINR>0){
    const rate=tINR/tUSD,pUSD=tUSD*(pINR/tINR),shUSD=tUSD-pUSD;
    $('pr-rate').textContent='1 USD = '+rate.toFixed(1)+' INR';
    $('pr-product').textContent=fmt(pUSD);$('pr-shipping').textContent=fmt(shUSD);
    $('pr-total').textContent=fmt(tUSD);box.style.display='block';
  } else if(tUSD>0){
    $('pr-rate').textContent='Enter INR to see rate';$('pr-product').textContent='—';
    $('pr-shipping').textContent='—';$('pr-total').textContent=fmt(tUSD);box.style.display='block';
  } else { box.style.display='none'; }
}
function submitShipment(){
  let supplier=$('shp-supplier').value;
  if(supplier==='__new__')supplier=$('shp-new-supplier').value.trim();
  if(!supplier)return toast('Please select or enter a supplier name','err');
  const coll=$('shp-collection').value;
  const collection=coll==='__other__'?$('shp-custom-collection').value.trim():coll;
  if(!collection)return toast('Please enter a collection name','err');
  if(!$('shp-usd').value)return toast('Please enter total USD paid','err');
  const d={supplier,collection,date:$('shp-date').value,totalUnits:$('shp-units').value,
    productINR:$('shp-pinr').value,shippingINR:$('shp-sinr').value,
    totalUSD:$('shp-usd').value,paymentMethod:$('shp-payment').value,notes:$('shp-notes').value};
  toast('Saving shipment...');
  google.script.run
    .withSuccessHandler(r=>{
      if(r.ok){toast('✅ Shipment '+r.id+' logged!','ok');
        $('shp-pinr').value='';$('shp-sinr').value='';$('shp-usd').value='';
        $('shp-units').value='';$('shp-notes').value='';$('shp-preview').style.display='none';
      } else toast('Error: '+r.error,'err');
    })
    .withFailureHandler(e=>toast('Error: '+e.message,'err'))
    .logShipment(d);
}

// ─── Receive Items ────────────────────────────────────────────────
function loadReceiveForm(){
  google.script.run
    .withSuccessHandler(shipments=>{
      const sel=$('rcv-shipment');
      sel.innerHTML='<option value="">— Choose a shipment —</option>';
      shipments.forEach(s=>{const opt=document.createElement('option');opt.value=s.id;opt.textContent=s.label;sel.appendChild(opt);});
    })
    .withFailureHandler(()=>{})
    .getShipments();
}
function onShipmentSelect(){
  const id=$('rcv-shipment').value;
  if(!id){$('rcv-tally').style.display='none';$('rcv-empty').style.display='block';return;}
  $('rcv-empty').style.display='none';$('rcv-tally').style.display='block';
  google.script.run
    .withSuccessHandler(t=>{
      if(!t)return;
      const catSel=$('rcv-category'),cats=['Navratri','Diwali','New Year','Everyday Wear'];
      const matchIdx=cats.indexOf(t.collection);catSel.selectedIndex=matchIdx>=0?matchIdx:4;
      $('rcv-ship').value=t.perUnitShipping;
      $('rcv-ship-hint').textContent='Auto: '+fmt(t.totalShippingUSD)+'÷'+t.totalUnits;
      $('rcv-item-title').textContent='➕ Add Item '+t.nextItemId;
      $('rcv-submit-btn').textContent='➕ Add '+t.nextItemId+' · '+(t.totalUnits-t.itemsTagged)+' left';
      renderTally(t);
    })
    .withFailureHandler(()=>toast('Could not load shipment tally','err'))
    .getShipmentTally(id);
}
function renderTally(t){
  const costPct=t.totalProductUSD>0?Math.min(100,Math.round(t.costTotal/t.totalProductUSD*100)):0;
  const shipPct=t.totalShippingUSD>0?Math.min(100,Math.round(t.shipTotal/t.totalShippingUSD*100)):0;
  const itemPct=t.totalUnits>0?Math.min(100,Math.round(t.itemsTagged/t.totalUnits*100)):0;
  $('rcv-tally-box').innerHTML=
    '<div class="tally-title">📊 Tally — '+t.shipmentId+' · '+t.collection+' · '+t.totalUnits+' units</div>'+
    '<div class="trow"><span style="color:var(--md)">Items Tagged</span><span class="tval" style="color:var(--bl)">'+t.itemsTagged+' of '+t.totalUnits+'</span></div>'+
    '<div class="prog"><div class="prog-fill b" style="width:'+itemPct+'%"></div></div>'+
    '<div class="trow"><span style="color:var(--md)">Cost Total</span><span class="tval">'+fmt(t.costTotal)+' of '+fmt(t.totalProductUSD)+'</span></div>'+
    '<div class="prog"><div class="prog-fill o" style="width:'+costPct+'%"></div></div>'+
    '<div class="trow"><span style="color:var(--md)">Shipping Total</span><span class="tval" style="color:var(--gn)">'+fmt(t.shipTotal)+' of '+fmt(t.totalShippingUSD)+'</span></div>'+
    '<div class="prog"><div class="prog-fill g" style="width:'+shipPct+'%"></div></div>'+
    '<div style="margin-top:8px;font-size:12px;color:var(--md)">Next ID: <strong style="color:var(--o)">'+t.nextItemId+'</strong></div>';
}
function calcRetail(){
  const cost=parseFloat($('rcv-cost').value)||0,ship=parseFloat($('rcv-ship').value)||0;
  const total=cost+ship,sugg=total*2,prev=$('rcv-cost-preview');
  if(cost>0){
    $('rcp-cost').textContent=fmt(total);$('rcp-retail').textContent=fmt(sugg);
    prev.style.display='block';if(!$('rcv-retail').value)$('rcv-retail').value=sugg.toFixed(2);
  } else {prev.style.display='none';}
}
function submitReceiveItem(){
  if(!$('rcv-shipment').value)return toast('Select a shipment first','err');
  if(!$('rcv-cost').value)return toast('Enter item cost','err');
  if(!$('rcv-retail').value)return toast('Enter retail price','err');
  const coll=$('rcv-category').value;
  const category=coll==='__other__'?$('rcv-custom-category').value.trim():coll;
  if(!category)return toast('Enter a category name','err');
  const d={shipmentId:$('rcv-shipment').value,category,sku:$('rcv-sku').value,
    costUSD:$('rcv-cost').value,shippingPerUnit:$('rcv-ship').value,retailPrice:$('rcv-retail').value};
  toast('Adding item...');
  google.script.run
    .withSuccessHandler(r=>{
      if(r.ok){
        toast('✅ '+r.itemId+' added!','ok');
        $('rcv-cost').value='';$('rcv-retail').value='';$('rcv-sku').value='';
        $('rcv-cost-preview').style.display='none';
        onShipmentSelect();
      } else toast('Error: '+r.error,'err');
    })
    .withFailureHandler(e=>toast('Error: '+e.message,'err'))
    .receiveItem(d);
}

// ─── Sale ─────────────────────────────────────────────────────────
function loadSaleForm(){
  currentItemId='';currentItemCost=0;$('sale-item-found').style.display='none';
  $('sale-payment').onchange=()=>{$('sale-other-grp').style.display=$('sale-payment').value==='__other__'?'block':'none';};
  google.script.run
    .withSuccessHandler(items=>{
      const sel=$('sale-picker');sel.innerHTML='<option value="">— Choose an item —</option>';
      items.forEach(i=>{
        const opt=document.createElement('option');opt.value=i.itemId;
        opt.textContent=i.itemId+' — '+i.category+' ('+fmt(i.retail)+')';
        opt.dataset.cost=i.cost;opt.dataset.retail=i.retail;opt.dataset.category=i.category;
        sel.appendChild(opt);
      });
    })
    .withFailureHandler(()=>{})
    .getAvailableItems();
}
function searchItem(){
  const id=$('sale-search').value.trim().toUpperCase();if(!id)return;
  const opts=$('sale-picker').options;
  for(let o of opts){if(o.value===id){$('sale-picker').value=id;onItemPick();return;}}
  toast('Item '+id+' not found or not available','err');
}
function onItemPick(){
  const sel=$('sale-picker'),opt=sel.options[sel.selectedIndex];
  if(!sel.value){$('sale-item-found').style.display='none';return;}
  currentItemId=sel.value;currentItemCost=parseFloat(opt.dataset.cost)||0;
  $('si-name').textContent=opt.dataset.category||sel.value;
  $('si-tag').textContent=sel.value+' · ✅ Available';
  $('si-cost').textContent=fmt(opt.dataset.cost);
  $('si-retail').textContent=fmt(opt.dataset.retail);
  $('sale-price').value=parseFloat(opt.dataset.retail).toFixed(2);
  calcSaleMargin();$('sale-item-found').style.display='block';$('sale-search').value=sel.value;
}
function calcSaleMargin(){
  const price=parseFloat($('sale-price').value)||0;
  if(price>0&&currentItemCost>0){
    const margin=((price-currentItemCost)/price*100).toFixed(1);
    const profit=(price-currentItemCost).toFixed(2);
    $('sale-margin-hint').textContent='Profit: '+fmt(profit)+' · Margin: '+margin+'%';
  } else {$('sale-margin-hint').textContent='';}
}
function submitSale(){
  if(!currentItemId)return toast('Select an item to sell','err');
  if(!$('sale-price').value)return toast('Enter a sale price','err');
  let payment=$('sale-payment').value;
  if(payment==='__other__')payment=$('sale-other').value.trim()||'Other';
  const d={itemId:currentItemId,salePrice:$('sale-price').value,
    customerName:$('sale-name').value,customerPhone:$('sale-phone').value,paymentMethod:payment};
  toast('Recording sale...');
  google.script.run
    .withSuccessHandler(r=>{
      if(r.ok){
        toast('✅ Sale '+r.saleId+' recorded!','ok');
        currentItemId='';currentItemCost=0;
        $('sale-picker').value='';$('sale-search').value='';$('sale-price').value='';
        $('sale-name').value='';$('sale-phone').value='';
        $('sale-item-found').style.display='none';$('sale-margin-hint').textContent='';
        loadSaleForm();
      } else toast('Error: '+r.error,'err');
    })
    .withFailureHandler(e=>toast('Error: '+e.message,'err'))
    .logSale(d);
}

// ─── Pending Payments ─────────────────────────────────────────────
function loadPending(){
  google.script.run
    .withSuccessHandler(pending=>{
      const list=$('pend-list'),empty=$('pend-empty'),summary=$('pend-summary');
      list.innerHTML='';
      if(!pending.length){empty.style.display='flex';summary.style.display='none';return;}
      empty.style.display='none';
      const total=pending.reduce((s,p)=>s+p.salePrice,0);
      summary.style.display='flex';
      summary.textContent='⚠️ '+pending.length+' unpaid item'+(pending.length>1?'s':'')+' · Total '+fmt(total);
      pending.forEach(p=>{
        const el=document.createElement('div');el.className='pc';
        el.innerHTML=
          '<div class="pc-id">'+p.itemId+'</div>'+
          '<div class="pc-name">'+p.category+' · '+p.customerName+'</div>'+
          '<div class="pc-amt">'+fmt(p.salePrice)+'</div>'+
          '<div class="pc-foot">'+
            '<span class="badge rd">'+(p.daysAgo===0?'Today':p.daysAgo+' day'+(p.daysAgo===1?'':'s')+' ago')+'</span>'+
            '<span style="font-size:12px;color:var(--gy)">'+(p.customerPhone?'📞 '+p.customerPhone:'No phone')+'</span>'+
          '</div>'+
          '<button class="btn green sm" onclick="openMarkAsPaid(\''+p.saleId+'\','+p.salePrice+',\''+p.customerName+'\',\''+p.itemId+'\')">✅ Mark as Paid</button>';
        list.appendChild(el);
      });
    })
    .withFailureHandler(()=>toast('Error loading pending','err'))
    .getPendingPayments();
}
function openMarkAsPaid(saleId,amt,name,itemId){
  modalSaleId=saleId;selectedPayMethod='';
  $('modal-title').textContent='How did '+name+' pay?';
  $('modal-amt').textContent=fmt(amt);$('modal-sub').textContent=itemId;
  $('modal-confirm-btn').textContent='✅ Confirm · '+fmt(amt);
  $('modal-other').value='';
  document.querySelectorAll('.pay-btn').forEach(b=>b.classList.remove('sel'));
  $('modal-overlay').classList.add('open');
}
function selectPayMethod(btn,method){
  document.querySelectorAll('.pay-btn').forEach(b=>b.classList.remove('sel'));
  btn.classList.add('sel');selectedPayMethod=method;$('modal-other').value='';
}
function confirmMarkAsPaid(){
  const method=$('modal-other').value.trim()||selectedPayMethod;
  if(!method)return toast('Please select a payment method','err');
  google.script.run
    .withSuccessHandler(r=>{
      if(r.ok){toast('✅ Payment recorded as '+method,'ok');closeModalDirect();loadPending();}
      else toast('Error: '+r.error,'err');
    })
    .withFailureHandler(e=>toast('Error: '+e.message,'err'))
    .markAsPaid(modalSaleId,method);
}
function closeModal(e){if(e.target.id==='modal-overlay')closeModalDirect();}
function closeModalDirect(){$('modal-overlay').classList.remove('open');}
function sendReminder(){
  toast('Sending reminder emails...');
  google.script.run
    .withSuccessHandler(r=>{
      if(r.ok)toast('📧 Reminder sent! ('+r.count+' items, '+fmt(r.total)+')','ok');
      else toast(r.msg||'Sent','');
    })
    .withFailureHandler(()=>toast('Error sending email','err'))
    .sendPendingReminder();
}

// ─── Inventory ────────────────────────────────────────────────────
let invFilter='all';
function loadInventory(){
  $('inv-list').innerHTML='<div class="loading">Loading inventory...</div>';
  google.script.run
    .withSuccessHandler(data=>{invData=data;renderInv();})
    .withFailureHandler(()=>toast('Error loading inventory','err'))
    .getInventoryAge();
}
function filterInv(f,el){
  invFilter=f;
  document.querySelectorAll('#inv-filter .tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');renderInv();
}
function renderInv(){
  const list=$('inv-list'),empty=$('inv-empty');
  const filtered=invData.filter(i=>{
    if(invFilter==='30')return i.days>=30;
    if(invFilter==='60')return i.days>=60;
    return true;
  });
  list.innerHTML='';
  if(!filtered.length){empty.style.display='block';return;}
  empty.style.display='none';
  filtered.forEach(i=>{
    const el=document.createElement('div');el.className='ii';
    const retail=i.level==='critical'?
      '<span class="xout">'+fmt(i.retail)+'</span>':
      '<span class="cur">Retail: '+fmt(i.retail)+'</span>';
    el.innerHTML=
      '<div class="ii-top">'+
        '<div><div class="ii-id">'+i.itemId+'</div><div class="ii-name">'+i.category+(i.sku?' · '+i.sku:'')+'</div></div>'+
        '<div class="age-badge '+i.level+'">'+i.days+' day'+(i.days===1?'':'s')+'</div>'+
      '</div>'+
      '<div class="ii-prices">'+retail+'<span style="color:var(--gy)">Cost: '+fmt(i.cost)+'</span></div>'+
      '<div class="sug '+i.level+'">'+i.suggest+'</div>';
    list.appendChild(el);
  });
}

// ─── Reports ──────────────────────────────────────────────────────
function loadReports(){
  $('rp-shipments').innerHTML='<div class="loading">Loading...</div>';
  google.script.run
    .withSuccessHandler(r=>{
      const s=r.summary;
      $('rp-invested').textContent=fmt(s.invested);$('rp-rev').textContent=fmt(s.revenue);
      $('rp-profit').textContent=fmt(s.profit);$('rp-margin').textContent=s.margin+'%';
      $('rp-costoh').textContent=fmt(s.costOH);$('rp-retailoh').textContent=fmt(s.retailOH);
      const h=r.health,tot=h.total||1;
      const fp=Math.round(h.fresh/tot*100),ap=Math.round(h.aging30/tot*100),cp=Math.round(h.aging60/tot*100);
      $('rp-health').innerHTML=
        '<div style="display:flex;flex-direction:column;gap:9px;font-size:12px">'+
        '<div><div class="metric-row" style="border:none;padding:0 0 4px"><span class="mk">✅ Fresh (under 30 days)</span><span class="mv gn">'+h.fresh+' · '+fp+'%</span></div><div class="miniprog"><div class="minifill" style="width:'+fp+'%;background:var(--gn)"></div></div></div>'+
        '<div><div class="metric-row" style="border:none;padding:0 0 4px"><span class="mk">⚠️ Aging (30–60 days)</span><span class="mv" style="color:#C2410C">'+h.aging30+' · '+ap+'%</span></div><div class="miniprog"><div class="minifill" style="width:'+ap+'%;background:#F97316"></div></div></div>'+
        '<div><div class="metric-row" style="border:none;padding:0 0 4px"><span class="mk">🔴 Old (60+ days)</span><span class="mv rd">'+h.aging60+' · '+cp+'%</span></div><div class="miniprog"><div class="minifill" style="width:'+cp+'%;background:var(--rd)"></div></div></div>'+
        '</div>';
      const cont=$('rp-shipments');
      if(!r.byShipment.length){cont.innerHTML='<div class="loading">No shipments yet.</div>';return;}
      cont.innerHTML='';
      r.byShipment.forEach(shp=>{
        const pct=shp.sellThru;
        const barColor=pct>=80?'var(--gn)':pct>=40?'var(--o)':'var(--bl)';
        let status='';
        if(shp.available>0&&shp.oldestDays>=60)status='<div style="font-size:11px;color:var(--rd);font-weight:600">🔴 Oldest item '+shp.oldestDays+' days · Mark down!</div>';
        else if(shp.available>0&&shp.oldestDays>=30)status='<div style="font-size:11px;color:#C2410C;font-weight:600">⚠️ '+shp.available+' items unsold · oldest '+shp.oldestDays+' days</div>';
        else if(shp.received>0)status='<div style="font-size:11px;color:var(--bl);font-weight:600">🔄 '+shp.sellThru+'% sold through</div>';
        const el=document.createElement('div');el.className='sr';
        el.innerHTML=
          '<div class="sr-top">'+
            '<div><div class="sr-id">'+shp.id+'</div><div class="sr-sub">'+shp.supplier+' · '+shp.date+'</div></div>'+
            '<div class="coll-tag">'+shp.collection+'</div>'+
          '</div>'+
          '<div class="sr-stats">'+
            '<div class="srs"><div class="sv">'+shp.received+'</div><div class="sk">Received</div></div>'+
            '<div class="srs"><div class="sv" style="color:var(--gn)">'+shp.sold+'</div><div class="sk">Sold</div></div>'+
            '<div class="srs"><div class="sv" style="color:'+(shp.available>0?'var(--o)':'var(--gy)')+'">'+shp.available+'</div><div class="sk">Left</div></div>'+
          '</div>'+
          '<div class="metric-row" style="padding:4px 0 0;border:none;font-size:12px">'+
            '<span class="mk">Revenue · Cost · Profit</span>'+
            '<span class="mv gn" style="font-size:11px">'+fmt(shp.revenue)+' · '+fmt(shp.cogs)+' · '+fmt(shp.profit)+'</span>'+
          '</div>'+
          '<div class="miniprog"><div class="minifill" style="width:'+pct+'%;background:'+barColor+'"></div></div>'+
          status;
        cont.appendChild(el);
      });
    })
    .withFailureHandler(()=>toast('Error loading reports','err'))
    .getReports();
}

// ─── Init ─────────────────────────────────────────────────────────
window.onload=function(){
  fixHeight();
  loadDashboard();
  const today=new Date().toISOString().split('T')[0];
  if($('shp-date'))$('shp-date').value=today;
};
</script>
</body>
</html>`;
}
