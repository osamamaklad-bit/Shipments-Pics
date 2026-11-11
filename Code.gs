/** Shipment photos + Bosta bulky tracker. Script properties:
 *  APP_PIN, SHEET_ID (if not container-bound), PHOTO_FOLDER_ID (optional),
 *  BOSTA_API_KEY, BOSTA_AUTH_HEADER_NAME (same as your fees script),
 *  BOSTA_DELIVERY_URL (optional, default below, {tn} = tracking number),
 *  SHOPIFY_STORE (xxx.myshopify.com), SHOPIFY_TOKEN (optional, for item SKUs),
 *  FLAG_REGEX (optional, default "bulky"; e.g. "bulky|white bag").
 *  Run installTrigger() once. */
const P = PropertiesService.getScriptProperties();
const DEFAULT_URL = 'https://app.bosta.co/api/v2/deliveries/business/{tn}';
const HEAD = {
  Shipments: ['Time', 'Tracking', 'FileId', 'PhotoURL', 'Note', 'ScannedBy'],
  Tracking: ['Tracking', 'FirstScanned', 'Photos', 'BostaState', 'Size', 'Done'],
  Bulky: ['Tracking', 'Order', 'Size', 'ExpectedFeeVAT', 'ShippingFee', 'SizeCost', 'MaterialFee',
    'FinalBostaFees', 'State', 'Items', 'PhotoLinks', 'DetectedAt', 'FinalizedAt', 'Status']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Shipment photos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
/** JSON API for the hosted PWA (POST body is text/plain JSON: {fn, args}). */
function doPost(e) {
  const fns = { saveShipment: saveShipment, getShipment: getShipment };
  let out;
  try {
    const q = JSON.parse(e.postData.contents);
    if (!fns[q.fn]) throw new Error('Bad request');
    out = { result: fns[q.fn].apply(null, q.args) };
  } catch (err) { out = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function sh_(name) {
  const ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(HEAD[name]); s.setFrozenRows(1); }
  return s;
}
function pin_(p) { if (String(p) !== P.getProperty('APP_PIN')) throw new Error('Wrong PIN'); }
function folder_() {
  let id = P.getProperty('PHOTO_FOLDER_ID');
  if (!id) { id = DriveApp.createFolder('Shipment Photos').getId(); P.setProperty('PHOTO_FOLDER_ID', id); }
  return DriveApp.getFolderById(id);
}
function rows_(name) { const v = sh_(name).getDataRange().getValues(); v.shift(); return v; }
function findRow_(name, tn) {
  const v = sh_(name).getRange(1, 1, sh_(name).getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === tn) return i + 1;
  return 0;
}

/** photos: array of data URLs (jpeg) */
function saveShipment(pin, tn, photos, note) {
  pin_(pin); tn = String(tn).trim();
  if (!/^\d{6,}$/.test(tn)) throw new Error('Tracking number must be digits only');
  if (!photos || !photos.length) throw new Error('Add at least one photo');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const user = Session.getActiveUser().getEmail() || 'app', now = new Date(), f = folder_();
    photos.forEach((d, i) => {
      const blob = Utilities.newBlob(Utilities.base64Decode(d.split(',')[1]), 'image/jpeg',
        tn + '_' + Utilities.formatDate(now, 'Africa/Cairo', 'yyyyMMdd-HHmmss') + '_' + (i + 1) + '.jpg');
      const file = f.createFile(blob);
      sh_('Shipments').appendRow([now, "'" + tn, file.getId(), file.getUrl(), note || '', user]);
    });
    const r = findRow_('Tracking', tn), t = sh_('Tracking');
    if (r) t.getRange(r, 3).setValue(Number(t.getRange(r, 3).getValue()) + photos.length);
    else t.appendRow(["'" + tn, now, photos.length, '', '', false]);
    return { saved: photos.length };
  } finally { lock.releaseLock(); }
}

function getShipment(pin, tn) {
  pin_(pin); tn = String(tn).trim();
  const photos = rows_('Shipments').filter(r => String(r[1]) === tn).slice(-8).map(r => ({
    time: Utilities.formatDate(new Date(r[0]), 'Africa/Cairo', 'dd MMM HH:mm'), note: r[4], url: r[3],
    img: 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(r[2]).getBlob().getBytes())
  }));
  const t = rows_('Tracking').find(r => String(r[0]) === tn);
  const b = rows_('Bulky').find(r => String(r[0]) === tn);
  return { tn, photos, state: t ? t[3] : '', size: t ? t[4] : '',
    bulky: b ? { size: b[2], expected: b[3], final: b[7], status: b[13], items: b[9] } : null };
}

/* ---------- Bosta sync (hourly) ---------- */
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncBosta').forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('syncBosta').timeBased().everyHours(1).create();
}
function bosta_(tn) {
  const url = (P.getProperty('BOSTA_DELIVERY_URL') || DEFAULT_URL).replace('{tn}', tn), h = {};
  h[P.getProperty('BOSTA_AUTH_HEADER_NAME') || 'Authorization'] = P.getProperty('BOSTA_API_KEY');
  const res = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Bosta ' + res.getResponseCode() + ' for ' + tn);
  return JSON.parse(res.getContentText()).data;
}
/** Walk the delivery log oldest->newest; the last pricing values win. */
function pricing_(d) {
  const o = { size: '', ship: '', sizeCost: '', material: '', vat: '' };
  (d.log || []).forEach(e => {
    const a = e.actionsList || {};
    if (a.pricingPackageSize && a.pricingPackageSize.after) o.size = a.pricingPackageSize.after.name;
    if (a.pricing_size_name) o.size = a.pricing_size_name.after;
    const p = a.pricing && a.pricing.after;
    if (p) { o.ship = p.shippingFee; o.sizeCost = p.size && p.size.cost; o.vat = p.priceAfterVat;
      o.material = p.bostaMaterialFee ? p.bostaMaterialFee.amount : ''; }
    if (a.pricing_shippingFee) o.ship = a.pricing_shippingFee.after;
    if (a.pricing_size_cost) o.sizeCost = a.pricing_size_cost.after;
    if (a.pricing_priceAfterVat) o.vat = a.pricing_priceAfterVat.after;
    if (a.pricing_bostaMaterialFee) o.material = a.pricing_bostaMaterialFee.after.amount;
  });
  const cc = d.wallet && d.wallet.cashCycle;
  if (cc && cc.size) o.size = cc.size;           // final word after delivery
  o.final = cc ? Number(cc.bosta_fees) : '';
  return o;
}
function items_(d) {
  const store = P.getProperty('SHOPIFY_STORE'), tok = P.getProperty('SHOPIFY_TOKEN');
  if (store && tok && d.shopifyOrderId) try {
    const r = UrlFetchApp.fetch('https://' + store + '/admin/api/2025-01/orders/' + d.shopifyOrderId + '.json?fields=name,line_items',
      { headers: { 'X-Shopify-Access-Token': tok }, muteHttpExceptions: true });
    const o = JSON.parse(r.getContentText()).order;
    return o.line_items.map(l => (l.sku || '-') + ' | ' + l.title + ' x' + l.quantity).join('\n');
  } catch (e) { /* fall back below */ }
  const pd = d.specs && d.specs.packageDetails;
  return pd ? pd.description + ' (items: ' + pd.itemsCount + ')' : '';
}
/** Pull tracking numbers of ALL shipments from the same Bosta fees feed your fees script uses
 *  (BOSTA_FEES_URL, covering recent days) and add unseen ones to Tracking with 0 photos. */
function importFromFees() {
  const url = P.getProperty('BOSTA_FEES_URL'); if (!url) return;
  const h = {}; h[P.getProperty('BOSTA_AUTH_HEADER_NAME') || 'Authorization'] = P.getProperty('BOSTA_API_KEY');
  const res = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Fees feed ' + res.getResponseCode());
  const found = {};
  (function walk(o) {
    if (Array.isArray(o)) return o.forEach(walk);
    if (o && typeof o === 'object') Object.keys(o).forEach(k => {
      if (/tracking/i.test(k) && /^\d{6,}$/.test(String(o[k]))) found[String(o[k])] = 1; else walk(o[k]);
    });
  })(JSON.parse(res.getContentText()));
  const T = sh_('Tracking'), have = {};
  T.getDataRange().getValues().forEach(r => have[String(r[0])] = 1);
  const add = Object.keys(found).filter(tn => !have[tn]).map(tn => ["'" + tn, new Date(), 0, '', '', false]);
  if (add.length) T.getRange(T.getLastRow() + 1, 1, add.length, 6).setValues(add);
}
function syncBosta() {
  try { importFromFees(); } catch (e) { console.error('importFromFees', e.message); }
  const t0 = Date.now(), flag = new RegExp(P.getProperty('FLAG_REGEX') || 'bulky', 'i');
  const T = sh_('Tracking'), data = T.getDataRange().getValues();
  for (let i = 1; i < data.length && Date.now() - t0 < 300000; i++) {
    if (data[i][5] === true) continue;
    const tn = String(data[i][0]);
    try {
      const d = bosta_(tn), p = pricing_(d), old = (Date.now() - new Date(data[i][1])) > 60 * 864e5;
      T.getRange(i + 1, 4, 1, 3).setValues([[d.state && d.state.value, p.size, p.final !== '' || old]]);
      if (!flag.test(p.size || '')) continue;
      const B = sh_('Bulky'), r = findRow_('Bulky', tn), now = new Date();
      const links = rows_('Shipments').filter(x => String(x[1]) === tn).map(x => x[3]).join('\n');
      const row = [ "'" + tn, d.businessReference, p.size, p.vat, p.ship, p.sizeCost, p.material, p.final,
        d.state && d.state.value, items_(d), links, now, p.final !== '' ? now : '', p.final !== '' ? 'Charged' : 'Flagged' ];
      if (r) { const cur = B.getRange(r, 1, 1, 14).getValues()[0]; row[11] = cur[11]; if (!row[12]) row[12] = cur[12];
        B.getRange(r, 1, 1, 14).setValues([row]); }
      else B.appendRow(row);
    } catch (e) { console.error(tn, e.message); }
  }
}
