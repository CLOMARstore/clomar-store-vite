/* Clomar Store V03.2 — Control gerencial + IA comercial guiada por datos */
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { supabase, hasSupabaseConfig } from './supabaseClient';
import * as XLSX from 'xlsx';
import { LogOut, Menu, Search, ShoppingCart, X } from 'lucide-react';
import './styles.css';

const money = (value) => `S/ ${Number(value || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const todayISO = () => {
  const d = new Date();
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
};
const asNum = (v) => Number(v || 0);
const fmtDate = (value) => value ? new Date(value).toLocaleString('es-PE') : '';
const ticketText = (sale, items = [], payments = []) => {
  const lines = [];
  lines.push('CLOMAR STORE');
  lines.push('Comprobante interno');
  lines.push('------------------------------');
  lines.push(`Boleta: B${sale?.receipt_number || '—'}`);
  lines.push(`Fecha: ${fmtDate(sale?.created_at || new Date())}`);
  lines.push(`Cliente: ${sale?.customer_name || 'Cliente'}`);
  lines.push(`Pago: ${sale?.payment_method || 'Efectivo'}`);
  lines.push('------------------------------');
  items.forEach(it => lines.push(`${it.qty} x ${it.product_name || it.name}  ${money(asNum(it.subtotal) || asNum(it.qty)*asNum(it.price))}`));
  lines.push('------------------------------');
  lines.push(`TOTAL: ${money(sale?.total || 0)}`);
  if (payments.length) {
    lines.push('');
    lines.push('ABONOS:');
    payments.forEach(p => lines.push(`${fmtDate(p.created_at)} · ${p.method || p.payment_method}: ${money(p.amount)}`));
  }
  lines.push('Gracias por su compra.');
  return lines.join('\n');
};
const downloadText = (filename, content) => {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
};


const escapeHtml = (value = '') => String(value ?? '').replace(/[&<>'"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[ch]));
const receiptNumber = (sale) => sale?.receipt_number ? `B${sale.receipt_number}` : `B${String(sale?.id || '').slice(0, 8) || '—'}`;
const publicAssetUrl = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^https?:\/\//i.test(raw) || raw.startsWith('data:')) return raw;
  return `${window.location.origin}${raw.startsWith('/') ? raw : `/${raw}`}`;
};
const receiptQrData = (sale, store) => [
  store?.name || 'Clomar Store',
  `Comprobante ${receiptNumber(sale)}`,
  `Total ${money(sale?.total || 0)}`,
  `Fecha ${fmtDate(sale?.created_at || new Date())}`,
].join(' | ');
const receiptTotals = (sale, items = []) => {
  const subtotal = items.reduce((sum, it) => sum + asNum(it.subtotal || asNum(it.qty) * asNum(it.price)), 0) || asNum(sale?.total);
  return { subtotal, total: asNum(sale?.total || subtotal) };
};
const buildReceiptHTML = ({ sale, items = [], store = {}, profile = {}, format = '80mm' }) => {
  const isA4 = format === 'a4';
  const is58 = format === '58mm';
  const width = isA4 ? '190mm' : is58 ? '54mm' : '76mm';
  const qr = qrUrl(receiptQrData(sale, store));
  const totals = receiptTotals(sale, items);
  const docType = sale?.document_type || 'Interno';
  const docMeta = documentMeta(docType);
  const receiptTitle = docType === 'Boleta' ? 'BOLETA ELECTRÓNICA — PENDIENTE SUNAT' : docType === 'Factura' ? 'FACTURA ELECTRÓNICA — PENDIENTE SUNAT' : sale?.payment_method === 'Crédito' ? 'COMPROBANTE DE CRÉDITO' : 'COMPROBANTE INTERNO';
  const rows = items.map((it) => {
    const qty = asNum(it.qty);
    const price = asNum(it.price);
    const subtotal = asNum(it.subtotal || qty * price);
    return `<tr><td>${escapeHtml(it.product_name || it.name || 'Producto')}</td><td class="num">${qty}</td><td class="num">${money(price)}</td><td class="num">${money(subtotal)}</td></tr>`;
  }).join('') || '<tr><td colspan="4" class="muted">Sin productos registrados.</td></tr>';
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(receiptNumber(sale))}</title><style>
  *{box-sizing:border-box} body{margin:0;background:#fff;color:#111827;font-family:Arial,Helvetica,sans-serif;font-size:${is58?'10px':'11px'}}
  .receipt{width:${width};max-width:100%;margin:0 auto;padding:${isA4?'16mm':'4mm'};background:#fff}
  .center{text-align:center}.logo{width:${isA4?'70px':'42px'};height:${isA4?'70px':'42px'};object-fit:contain;margin:0 auto 4px;display:block}.store{font-weight:900;font-size:${isA4?'24px':'14px'};letter-spacing:.02em}.muted{color:#64748b}.line{border-top:1px dashed #94a3b8;margin:8px 0}.meta{display:grid;gap:2px;margin:6px 0}.meta div{display:flex;justify-content:space-between;gap:8px}.title{font-weight:900;margin:7px 0;text-align:center}table{width:100%;border-collapse:collapse;margin-top:6px}th,td{padding:3px 0;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top}th{font-size:9px;text-transform:uppercase;color:#64748b}.num{text-align:right;white-space:nowrap}.total{font-size:${isA4?'18px':'14px'};font-weight:900}.qr{width:${isA4?'90px':'58px'};height:${isA4?'90px':'58px'};object-fit:contain;margin:8px auto 3px;display:block}.thanks{font-weight:800;margin-top:8px}.a4-grid{display:${isA4?'grid':'block'};grid-template-columns:1fr 1fr;gap:14mm}.a4-box{border:${isA4?'1px solid #e5e7eb':'0'};border-radius:${isA4?'12px':'0'};padding:${isA4?'10mm':'0'}}
  @media print{body{background:#fff}.receipt{margin:0}.no-print{display:none!important}@page{size:${isA4?'A4':'auto'};margin:${isA4?'10mm':'2mm'}}}

@media print{.price strong{white-space:nowrap!important}.composition-two-column .price strong{white-space:nowrap!important}.composition-two-column.price-max .price strong{font-size:13.6px!important}.label.orientation-vertical.vertical-compact{gap:.38mm!important;align-content:start!important}.label.orientation-vertical.vertical-compact .codes{align-items:center!important}.label.orientation-vertical.price-max .price strong{font-size:26px!important}.density-vertical-medium.orientation-vertical{grid-template-rows:3.95mm 11.9mm 14.3mm 25.4mm!important;gap:.38mm!important;padding:1.95mm 2.05mm 1.6mm!important}.density-vertical-medium.orientation-vertical.price-max .price strong{font-size:25.8px!important}.density-vertical-medium.orientation-vertical .price strong{font-size:22.8px!important}.density-vertical-medium.orientation-vertical.mode-barcode .barcode-svg{height:17.4mm!important}.density-vertical-compact.orientation-vertical{grid-template-rows:3.65mm 10.4mm 13.4mm 24.2mm!important;gap:.34mm!important;padding:1.7mm 1.7mm 1.45mm!important}.density-vertical-compact.orientation-vertical .price strong{font-size:20.2px!important}.density-vertical-compact.orientation-vertical.price-max .price strong{font-size:22.8px!important}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-svg{height:16.8mm!important}.density-vertical-large.orientation-vertical{grid-template-rows:5.1mm 18.6mm 18.7mm 32.4mm!important;gap:.58mm!important;padding:2.75mm 2.95mm 2.35mm!important}.density-vertical-large.orientation-vertical.price-max .price strong{font-size:33px!important}.density-vertical-large.orientation-vertical .price strong{font-size:29.2px!important}.label.orientation-vertical.composition-two-column .price{display:flex!important;flex-direction:row!important;align-items:baseline!important}.label.orientation-vertical.composition-two-column .price strong{white-space:nowrap!important;font-size:22px!important}.qr-large .qr-wrap{grid-template-columns:13.8mm!important;grid-template-rows:13.8mm auto!important}.qr-large .qr{width:13.8mm!important;height:13.8mm!important}.qr-max .qr-wrap{grid-template-columns:17mm!important;grid-template-rows:17mm auto!important}.qr-max .qr{width:17mm!important;height:17mm!important}}
</style></head><body><main class="receipt"><section class="${isA4?'a4-box':''}">
  <div class="center"><img class="logo" src="${escapeHtml(publicAssetUrl(store?.logo_url || APP_ICON))}"/><div class="store">${escapeHtml(store?.name || 'Clomar Store')}</div><div class="muted">${escapeHtml(store?.ruc ? `RUC: ${store.ruc}` : '')}</div><div class="muted">${escapeHtml(store?.address || '')}</div><div class="muted">${escapeHtml(store?.phone ? `Tel: ${store.phone}` : '')}</div></div>
  <div class="line"></div><div class="title">${receiptTitle}</div>
  <div class="meta"><div><strong>N°</strong><span>${escapeHtml(receiptNumber(sale))}</span></div><div><strong>Fecha</strong><span>${escapeHtml(fmtDate(sale?.created_at || new Date()))}</span></div><div><strong>Cliente</strong><span>${escapeHtml(sale?.customer_name || 'Cliente')}</span></div><div><strong>Vendedor</strong><span>${escapeHtml(profile?.full_name || sale?.seller_email || profile?.email || 'Usuario')}</span></div><div><strong>Pago</strong><span>${escapeHtml(sale?.payment_method || 'Efectivo')}</span></div><div><strong>Estado</strong><span>${escapeHtml(sale?.status || 'Pagado')}</span></div><div><strong>SUNAT</strong><span>${escapeHtml(sale?.sunat_status || docMeta.status)}</span></div></div>
  <table><thead><tr><th>Producto</th><th class="num">Cant.</th><th class="num">P.U.</th><th class="num">Importe</th></tr></thead><tbody>${rows}</tbody></table>
  <div class="line"></div><div class="meta"><div><strong>Subtotal</strong><span>${money(totals.subtotal)}</span></div><div class="total"><strong>Total</strong><span>${money(totals.total)}</span></div>${sale?.payment_method === 'Crédito' ? `<div><strong>Saldo pendiente</strong><span>${money(totals.total)}</span></div>` : ''}</div>
  <img class="qr" src="${qr}"/><div class="center muted">${escapeHtml(receiptNumber(sale))}</div><div class="center thanks">Gracias por su compra</div><div class="center muted">${docType === 'Interno' ? 'Comprobante interno. No reemplaza comprobante electrónico SUNAT.' : 'Documento preparado para futura integración SUNAT/PSE/OSE. No enviado todavía.'}</div>
</section></main><script>window.onload=()=>setTimeout(()=>window.print(),250)</script></body></html>`;
};
const printReceipt = ({ sale, items = [], store = {}, profile = {}, format = '80mm' }) => {
  const html = buildReceiptHTML({ sale, items, store, profile, format });
  const win = window.open('', '_blank', 'width=420,height=720');
  if (!win) return alert('El navegador bloqueó la ventana de impresión. Permite ventanas emergentes para imprimir el comprobante.');
  win.document.open();
  win.document.write(html);
  win.document.close();
};

const demoProducts = [
  { id: 'demo-1', code: '0001', name: 'Zapatillas deportivas Newton Nimble Leather', category: 'Calzado', price: 380, cost: 220, stock: 5, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
  { id: 'demo-2', code: '0002', name: 'Sombrero para el sol Bora Bora Booney', category: 'Accesorios', price: 80, cost: 40, stock: 2, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
  { id: 'demo-3', code: '0003', name: 'Camisa de popelina de manga larga para hombre', category: 'Ropa', price: 120, cost: 55, stock: 1, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
];

const DEFAULT_STORE_ID = '00000000-0000-0000-0000-000000000001';
const APP_ICON = '/logo-clomar-icon.png';
const APP_LOGO_FULL = '/logo-clomar-full.png';
const APP_VERSION = 'V03.4.4-R4 · Editor visual de etiquetas + vertical personalizable';
const DOCUMENT_TYPES = ['Interno', 'Boleta', 'Factura'];
const documentMeta = (type = 'Interno') => {
  if (type === 'Boleta') return { label: 'Boleta electrónica', series: 'B001', status: 'Pre-emisión', action: 'Registrar boleta pendiente', note: 'Se registrará como pre-emisión. El envío real requerirá un backend seguro y un PSE/OSE.' };
  if (type === 'Factura') return { label: 'Factura electrónica', series: 'F001', status: 'Pre-emisión', action: 'Registrar factura pendiente', note: 'Requiere RUC y razón social. El envío real requerirá un backend seguro y un PSE/OSE.' };
  return { label: 'Comprobante interno', series: 'INT', status: 'Interno', action: 'Registrar venta interna', note: 'Control interno. No reemplaza comprobante electrónico SUNAT.' };
};
const sunatStatusClass = (status = '') => {
  const s = String(status).toLowerCase();
  if (s.includes('acept')) return 'accepted';
  if (s.includes('rechaz')) return 'rejected';
  if (s.includes('pend') || s.includes('pre')) return 'pending';
  return 'internal';
};
const inferDocumentType = (value = '') => {
  const digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 11) return 'RUC';
  if (digits.length === 8) return 'DNI';
  return 'DNI';
};
const cleanDocument = (value = '') => String(value || '').replace(/[^0-9A-Za-z-]/g, '');
const logoSrc = (store) => store?.logo_url || APP_ICON;
const productImageSrc = (product) => product?.image_url || APP_ICON;

const CODE128_PATTERNS = [
  '212222','222122','222221','121223','121322','131222','122213','122312','132212','221213','221312','231212','112232','122132','122231','113222','123122','123221','223211','221132','221231','213212','223112','312131','311222','321122','321221','312212','322112','322211','212123','212321','232121','111323','131123','131321','112313','132113','132311','211313','231113','231311','112133','112331','132131','113123','113321','133121','313121','211331','231131','213113','213311','213131','311123','311321','331121','312113','312311','332111','314111','221411','431111','111224','111422','121124','121421','141122','141221','112214','112412','122114','122411','142112','142211','241211','221114','413111','241112','134111','111242','121142','121241','114212','124112','124211','411212','421112','421211','212141','214121','412121','111143','111341','131141','114113','114311','411113','411311','113141','114131','311141','411131','211412','211214','211232','2331112'
];
const code128Sequence = (value = '') => {
  const text = String(value || '').trim();
  if (!text) return [104, 16, 106];
  const chars = Array.from(text).map((ch) => {
    const code = ch.charCodeAt(0);
    return code >= 32 && code <= 127 ? code - 32 : 16;
  });
  let checksum = 104;
  chars.forEach((code, idx) => { checksum += code * (idx + 1); });
  return [104, ...chars, checksum % 103, 106];
};
function BarcodeSVG({ value, height = 46 }) {
  const sequence = code128Sequence(value);
  let x = 0;
  const bars = [];
  sequence.forEach((code, codeIndex) => {
    const pattern = CODE128_PATTERNS[code] || CODE128_PATTERNS[16];
    let black = true;
    for (const widthChar of pattern) {
      const w = Number(widthChar || 1);
      if (black) bars.push(<rect key={`${codeIndex}-${x}`} x={x} y="0" width={w} height={height} />);
      x += w;
      black = !black;
    }
  });
  return <svg className="barcode-svg" viewBox={`0 0 ${x} ${height}`} preserveAspectRatio="none" role="img" aria-label={`Código de barras ${value}`}>{bars}</svg>;
}

/* Genera el SVG para la ventana de impresión independiente, sin depender del CSS de la aplicación. */
const barcodeSvgMarkup = (value, height = 46) => {
  const sequence = code128Sequence(value);
  let x = 0;
  const bars = [];
  sequence.forEach((code) => {
    const pattern = CODE128_PATTERNS[code] || CODE128_PATTERNS[16];
    let black = true;
    for (const widthChar of pattern) {
      const width = Number(widthChar || 1);
      if (black) bars.push(`<rect x="${x}" y="0" width="${width}" height="${height}"/>`);
      x += width;
      black = !black;
    }
  });
  return `<svg class="barcode-svg" viewBox="0 0 ${x} ${height}" preserveAspectRatio="none" role="img" aria-label="Código de barras ${escapeHtml(value)}">${bars.join('')}</svg>`;
};
const qrUrl = (value) => `https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=8&data=${encodeURIComponent(String(value || ''))}`;
const productScanCode = (product) => String(product?.barcode || product?.code || product?.id || '').trim();
const CATALOG_WHATSAPP_FALLBACK = '51931709871';
/*
  El QR debe apuntar a un dominio público estable, no a un preview temporal de Vercel.
  Prioridad: URL guardada por el dueño > VITE_PUBLIC_CATALOG_URL > dominio actual.
  Si no se configura una URL estable, el módulo muestra una advertencia antes de imprimir QR.
*/
const CATALOG_URL_STORAGE_KEY = 'clomar_public_catalog_url';
const normalizeCatalogBaseUrl = (value = '') => {
  const raw = String(value || '').trim().replace(/\/+$/, '');
  if (!raw) return '';
  const withProtocol = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  return withProtocol.includes('#/catalogo') ? withProtocol : `${withProtocol}/#/catalogo`;
};
const getSavedCatalogBaseUrl = () => {
  try { return normalizeCatalogBaseUrl(localStorage.getItem(CATALOG_URL_STORAGE_KEY) || ''); } catch (_) { return ''; }
};
const RUNTIME_CATALOG_FALLBACK = typeof window !== 'undefined' ? `${window.location.origin}/#/catalogo` : '';
const CONFIGURED_CATALOG_URL = normalizeCatalogBaseUrl(String(import.meta.env.VITE_PUBLIC_CATALOG_URL || '').trim());
const catalogBaseUrl = (override = '') => normalizeCatalogBaseUrl(override) || getSavedCatalogBaseUrl() || CONFIGURED_CATALOG_URL || normalizeCatalogBaseUrl(RUNTIME_CATALOG_FALLBACK);
const hasStableCatalogUrl = (value = '') => Boolean(normalizeCatalogBaseUrl(value) || getSavedCatalogBaseUrl() || CONFIGURED_CATALOG_URL);
const catalogProductUrl = (product, override = '') => `${catalogBaseUrl(override)}?product=${encodeURIComponent(String(product?.id || ''))}`;
const catalogAvailabilityLabel = (value) => ({ 'Disponible': 'Disponible', 'Últimas unidades': 'Últimas unidades', 'Agotado': 'Agotado' }[value] || 'Disponible');
const normalizeWhatsappNumber = (value) => {
  let digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 9 && digits.startsWith('9')) digits = `51${digits}`;
  return digits || CATALOG_WHATSAPP_FALLBACK;
};
const publicWhatsAppLink = (phone, message) => `https://wa.me/${normalizeWhatsappNumber(phone)}?text=${encodeURIComponent(String(message || ''))}`;
const catalogQrValue = (product, override = '') => catalogProductUrl(product, override);

const PRICE_STATUS_OPTIONS = ['Pendiente', 'Validado', 'Revisar'];
const normalizePriceStatus = (value) => {
  const n = normalizeText(value);
  if (['validado', 'validada', 'ok', 'si', 'sí', 'aprobado', 'aprobada'].includes(n)) return 'Validado';
  if (['revisar', 'revision', 'revisión', 'observado', 'observada'].includes(n)) return 'Revisar';
  return 'Pendiente';
};
const productPriceStatus = (product) => product?.price_status || 'Pendiente';


/* =========================================================
   V02.3 — Impresión profesional de etiquetas
   La impresión se genera en una ventana limpia. Así se evita
   que el menú, la barra inferior móvil o estilos del ERP se
   mezclen con la hoja de etiquetas.
   ========================================================= */
const LABEL_LAYOUTS = {
  a4_2x6: { key: 'a4_2x6', label: 'A4 · 2 columnas × 6 filas', paper: 'a4', columns: 2, rows: 6, width: 90, height: 43, gapX: 5, gapY: 3, density: 'showcase' },
  a4_3x7: { key: 'a4_3x7', label: 'A4 · 3 columnas × 7 filas', paper: 'a4', columns: 3, rows: 7, width: 58, height: 37, gapX: 4, gapY: 3, density: 'commercial' },
  a4_4x8: { key: 'a4_4x8', label: 'A4 · 4 columnas × 8 filas', paper: 'a4', columns: 4, rows: 8, width: 43, height: 32, gapX: 2, gapY: 2, density: 'compact' },
  a4_vertical_2x3: { key: 'a4_vertical_2x3', label: 'A4 vertical premium · 2 columnas × 3 filas', paper: 'a4', columns: 2, rows: 3, width: 90, height: 84, gapX: 6, gapY: 4, density: 'vertical-large' },
  a4_vertical_3x4: { key: 'a4_vertical_3x4', label: 'A4 vertical retail · 3 columnas × 4 filas', paper: 'a4', columns: 3, rows: 4, width: 58, height: 63, gapX: 4, gapY: 3, density: 'vertical-medium' },
  a4_vertical_4x4: { key: 'a4_vertical_4x4', label: 'A4 vertical compacto · 4 columnas × 4 filas', paper: 'a4', columns: 4, rows: 4, width: 43, height: 63, gapX: 2, gapY: 3, density: 'vertical-compact' },
  a4_vertical_5x4: { key: 'a4_vertical_5x4', label: 'A4 vertical fino · 5 columnas × 4 filas', paper: 'a4', columns: 5, rows: 4, width: 33, height: 63, gapX: 2, gapY: 3, density: 'vertical-mini' },
  a4_vertical_3x3: { key: 'a4_vertical_3x3', label: 'A4 vertical largo · 3 columnas × 3 filas', paper: 'a4', columns: 3, rows: 3, width: 52, height: 80, gapX: 5, gapY: 5, density: 'vertical' },
  roll_1col: { key: 'roll_1col', label: 'Rollo térmico · 1 columna', paper: 'roll', columns: 1, rows: 1, width: 60, height: 40, gapX: 0, gapY: 2, density: 'commercial' },
};
const LABEL_TEMPLATE_INFO = {
  commercial: { label: 'Venta profesional', help: 'Para ropa, calzado y artículos con precio, variante y código POS.' },
  compact: { label: 'Control / almacén', help: 'Para reposición, caja o artículos pequeños. Prioriza código de barras.' },
  showcase: { label: 'Góndola / exhibición', help: 'Para mostrador. Precio grande y QR para que el cliente vea la ficha pública.' },
  minimal: { label: 'Solo tienda + precio + códigos', help: 'Etiqueta minimalista: muestra solo la tienda, precio, QR catálogo y código de barras.' },
  brand_codes: { label: 'Solo tienda + QR + barras', help: 'Etiqueta vertical fina: solo marca de la tienda, QR catálogo y código de barras.' },
};
const LABEL_USE_INFO = {
  auto: { label: 'Automático por rubro', help: 'Detecta ropa, calzado, hogar o accesorio y adapta los campos.' },
  apparel: { label: 'Ropa y prendas', help: 'Resalta talla, color, precio y código de barras.' },
  footwear: { label: 'Calzado', help: 'Resalta talla, modelo/color, precio y barras para caja.' },
  home: { label: 'Hogar / bazar', help: 'Muestra medida, diseño o característica y QR cuando corresponde.' },
  accessories: { label: 'Accesorios', help: 'Nombre visual, marca/color y precio.' },
  gondola: { label: 'Góndola / cliente', help: 'Etiqueta comercial: precio y QR; no sobrecarga con código interno.' },
  inventory: { label: 'Almacén / control', help: 'Etiqueta operativa: código de barras, SKU y variante.' },
};

const LABEL_VISUAL_PRESETS = {
  retail_auto: { label: 'Automática retail', help: 'Decide rubro, orientación y código según producto.', labelUse: 'auto', orientation: 'auto', composition: 'auto', mode: 'auto', labelTemplate: 'commercial', sheetLayout: 'a4_3x7', priceEmphasis: 'featured', qrEmphasis: 'balanced', verticalDensity: 'compact' },
  ropa_colgante: { label: 'Ropa colgante', help: 'Vertical retail, precio fuerte y barras POS para prendas.', labelUse: 'apparel', orientation: 'vertical', composition: 'classic', mode: 'barcode', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_3x4', priceEmphasis: 'max', qrEmphasis: 'small', verticalDensity: 'compact' },
  ropa_colgante_premium: { label: 'Ropa colgante premium', help: 'Etiqueta vertical grande, más aire y precio protagonista.', labelUse: 'apparel', orientation: 'vertical', composition: 'price', mode: 'barcode', labelTemplate: 'showcase', sheetLayout: 'a4_vertical_2x3', priceEmphasis: 'max', qrEmphasis: 'small', verticalDensity: 'relaxed' },
  ropa_vertical_qr: { label: 'Ropa vertical QR', help: 'Vertical con QR de catálogo y barras POS en dos columnas.', labelUse: 'apparel', orientation: 'vertical', composition: 'two-column', mode: 'both', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_3x4', priceEmphasis: 'max', qrEmphasis: 'large', verticalDensity: 'compact' },
  accesorio_vertical: { label: 'Accesorio vertical', help: 'Para lentes, billeteras o accesorios pequeños con precio visible.', labelUse: 'accessories', orientation: 'vertical', composition: 'price', mode: 'barcode', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_3x4', priceEmphasis: 'max', qrEmphasis: 'small', verticalDensity: 'compact' },
  ropa_vertical_compacta: { label: 'Ropa vertical compacta', help: 'Más contenido útil y menos espacio muerto en formato vertical mediano.', labelUse: 'apparel', orientation: 'vertical', composition: 'classic', mode: 'barcode', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_3x4', priceEmphasis: 'max', qrEmphasis: 'small', verticalDensity: 'compact' },
  ropa_vertical_destacada: { label: 'Ropa vertical destacada', help: 'Precio muy visible y diseño formal para exhibición vertical.', labelUse: 'apparel', orientation: 'vertical', composition: 'price', mode: 'both', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_2x3', priceEmphasis: 'max', qrEmphasis: 'large', verticalDensity: 'compact' },
  calzado_caja: { label: 'Calzado caja', help: 'Horizontal de dos columnas para talla, color, QR y barras.', labelUse: 'footwear', orientation: 'horizontal', composition: 'two-column', mode: 'both', labelTemplate: 'commercial', sheetLayout: 'a4_3x7', priceEmphasis: 'max', qrEmphasis: 'large', verticalDensity: 'compact' },
  hogar_detalle: { label: 'Hogar / bazar detalle', help: 'Dos columnas: precio y QR a un lado, detalle al otro.', labelUse: 'home', orientation: 'horizontal', composition: 'two-column', mode: 'both', labelTemplate: 'commercial', sheetLayout: 'a4_3x7', priceEmphasis: 'featured', qrEmphasis: 'large', verticalDensity: 'compact' },
  gondola_qr: { label: 'Góndola QR premium', help: 'Precio y QR grandes para cliente; sin barras POS.', labelUse: 'gondola', orientation: 'horizontal', composition: 'qr', mode: 'qr', labelTemplate: 'showcase', sheetLayout: 'a4_2x6', priceEmphasis: 'max', qrEmphasis: 'max', verticalDensity: 'compact' },
  almacen_barra: { label: 'Almacén operativo', help: 'Compacta, código escrito y barras grandes para control.', labelUse: 'inventory', orientation: 'horizontal', composition: 'barcode', mode: 'barcode', labelTemplate: 'compact', sheetLayout: 'a4_4x8', priceEmphasis: 'normal', qrEmphasis: 'small', verticalDensity: 'compact' },
  vertical_4x4_compacta: { label: 'Vertical 4 por fila', help: 'Formato vertical compacto con 4 etiquetas por fila.', labelUse: 'apparel', orientation: 'vertical', composition: 'classic', mode: 'barcode', labelTemplate: 'commercial', sheetLayout: 'a4_vertical_4x4', priceEmphasis: 'max', qrEmphasis: 'small', verticalDensity: 'compact' },
  precio_qr_barras: { label: 'Solo tienda + precio + QR + barras', help: 'Etiqueta minimalista con marca, precio protagonista, QR catálogo y código de barras.', labelUse: 'gondola', orientation: 'vertical', composition: 'price', mode: 'both', labelTemplate: 'minimal', sheetLayout: 'a4_vertical_4x4', priceEmphasis: 'max', qrEmphasis: 'large', verticalDensity: 'compact', showPrice: true, showLogo: true, showCodeText: true },
  tienda_qr_barras_vertical: { label: 'Solo tienda + QR + barras', help: 'Para productos donde solo desea una etiqueta vertical fina con la tienda y los códigos.', labelUse: 'inventory', orientation: 'vertical', composition: 'barcode', mode: 'both', labelTemplate: 'brand_codes', sheetLayout: 'a4_vertical_4x4', priceEmphasis: 'normal', qrEmphasis: 'large', verticalDensity: 'compact', showPrice: false, showLogo: true, showCodeText: true },
  vertical_5x4_fina: { label: 'Vertical fina 5 por fila', help: 'Etiqueta vertical delgada con 5 etiquetas por fila para productos pequeños.', labelUse: 'inventory', orientation: 'vertical', composition: 'barcode', mode: 'both', labelTemplate: 'brand_codes', sheetLayout: 'a4_vertical_5x4', priceEmphasis: 'normal', qrEmphasis: 'balanced', verticalDensity: 'compact', showPrice: false, showLogo: true, showCodeText: false },
};
const labelText = (value = '') => normalizeText(value);
const inferredLabelProfile = (product = {}) => {
  const source = labelText([product.category, product.subcategory, product.name, product.brand].filter(Boolean).join(' '));
  if (/(calzado|zapatilla|zapatillas|bota|botas|zapato|sandalia|taco)/.test(source)) return 'footwear';
  if (/(ropa|prenda|camisa|camiseta|polo|polera|casaca|sudadera|boxer|bóxer|calzoncillo|calcetin|calcetines|media|medias|pantalon|pantalón|falda|vestido)/.test(source)) return 'apparel';
  if (/(hogar|alfombra|decoracion|decoración|bazar|utilitario|electro|vaso|taza|cocina|mueble)/.test(source)) return 'home';
  if (/(accesorio|lente|gafa|mochila|billetera|morral|sombrero|llavero|cartera|bolso)/.test(source)) return 'accessories';
  return 'general';
};
const resolveLabelProfile = (product, labelUse = 'auto') => labelUse === 'auto' ? inferredLabelProfile(product) : labelUse;
const resolveLabelTemplate = (template = 'commercial', labelUse = 'auto') => labelUse === 'gondola' ? 'showcase' : labelUse === 'inventory' ? 'compact' : template;
const resolveLabelOrientation = (orientation = 'auto', profile = 'general') => {
  if (orientation !== 'auto') return orientation;
  return profile === 'apparel' ? 'vertical' : 'horizontal';
};
const VERTICAL_LABEL_LAYOUTS = ['a4_vertical_2x3', 'a4_vertical_3x4', 'a4_vertical_3x3', 'a4_vertical_4x4', 'a4_vertical_5x4'];
const resolveAdaptiveLayout = (layoutKey = 'a4_3x7', orientation = 'auto', labelUse = 'auto') => {
  const preferredProfile = labelUse === 'auto' ? 'general' : labelUse;
  const resolvedOrientation = resolveLabelOrientation(orientation, preferredProfile);
  if (resolvedOrientation === 'vertical') return VERTICAL_LABEL_LAYOUTS.includes(layoutKey) ? layoutKey : 'a4_vertical_3x4';
  if (VERTICAL_LABEL_LAYOUTS.includes(layoutKey)) return 'a4_3x7';
  return layoutKey;
};
const resolveLabelComposition = (composition = 'auto', profile = 'general', orientation = 'horizontal', layout = {}) => {
  if (composition !== 'auto') {
    if (composition === 'two-column' && Number(layout.width || 0) < 52) return 'classic';
    return composition;
  }
  if (profile === 'gondola') return 'qr';
  if (profile === 'inventory') return 'barcode';
  if (orientation === 'vertical') return 'classic';
  if ((profile === 'home' || profile === 'footwear') && Number(layout.width || 0) >= 58) return 'two-column';
  return 'classic';
};
const resolveCompositionCodeMode = (baseMode = 'barcode', composition = 'classic', profile = 'general', layout = {}) => {
  if (composition === 'qr') return 'qr';
  if (composition === 'barcode') return 'barcode';
  if (composition === 'two-column' && Number(layout.width || 0) >= 58 && profile !== 'inventory') return 'both';
  if (composition === 'price' && profile === 'gondola') return 'qr';
  return baseMode;
};
const resolveLabelCodeMode = (mode = 'auto', profile = 'general', layout = {}) => {
  if (mode !== 'auto') return mode;
  if (profile === 'gondola') return 'qr';
  if (profile === 'inventory') return 'barcode';
  if (layout.width <= 43) return 'barcode';
  if (profile === 'home' || profile === 'accessories') return 'both';
  return 'barcode';
};
const labelMetaFor = (product = {}, profile = 'general') => {
  const code = product?.code || productScanCode(product);
  const size = String(product?.size || '').trim();
  const color = String(product?.color || '').trim();
  const brand = String(product?.brand || '').trim();
  if (profile === 'apparel') return [size ? `Talla ${size}` : '', color ? `Color ${color}` : '', code].filter(Boolean).join(' · ');
  if (profile === 'footwear') return [size ? `Talla ${size}` : '', color ? color : '', code].filter(Boolean).join(' · ');
  if (profile === 'home') return [size ? `Medida ${size}` : '', color ? `Diseño ${color}` : '', brand].filter(Boolean).join(' · ') || code;
  if (profile === 'accessories') return [brand, color, size].filter(Boolean).join(' · ') || code;
  if (profile === 'gondola') return [product?.category, size].filter(Boolean).join(' · ') || 'Consulte disponibilidad';
  if (profile === 'inventory') return [code, size, color].filter(Boolean).join(' · ');
  return [code, size, color].filter(Boolean).join(' · ') || code;
};

const labelPrintEsc = (value = '') => escapeHtml(String(value ?? ''));
const labelPriceEsc = (value = '') => labelPrintEsc(String(value ?? '').replace(/^S\/\s+/, 'S/\u00A0'));
const splitEvery = (items, chunkSize) => {
  const groups = [];
  for (let i = 0; i < items.length; i += chunkSize) groups.push(items.slice(i, i + chunkSize));
  return groups;
};

const buildLabelsPrintHTML = ({
  items = [],
  store = {},
  mode = 'auto',
  showPrice = true,
  showLogo = true,
  showCodeText = true,
  sheetLayout = 'a4_3x7',
  labelStyle = 'medium',
  labelTemplate = 'commercial',
  labelUse = 'auto',
  orientation = 'auto',
  composition = 'auto',
  catalogUrl = '',
  priceEmphasis = 'featured',
  qrEmphasis = 'balanced',
  verticalDensity = 'compact',
  visualPreset = 'manual',
}) => {
  const layoutKey = resolveAdaptiveLayout(sheetLayout, orientation, labelUse);
  const layout = LABEL_LAYOUTS[layoutKey] || LABEL_LAYOUTS.a4_3x7;
  const perPage = layout.paper === 'a4' ? layout.columns * layout.rows : 1;
  const pages = splitEvery(items, perPage);
  const logo = publicAssetUrl(store?.logo_url || APP_ICON);
  const storeName = store?.name || 'Clomar Store';
  const selectedTemplate = ['commercial', 'compact', 'showcase', 'minimal', 'brand_codes'].includes(labelTemplate) ? labelTemplate : 'commercial';
  const styleClass = labelStyle === 'small' ? 'style-compact' : labelStyle === 'large' ? 'style-detailed' : 'style-standard';
  const compactTitle = (value, max = 56) => {
    const raw = String(value || 'Producto').replace(/\s+/g, ' ').trim();
    if (raw.length <= max) return raw;
    const target = raw.slice(0, max + 1);
    const wordSafe = target.replace(/\s+\S*$/, '').trim();
    return `${wordSafe || raw.slice(0, max).trim()}…`;
  };
  const labelMarkup = ({ product }) => {
    const code = productScanCode(product);
    const profile = resolveLabelProfile(product, labelUse);
    const effectiveOrientation = resolveLabelOrientation(orientation, profile);
    const effectiveComposition = resolveLabelComposition(composition, profile, effectiveOrientation, layout);
    const effectiveTemplate = effectiveComposition === 'qr' ? 'showcase' : resolveLabelTemplate(selectedTemplate, labelUse);
    const baseMode = resolveLabelCodeMode(mode, profile, layout);
    const effectiveMode = resolveCompositionCodeMode(baseMode, effectiveComposition, profile, layout);
    const priceIsReady = productPriceStatus(product) === 'Validado' && Number(product?.price || 0) > 0;
    const hasQr = effectiveMode === 'qr' || effectiveMode === 'both';
    const hasBarcode = effectiveMode === 'barcode' || effectiveMode === 'both';
    const variant = labelMetaFor(product, profile);
    const titleLimit = layout.width <= 43 ? 46 : layout.width <= 58 ? 60 : 80;
    const rawTitle = String(product?.name || 'Producto').replace(/\s+/g, ' ').trim();
    const visibleTitle = compactTitle(rawTitle, titleLimit);
    const titleClass = rawTitle.length > titleLimit ? 'name-long' : rawTitle.length > Math.round(titleLimit * .68) ? 'name-medium' : 'name-short';
    const logoBlock = showLogo ? `<div class="brand"><img src="${labelPrintEsc(logo)}" alt="" onerror="this.style.display='none'"/><span>${labelPrintEsc(storeName)}</span></div>` : '';
    const priceBlock = !showPrice ? '' : priceIsReady
      ? `<div class="price"><span>${profile === 'inventory' ? 'PRECIO REF.' : 'PRECIO'}</span><strong>${labelPriceEsc(money(product.price))}</strong></div>`
      : `<div class="price pending"><strong>PRECIO PENDIENTE</strong></div>`;
    const qrCaption = profile === 'gondola' ? 'Escanea y consulta' : layout.width <= 43 ? 'Catálogo' : 'Ver catálogo';
    const qrBlock = hasQr ? `<div class="qr-wrap"><img class="qr" src="${labelPrintEsc(qrUrl(catalogQrValue(product, catalogUrl)))}" alt="QR catálogo ${labelPrintEsc(code)}"/><small>${qrCaption}</small></div>` : '';
    const barcodeBlock = hasBarcode ? `<div class="barcode-wrap">${barcodeSvgMarkup(code, 46)}${showCodeText ? `<div class="code-text">${labelPrintEsc(code)}</div>` : ''}</div>` : '';
    const meta = `<div class="meta">${labelPrintEsc(variant)}</div>`;
    const mainMarkup = effectiveTemplate === 'minimal'
      ? `${logoBlock}${priceBlock}<div class="codes ${hasQr && hasBarcode ? 'codes-both' : ''}">${qrBlock}${barcodeBlock}</div>`
      : effectiveTemplate === 'brand_codes'
        ? `${logoBlock}<div class="codes ${hasQr && hasBarcode ? 'codes-both' : ''}">${qrBlock}${barcodeBlock}</div>`
        : `${logoBlock}<div class="label-title"><div class="name ${titleClass}">${labelPrintEsc(visibleTitle)}</div>${meta}</div>${priceBlock}<div class="codes ${hasQr && hasBarcode ? 'codes-both' : ''}">${qrBlock}${barcodeBlock}</div>`;
    return `<article class="label template-${effectiveTemplate} profile-${profile} orientation-${effectiveOrientation} composition-${effectiveComposition} ${styleClass} density-${layout.density} mode-${effectiveMode} ${titleClass} price-${priceEmphasis} qr-${qrEmphasis} vertical-${verticalDensity} preset-${visualPreset}">${mainMarkup}</article>`;
  };

  const pagesMarkup = pages.map((pageItems, pageIndex) => {
    const filler = layout.paper === 'a4' ? Array.from({ length: Math.max(0, perPage - pageItems.length) }, (_, i) => `<div class="label empty" aria-hidden="true" data-empty="${i}"></div>`).join('') : '';
    return `<section class="label-page ${layout.paper === 'roll' ? 'roll-page' : 'a4-page'}" data-page="${pageIndex + 1}"><div class="label-grid">${pageItems.map(labelMarkup).join('')}${filler}</div></section>`;
  }).join('');

  const pageCss = layout.paper === 'a4'
    ? `@page { size: A4 portrait; margin: 0; } .label-page{width:210mm;height:297mm;padding:10mm;page-break-after:always;break-after:page;} .label-page:last-child{page-break-after:auto;break-after:auto;}`
    : `@page { size: 62mm auto; margin: 0; } .label-page{width:62mm;min-height:40mm;padding:1mm;page-break-after:always;break-after:page;} .label-page:last-child{page-break-after:auto;break-after:auto;}`;

  return `<!doctype html>
<html lang="es"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Etiquetas comerciales · ${labelPrintEsc(storeName)}</title>
<style>
  :root{--label-w:${layout.width}mm;--label-h:${layout.height}mm;--cols:${layout.columns};--gap-x:${layout.gapX}mm;--gap-y:${layout.gapY}mm;--navy:#14213d;--ink:#111827;--muted:#64748b;--line:#cbd5e1;--blush:#fff4f1;}
  *{box-sizing:border-box} html,body{margin:0;padding:0;background:#fff;color:#000;font-family:Arial,Helvetica,sans-serif} body{print-color-adjust:exact;-webkit-print-color-adjust:exact}
  ${pageCss}
  .label-grid{display:grid;grid-template-columns:repeat(var(--cols),var(--label-w));grid-auto-rows:var(--label-h);column-gap:var(--gap-x);row-gap:var(--gap-y);justify-content:center;align-content:start}.roll-page .label-grid{justify-content:start;grid-template-columns:var(--label-w)}
  .label{width:var(--label-w);height:var(--label-h);overflow:hidden;border:.22mm solid var(--line);border-radius:2.2mm;background:#fff;padding:1.7mm 1.8mm 1.45mm;display:grid;grid-template-rows:3.4mm 7.1mm 6mm 12mm;gap:.5mm;align-content:space-between;text-align:center;break-inside:avoid;page-break-inside:avoid}.label.empty{visibility:hidden}
  .brand{min-height:3.4mm;display:flex;align-items:center;justify-content:center;gap:1mm;color:var(--navy);font-size:7.2px;font-weight:900;line-height:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.brand img{width:3.3mm;height:3.3mm;object-fit:contain;flex:0 0 auto}.label-title{height:7.1mm;overflow:hidden;display:block}.name{display:block;height:5.3mm;max-height:5.3mm;overflow:hidden;word-break:break-word;font-size:9.7px;line-height:1.06;font-weight:950;color:#05070b}.name.name-medium{font-size:9.15px;line-height:1.03}.name.name-long{font-size:8.45px;line-height:1.01;letter-spacing:-.012em}.meta{height:1.35mm;font-size:6.2px;line-height:1.05;font-weight:700;letter-spacing:.025em;color:#64748b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .price{min-height:6mm;display:flex;align-items:center;justify-content:center;gap:1.3mm;padding:.85mm 1.3mm;border-radius:1.25mm;background:var(--navy);color:#fff}.price span{font-size:5.4px;font-weight:900;letter-spacing:.09em}.price strong{font-size:12.3px;line-height:1;font-weight:950}.price.pending{background:#fff3e0;color:#9a3412;font-size:6.2px;letter-spacing:.035em}
  .codes{height:12mm;display:flex;align-items:center;justify-content:center;gap:1.55mm;min-width:0}.codes-both{justify-content:space-between}.qr-wrap{display:grid;grid-template-columns:11.4mm;grid-template-rows:11.4mm auto;place-items:center;gap:.2mm}.qr{width:11.4mm;height:11.4mm;object-fit:contain}.qr-wrap small{font-size:4.7px;line-height:1;color:#64748b;white-space:nowrap}.barcode-wrap{display:grid;grid-template-rows:9.6mm auto;align-items:center;justify-items:center;min-width:0;flex:1}.barcode-svg{width:100%;height:9.6mm;display:block;fill:#000;background:#fff}.codes-both .barcode-wrap{width:calc(var(--label-w) - 21.5mm)}.code-text{margin-top:.35mm;font-size:5.8px;line-height:1;font-weight:850;letter-spacing:.065em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#111827}
  .template-showcase .price{background:#0f172a;border-radius:1.6mm;min-height:7.2mm}.template-showcase .price strong{font-size:15px}.template-showcase .name{font-size:10.6px}.template-minimal{padding:1.25mm 1.25mm 1.1mm;grid-template-rows:3.2mm 8.4mm 13.2mm;gap:.55mm;align-content:center}.template-minimal .brand{min-height:3.2mm;font-size:6.35px;justify-content:center}.template-minimal .brand img{width:2.8mm;height:2.8mm}.template-minimal .price{min-height:8.4mm;border-radius:1.7mm;padding:.8mm 1mm}.template-minimal .price span{font-size:4.7px}.template-minimal .price strong{font-size:14.6px}.template-minimal .codes{height:13.2mm;gap:1mm}.template-minimal .qr-wrap{grid-template-columns:10.4mm;grid-template-rows:10.4mm auto}.template-minimal .qr{width:10.4mm;height:10.4mm}.template-minimal .qr-wrap small{font-size:4.4px;font-weight:800;line-height:1.02}.template-minimal .barcode-wrap{grid-template-rows:8.3mm auto}.template-minimal .barcode-svg{height:8.3mm}.template-minimal .code-text{font-size:4.7px}.template-minimal .codes-both .barcode-wrap{width:auto;flex:1}.template-minimal.orientation-vertical{padding:1.8mm 1.8mm 1.5mm;grid-template-rows:4mm 15mm 25mm;gap:.55mm;border-radius:2.8mm}.template-minimal.orientation-vertical .brand{min-height:4mm;font-size:7.1px}.template-minimal.orientation-vertical .brand img{width:3.4mm;height:3.4mm}.template-minimal.orientation-vertical .price{min-height:15mm;padding:1.05mm 1.1mm}.template-minimal.orientation-vertical .price span{font-size:5.8px}.template-minimal.orientation-vertical .price strong{font-size:22px}.template-minimal.orientation-vertical.price-max .price strong{font-size:25px!important}.template-minimal.orientation-vertical .codes{height:25mm}.template-minimal.orientation-vertical.mode-both .qr-wrap{grid-template-columns:14.5mm;grid-template-rows:14.5mm auto}.template-minimal.orientation-vertical.mode-both .qr{width:14.5mm;height:14.5mm}.template-minimal.orientation-vertical.mode-both .barcode-wrap{grid-template-rows:13.6mm auto}.template-minimal.orientation-vertical.mode-both .barcode-svg{height:13.6mm}.template-brand_codes{padding:1.15mm 1.15mm 1mm;grid-template-rows:3.1mm 12.6mm;gap:.55mm;align-content:center}.template-brand_codes .brand{min-height:3.1mm;font-size:6.15px;justify-content:center}.template-brand_codes .brand img{width:2.7mm;height:2.7mm}.template-brand_codes .codes{height:12.6mm;gap:.9mm;align-items:center}.template-brand_codes .qr-wrap{grid-template-columns:9.8mm;grid-template-rows:9.8mm auto}.template-brand_codes .qr{width:9.8mm;height:9.8mm}.template-brand_codes .qr-wrap small{display:none}.template-brand_codes .barcode-wrap{grid-template-rows:8.1mm auto}.template-brand_codes .barcode-svg{height:8.1mm}.template-brand_codes .code-text{font-size:4.6px}.template-brand_codes .codes-both .barcode-wrap{width:auto;flex:1}.template-brand_codes.orientation-vertical{padding:1.55mm 1.55mm 1.3mm;grid-template-rows:3.8mm 24.8mm;gap:.55mm;border-radius:2.5mm}.template-brand_codes.orientation-vertical .brand{min-height:3.8mm;font-size:6.75px}.template-brand_codes.orientation-vertical .brand img{width:3.15mm;height:3.15mm}.template-brand_codes.orientation-vertical .codes{height:24.8mm;align-items:center}.template-brand_codes.orientation-vertical.mode-both .qr-wrap{grid-template-columns:14.2mm;grid-template-rows:14.2mm auto}.template-brand_codes.orientation-vertical.mode-both .qr{width:14.2mm;height:14.2mm}.template-brand_codes.orientation-vertical.mode-both .barcode-wrap{grid-template-rows:12.9mm auto}.template-brand_codes.orientation-vertical.mode-both .barcode-svg{height:12.9mm}.template-brand_codes.orientation-vertical.mode-qr .qr-wrap{grid-template-columns:23mm;grid-template-rows:23mm auto}.template-brand_codes.orientation-vertical.mode-qr .qr{width:23mm;height:23mm}.template-compact{padding:1.15mm 1.2mm;grid-template-rows:2.7mm 5.3mm 4.8mm 8.6mm;gap:.3mm}.template-compact .brand{font-size:6.1px;min-height:2.7mm}.template-compact .brand img{width:2.7mm;height:2.7mm}.template-compact .label-title{height:5.3mm}.template-compact .name{height:4.1mm;max-height:4.1mm;font-size:7.3px}.template-compact .meta{height:1.1mm;font-size:5.3px}.template-compact .price{min-height:4.8mm;padding:.5mm 1mm}.template-compact .price strong{font-size:9.5px}.template-compact .price span{font-size:4.5px}.template-compact .codes{height:8.6mm;gap:1mm}.template-compact .qr-wrap{grid-template-columns:8.3mm;grid-template-rows:8.3mm auto}.template-compact .qr{width:8.3mm;height:8.3mm}.template-compact .qr-wrap small{display:none}.template-compact .barcode-wrap{grid-template-rows:7.3mm auto}.template-compact .barcode-svg{height:7.3mm}.template-compact .code-text{font-size:4.8px}.template-compact .codes-both .barcode-wrap{width:calc(var(--label-w) - 15mm)}
  /* Impresión estable: evita que Chrome o Adobe PDF colapsen el título cuando la etiqueta es pequeña. */
  .density-compact.template-commercial{padding:1.25mm 1.25mm 1.05mm;grid-template-rows:2.85mm 6.25mm 5.2mm 10.35mm;gap:.28mm;align-content:space-between}.density-compact.template-commercial .brand{min-height:2.85mm;font-size:6.1px}.density-compact.template-commercial .brand img{width:2.55mm;height:2.55mm}.density-compact.template-commercial .label-title{height:6.25mm;display:block;overflow:hidden}.density-compact.template-commercial .name{display:block;min-height:4.85mm;max-height:4.85mm;overflow:hidden;word-break:break-word;font-size:7.15px;line-height:1.01;font-weight:950;letter-spacing:-.008em}.density-compact.template-commercial .name.name-medium{font-size:6.82px;line-height:1.005}.density-compact.template-commercial .name.name-long{font-size:6.32px;line-height:1;letter-spacing:-.015em}.density-compact.template-commercial .meta{height:1.1mm;font-size:4.75px;line-height:1.02}.density-compact.template-commercial .price{min-height:5.2mm;padding:.52mm .85mm;border-radius:1.35mm;box-shadow:inset 0 -.18mm 0 rgba(255,255,255,.1)}.density-compact.template-commercial .price span{font-size:4.15px;letter-spacing:.09em}.density-compact.template-commercial .price strong{font-size:9.55px;line-height:1}.density-compact.template-commercial .codes{height:10.35mm;gap:.92mm}.density-compact.template-commercial .qr-wrap{grid-template-columns:9mm;grid-template-rows:9mm auto}.density-compact.template-commercial .qr{width:9mm;height:9mm}.density-compact.template-commercial .qr-wrap small{display:none}.density-compact.template-commercial .barcode-wrap{grid-template-rows:7.25mm auto}.density-compact.template-commercial .barcode-svg{height:7.25mm}.density-compact.template-commercial .code-text{font-size:4.4px}.density-compact.template-commercial .codes-both .barcode-wrap{width:calc(var(--label-w) - 15mm)}
  /* V03.4.1-R4 — Arquitectura de impresión por formato y plantilla.
     Evita que el diseño de Góndola herede medidas de Comercial Pro y garantiza
     la leyenda "Ver catálogo" junto al QR en los formatos A4. */
  .density-commercial.template-commercial,.density-commercial.template-showcase{padding:1.35mm 1.45mm 1.15mm;grid-template-rows:3.05mm 7.25mm 6.15mm 13.35mm;gap:.34mm;align-content:center}
  .density-commercial.template-commercial .brand,.density-commercial.template-showcase .brand{min-height:3.05mm;font-size:6.6px}.density-commercial.template-commercial .brand img,.density-commercial.template-showcase .brand img{width:2.85mm;height:2.85mm}
  .density-commercial.template-commercial .label-title,.density-commercial.template-showcase .label-title{height:7.25mm}.density-commercial.template-commercial .name,.density-commercial.template-showcase .name{height:5.65mm;max-height:5.65mm;font-size:8.75px;line-height:1.03}.density-commercial.template-commercial .name.name-medium,.density-commercial.template-showcase .name.name-medium{font-size:8.25px}.density-commercial.template-commercial .name.name-long,.density-commercial.template-showcase .name.name-long{font-size:7.65px;line-height:1;letter-spacing:-.014em}.density-commercial.template-commercial .meta,.density-commercial.template-showcase .meta{height:1.25mm;font-size:5.3px}
  .density-commercial.template-commercial .price,.density-commercial.template-showcase .price{min-height:6.15mm;padding:.65mm 1.05mm;border-radius:1.45mm}.density-commercial.template-commercial .price strong{font-size:11.1px}.density-commercial.template-showcase .price strong{font-size:12.15px}.density-commercial.template-commercial .price span,.density-commercial.template-showcase .price span{font-size:4.75px}
  .density-commercial.template-commercial .codes,.density-commercial.template-showcase .codes{height:13.35mm;gap:1.15mm}.density-commercial.template-commercial .qr-wrap,.density-commercial.template-showcase .qr-wrap{grid-template-columns:11.45mm;grid-template-rows:11.45mm 1.35mm;gap:.1mm}.density-commercial.template-commercial .qr,.density-commercial.template-showcase .qr{width:11.45mm;height:11.45mm}.density-commercial.template-commercial .qr-wrap small,.density-commercial.template-showcase .qr-wrap small{display:block;font-size:4.6px;font-weight:800;line-height:1.05;color:#475569}.density-commercial.template-commercial .barcode-wrap,.density-commercial.template-showcase .barcode-wrap{grid-template-rows:8.4mm auto}.density-commercial.template-commercial .barcode-svg,.density-commercial.template-showcase .barcode-svg{height:8.4mm}.density-commercial.template-commercial .code-text,.density-commercial.template-showcase .code-text{font-size:5.15px}.density-commercial.template-commercial .codes-both .barcode-wrap,.density-commercial.template-showcase .codes-both .barcode-wrap{width:auto;flex:1}

  .density-compact.template-commercial,.density-compact.template-showcase{padding:1.05mm 1.15mm .95mm;grid-template-rows:2.75mm 6.45mm 5.35mm 11.25mm;gap:.28mm;align-content:center}
  .density-compact.template-commercial .brand,.density-compact.template-showcase .brand{min-height:2.75mm;font-size:6.0px}.density-compact.template-commercial .brand img,.density-compact.template-showcase .brand img{width:2.45mm;height:2.45mm}
  .density-compact.template-commercial .label-title,.density-compact.template-showcase .label-title{height:6.45mm}.density-compact.template-commercial .name,.density-compact.template-showcase .name{height:5.05mm;max-height:5.05mm;font-size:7.2px;line-height:1.015;letter-spacing:-.009em}.density-compact.template-commercial .name.name-medium,.density-compact.template-showcase .name.name-medium{font-size:6.85px}.density-compact.template-commercial .name.name-long,.density-compact.template-showcase .name.name-long{font-size:6.3px;line-height:1;letter-spacing:-.016em}.density-compact.template-commercial .meta,.density-compact.template-showcase .meta{height:1.15mm;font-size:4.75px;line-height:1.03}
  .density-compact.template-commercial .price,.density-compact.template-showcase .price{min-height:5.35mm;padding:.56mm .85mm;border-radius:1.35mm}.density-compact.template-commercial .price strong{font-size:9.75px}.density-compact.template-showcase .price strong{font-size:10.45px}.density-compact.template-commercial .price span,.density-compact.template-showcase .price span{font-size:4.1px}
  .density-compact.template-commercial .codes,.density-compact.template-showcase .codes{height:11.25mm;gap:.95mm}.density-compact.template-commercial .qr-wrap,.density-compact.template-showcase .qr-wrap{grid-template-columns:9.15mm;grid-template-rows:9.15mm 1.2mm;gap:.1mm}.density-compact.template-commercial .qr,.density-compact.template-showcase .qr{width:9.15mm;height:9.15mm}.density-compact.template-commercial .qr-wrap small,.density-compact.template-showcase .qr-wrap small{display:block;font-size:4.25px;font-weight:800;line-height:1.02;color:#475569;white-space:nowrap}.density-compact.template-commercial .barcode-wrap,.density-compact.template-showcase .barcode-wrap{grid-template-rows:7.15mm auto}.density-compact.template-commercial .barcode-svg,.density-compact.template-showcase .barcode-svg{height:7.15mm}.density-compact.template-commercial .code-text,.density-compact.template-showcase .code-text{font-size:4.35px}.density-compact.template-commercial .codes-both .barcode-wrap,.density-compact.template-showcase .codes-both .barcode-wrap{width:auto;flex:1}

  .density-showcase.template-commercial,.density-showcase.template-showcase{padding:1.65mm 1.8mm 1.4mm;grid-template-rows:3.7mm 8.35mm 7.15mm 15.4mm;gap:.42mm;align-content:center}.density-showcase.template-commercial .brand,.density-showcase.template-showcase .brand{min-height:3.7mm;font-size:7.35px}.density-showcase.template-commercial .label-title,.density-showcase.template-showcase .label-title{height:8.35mm}.density-showcase.template-commercial .name,.density-showcase.template-showcase .name{height:6.65mm;max-height:6.65mm;font-size:10.15px;line-height:1.04}.density-showcase.template-commercial .meta,.density-showcase.template-showcase .meta{height:1.3mm;font-size:5.75px}.density-showcase.template-commercial .price,.density-showcase.template-showcase .price{min-height:7.15mm}.density-showcase.template-commercial .price strong{font-size:14.2px}.density-showcase.template-showcase .price strong{font-size:15.6px}.density-showcase.template-commercial .codes,.density-showcase.template-showcase .codes{height:15.4mm}.density-showcase.template-commercial .qr-wrap,.density-showcase.template-showcase .qr-wrap{grid-template-columns:13.6mm;grid-template-rows:13.6mm 1.4mm}.density-showcase.template-commercial .qr,.density-showcase.template-showcase .qr{width:13.6mm;height:13.6mm}.density-showcase.template-commercial .qr-wrap small,.density-showcase.template-showcase .qr-wrap small{display:block;font-size:5.2px;font-weight:800;color:#475569}.density-showcase.template-commercial .barcode-wrap,.density-showcase.template-showcase .barcode-wrap{grid-template-rows:10mm auto}.density-showcase.template-commercial .barcode-svg,.density-showcase.template-showcase .barcode-svg{height:10mm}.density-showcase.template-commercial .codes-both .barcode-wrap,.density-showcase.template-showcase .codes-both .barcode-wrap{width:auto;flex:1}

  .mode-qr .qr-wrap{grid-template-columns:16mm;grid-template-rows:16mm auto}.mode-qr .qr{width:16mm;height:16mm}.mode-qr .codes{height:17mm}.mode-barcode .barcode-wrap{width:100%;grid-template-rows:11mm auto}.mode-barcode .barcode-svg{height:11mm}.mode-barcode .codes{height:12.5mm}
  .density-compact.mode-qr .codes{height:11.25mm}.density-compact.mode-qr .qr-wrap{grid-template-columns:9.8mm;grid-template-rows:9.8mm 1.2mm;gap:.1mm}.density-compact.mode-qr .qr{width:9.8mm;height:9.8mm}.density-compact.mode-qr .qr-wrap small{display:block;font-size:4.25px;font-weight:800}.density-compact.mode-barcode .codes{height:11.25mm}.density-compact.mode-barcode .barcode-wrap{grid-template-rows:8.6mm auto}.density-compact.mode-barcode .barcode-svg{height:8.6mm}
  /* V03.4.4 — Editor visual: precio, QR y composición estable. */
  .price strong{white-space:nowrap}.price-normal .price strong{font-size:11.4px}.price-featured .price strong{font-size:13.8px}.price-max .price strong{font-size:15.8px}.template-showcase.price-max .price strong{font-size:19px}.qr-small .qr-wrap{grid-template-columns:9.5mm;grid-template-rows:9.5mm auto}.qr-small .qr{width:9.5mm;height:9.5mm}.qr-large .qr-wrap{grid-template-columns:13.8mm;grid-template-rows:13.8mm auto}.qr-large .qr{width:13.8mm;height:13.8mm}.qr-max .qr-wrap{grid-template-columns:18mm;grid-template-rows:18mm auto}.qr-max .qr{width:18mm;height:18mm}.qr-max .codes{height:19.5mm}.qr-max.mode-qr .qr-wrap small{font-size:5.8px;font-weight:900}.composition-two-column{grid-template-columns:39% 1fr;align-content:start}.composition-two-column .price{overflow:hidden;min-width:0}.composition-two-column .price strong{font-size:12.7px;white-space:nowrap;letter-spacing:-.02em}.composition-two-column.price-max .price strong{font-size:13.8px}.composition-two-column .price span{white-space:nowrap}.composition-two-column.qr-large .qr-wrap{grid-template-columns:11.8mm;grid-template-rows:11.8mm auto}.composition-two-column.qr-large .qr{width:11.8mm;height:11.8mm}.composition-two-column.qr-max .qr-wrap{grid-template-columns:13.5mm;grid-template-rows:13.5mm auto}.composition-two-column.qr-max .qr{width:13.5mm;height:13.5mm}.composition-two-column .codes{align-items:center;justify-content:space-between}
  /* V03.4.3 — Etiquetas adaptativas: orientación y composición profesional. */
  .label.orientation-vertical{padding:2.4mm 2.55mm 2.15mm;grid-template-rows:4.2mm 15mm 11mm 31.5mm;gap:.9mm;border-radius:3mm;align-content:start}.label.orientation-vertical .brand{min-height:4.2mm;font-size:8px}.label.orientation-vertical .brand img{width:4mm;height:4mm}.label.orientation-vertical .label-title{height:15mm}.label.orientation-vertical .name{height:11.2mm;max-height:11.2mm;font-size:12.1px;line-height:1.05}.label.orientation-vertical .name.name-medium{font-size:11px}.label.orientation-vertical .name.name-long{font-size:9.8px}.label.orientation-vertical .meta{height:2.25mm;font-size:7px}.label.orientation-vertical .price{min-height:11mm;border-radius:2mm;padding:1.1mm}.label.orientation-vertical .price span{font-size:6.2px}.label.orientation-vertical .price strong{font-size:21px}.label.orientation-vertical.price-max .price strong{font-size:24px}.label.orientation-vertical .codes{height:31.5mm;align-items:start;justify-content:center}.label.orientation-vertical .barcode-wrap{width:100%;grid-template-rows:18mm auto}.label.orientation-vertical .barcode-svg{height:18mm}.label.orientation-vertical .code-text{font-size:7.2px;margin-top:.7mm}.label.orientation-vertical.vertical-compact{grid-template-rows:4mm 14.2mm 11.2mm 28.5mm;gap:.75mm}.label.orientation-vertical.vertical-compact .label-title{height:14.2mm}.label.orientation-vertical.vertical-compact .codes{height:28.5mm;align-items:start}.label.orientation-vertical.mode-both .codes{align-items:start}.label.orientation-vertical.mode-both .qr-wrap{grid-template-columns:16mm;grid-template-rows:16mm auto}.label.orientation-vertical.mode-both .qr{width:16mm;height:16mm}.label.orientation-vertical.mode-both .barcode-wrap{width:calc(var(--label-w) - 22mm);grid-template-rows:14mm auto}.label.orientation-vertical.mode-both .barcode-svg{height:14mm}.label.orientation-vertical.mode-qr .codes{align-items:start}.label.orientation-vertical.mode-qr .qr-wrap{grid-template-columns:29mm;grid-template-rows:29mm auto}.label.orientation-vertical.mode-qr .qr{width:29mm;height:29mm}.label.orientation-vertical.mode-qr .qr-wrap small{font-size:6.3px;font-weight:900}
  .label.composition-two-column{grid-template-columns:34% 1fr;grid-template-rows:3.5mm 13.2mm 11.8mm;grid-template-areas:"brand brand" "price title" "codes codes";column-gap:1.4mm;row-gap:.65mm;text-align:left}.label.composition-two-column .brand{grid-area:brand;justify-content:flex-start}.label.composition-two-column .label-title{grid-area:title;height:13.2mm;padding-top:.2mm}.label.composition-two-column .name{height:9.8mm;max-height:9.8mm;font-size:10.25px;line-height:1.04}.label.composition-two-column .meta{height:2.6mm;font-size:6.1px}.label.composition-two-column .price{grid-area:price;display:grid;align-content:center;justify-items:center;gap:.5mm;min-height:13.2mm;border-radius:1.8mm;padding:1mm}.label.composition-two-column .price span{font-size:5px}.label.composition-two-column .price strong{font-size:13.4px}.label.composition-two-column .codes{grid-area:codes;height:11.8mm}.label.composition-two-column .qr-wrap{grid-template-columns:10.2mm;grid-template-rows:10.2mm auto}.label.composition-two-column .qr{width:10.2mm;height:10.2mm}.label.composition-two-column .barcode-wrap{grid-template-rows:8mm auto}.label.composition-two-column .barcode-svg{height:8mm}.label.composition-two-column .codes-both .barcode-wrap{width:auto;flex:1}

  /* V03.4.4-R2 — Vertical retail experto.
     Reduce espacios muertos, mejora el marco y convierte el precio
     en el protagonista real del formato vertical. */
  .label.orientation-vertical{
    padding:2.05mm 2.1mm 1.7mm;
    grid-template-rows:4mm 12.4mm 14.8mm 26.3mm;
    gap:.48mm;
    border:.22mm solid #cbd5e1;
    border-radius:2.8mm;
    box-shadow:inset 0 0 0 .16mm #ffffff;
    background:linear-gradient(180deg,#ffffff 0%,#ffffff 77%,#f8fbff 100%);
    align-content:start;
  }
  .label.orientation-vertical .brand{min-height:4mm;font-size:7.25px;justify-content:center}
  .label.orientation-vertical .brand img{width:3.5mm;height:3.5mm}
  .label.orientation-vertical .label-title{height:12.4mm;display:flex;flex-direction:column;justify-content:flex-start}
  .label.orientation-vertical .name{height:9.5mm;max-height:9.5mm;font-size:11px;line-height:1.03;font-weight:950;letter-spacing:-.012em}
  .label.orientation-vertical .name.name-medium{font-size:10.35px}
  .label.orientation-vertical .name.name-long{font-size:9.45px;line-height:1.01}
  .label.orientation-vertical .meta{height:2.15mm;font-size:6.6px;line-height:1.05}
  .label.orientation-vertical .price{
    min-height:14.8mm;
    border-radius:2.2mm;
    padding:1.2mm 1.4mm;
    background:#0f172a;
    color:#fff;
    display:flex;
    flex-wrap:nowrap;
    align-items:baseline;
    justify-content:center;
    gap:1.25mm;
    box-shadow:0 .7mm 1.4mm rgba(15,23,42,.16), inset 0 -.18mm 0 rgba(255,255,255,.12);
  }
  .label.orientation-vertical .price span{font-size:6.2px;font-weight:900;letter-spacing:.08em;white-space:nowrap;flex:0 0 auto}
  .label.orientation-vertical .price strong{font-size:24px;line-height:1;font-weight:950;white-space:nowrap!important;letter-spacing:-.035em;flex:0 0 auto}
  .label.orientation-vertical.price-max .price strong{font-size:27.8px!important}
  .label.orientation-vertical .codes{height:26.3mm;align-items:center;justify-content:center}
  .label.orientation-vertical .barcode-wrap{width:100%;grid-template-rows:17.2mm auto}
  .label.orientation-vertical .barcode-svg{height:17.2mm}
  .label.orientation-vertical .code-text{font-size:6.85px;margin-top:.55mm}
  .label.orientation-vertical.vertical-compact{padding:1.9mm 1.95mm 1.55mm;grid-template-rows:3.85mm 11.7mm 14.1mm 25mm;gap:.42mm}
  .label.orientation-vertical.vertical-compact .label-title{height:11.7mm}
  .label.orientation-vertical.vertical-compact .name{height:8.9mm;max-height:8.9mm;font-size:10.5px}
  .label.orientation-vertical.vertical-compact .meta{font-size:6.3px}
  .label.orientation-vertical.vertical-compact .price{min-height:14.1mm;padding:1.1mm 1.25mm}
  .label.orientation-vertical.vertical-compact .price strong{font-size:23px}
  .label.orientation-vertical.vertical-compact.price-max .price strong{font-size:26px!important}
  .label.orientation-vertical.vertical-compact .codes{height:25mm}
  .label.orientation-vertical.mode-both .codes{align-items:center}
  .label.orientation-vertical.mode-both .qr-wrap{grid-template-columns:15mm;grid-template-rows:15mm auto}
  .label.orientation-vertical.mode-both .qr{width:15mm;height:15mm}
  .label.orientation-vertical.mode-both .barcode-wrap{width:calc(var(--label-w) - 20.5mm);grid-template-rows:13.8mm auto}
  .label.orientation-vertical.mode-both .barcode-svg{height:13.8mm}
  .label.orientation-vertical.mode-qr .qr-wrap{grid-template-columns:26mm;grid-template-rows:26mm auto}
  .label.orientation-vertical.mode-qr .qr{width:26mm;height:26mm}
  .label.orientation-vertical.mode-qr .qr-wrap small{font-size:6px;font-weight:900}
  .density-vertical-medium.orientation-vertical{
    padding:1.95mm 2.05mm 1.6mm;
    grid-template-rows:3.95mm 11.9mm 14.3mm 25.4mm;
    gap:.42mm;
    border-radius:2.7mm;
  }
  .density-vertical-medium.orientation-vertical .brand{min-height:3.95mm;font-size:7px}
  .density-vertical-medium.orientation-vertical .brand img{width:3.25mm;height:3.25mm}
  .density-vertical-medium.orientation-vertical .label-title{height:11.9mm}
  .density-vertical-medium.orientation-vertical .name{height:9.1mm;max-height:9.1mm;font-size:10.55px;line-height:1.02}
  .density-vertical-medium.orientation-vertical .name.name-medium{font-size:10px}
  .density-vertical-medium.orientation-vertical .name.name-long{font-size:9.05px}
  .density-vertical-medium.orientation-vertical .meta{height:2.05mm;font-size:6.25px}
  .density-vertical-medium.orientation-vertical .price{min-height:14.3mm;border-radius:2mm}
  .density-vertical-medium.orientation-vertical .price strong{font-size:22.8px}
  .density-vertical-medium.orientation-vertical.price-max .price strong{font-size:25.8px!important}
  .density-vertical-medium.orientation-vertical .codes{height:25.4mm;align-items:center}
  .density-vertical-medium.orientation-vertical.mode-barcode .barcode-wrap{grid-template-rows:17.5mm auto}
  .density-vertical-medium.orientation-vertical.mode-barcode .barcode-svg{height:17.5mm}
  .density-vertical-medium.orientation-vertical.mode-both .qr-wrap{grid-template-columns:14.7mm;grid-template-rows:14.7mm auto}
  .density-vertical-medium.orientation-vertical.mode-both .qr{width:14.7mm;height:14.7mm}
  .density-vertical-medium.orientation-vertical.mode-both .barcode-wrap{grid-template-rows:13.5mm auto}
  .density-vertical-medium.orientation-vertical.mode-both .barcode-svg{height:13.5mm}
  .density-vertical-compact.orientation-vertical{padding:1.7mm 1.7mm 1.45mm;grid-template-rows:3.65mm 10.4mm 13.4mm 24.2mm;gap:.38mm;border-radius:2.45mm;align-content:start}.density-vertical-compact.orientation-vertical .brand{min-height:3.65mm;font-size:6.55px}.density-vertical-compact.orientation-vertical .brand img{width:3mm;height:3mm}.density-vertical-compact.orientation-vertical .label-title{height:10.4mm}.density-vertical-compact.orientation-vertical .name{height:7.95mm;max-height:7.95mm;font-size:9.55px;line-height:1.02}.density-vertical-compact.orientation-vertical .name.name-medium{font-size:9px}.density-vertical-compact.orientation-vertical .name.name-long{font-size:8.2px}.density-vertical-compact.orientation-vertical .meta{height:1.95mm;font-size:5.8px}.density-vertical-compact.orientation-vertical .price{min-height:13.4mm;border-radius:1.9mm;padding:1mm 1.05mm}.density-vertical-compact.orientation-vertical .price strong{font-size:20.2px}.density-vertical-compact.orientation-vertical.price-max .price strong{font-size:22.8px!important}.density-vertical-compact.orientation-vertical .price span{font-size:5.2px}.density-vertical-compact.orientation-vertical .codes{height:24.2mm;align-items:center}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-wrap{grid-template-rows:16.8mm auto}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-svg{height:16.8mm}.density-vertical-compact.orientation-vertical.mode-both .qr-wrap{grid-template-columns:13.5mm;grid-template-rows:13.5mm auto}.density-vertical-compact.orientation-vertical.mode-both .qr{width:13.5mm;height:13.5mm}.density-vertical-compact.orientation-vertical.mode-both .barcode-wrap{width:calc(var(--label-w) - 18.2mm);grid-template-rows:12.6mm auto}.density-vertical-compact.orientation-vertical.mode-both .barcode-svg{height:12.6mm}.density-vertical-compact.orientation-vertical.mode-qr .qr-wrap{grid-template-columns:23mm;grid-template-rows:23mm auto}.density-vertical-compact.orientation-vertical.mode-qr .qr{width:23mm;height:23mm}.density-vertical-compact.orientation-vertical.mode-qr .qr-wrap small{font-size:5.5px;font-weight:800}.density-vertical-mini.orientation-vertical{padding:1.45mm 1.25mm 1.2mm;grid-template-rows:3.4mm 9.2mm 12.6mm 23.4mm;gap:.34mm;border-radius:2.2mm;align-content:start}.density-vertical-mini.orientation-vertical .brand{min-height:3.4mm;font-size:5.95px}.density-vertical-mini.orientation-vertical .brand img{width:2.75mm;height:2.75mm}.density-vertical-mini.orientation-vertical .label-title{height:9.2mm}.density-vertical-mini.orientation-vertical .name{height:7.1mm;max-height:7.1mm;font-size:8.45px;line-height:1.01}.density-vertical-mini.orientation-vertical .name.name-medium{font-size:8px}.density-vertical-mini.orientation-vertical .name.name-long{font-size:7.35px}.density-vertical-mini.orientation-vertical .meta{height:1.75mm;font-size:5.15px}.density-vertical-mini.orientation-vertical .price{min-height:12.6mm;border-radius:1.7mm;padding:.9mm .9mm}.density-vertical-mini.orientation-vertical .price strong{font-size:17.2px}.density-vertical-mini.orientation-vertical.price-max .price strong{font-size:19.8px!important}.density-vertical-mini.orientation-vertical .price span{font-size:4.8px}.density-vertical-mini.orientation-vertical .codes{height:23.4mm;align-items:center}.density-vertical-mini.orientation-vertical.mode-barcode .barcode-wrap{grid-template-rows:15.5mm auto}.density-vertical-mini.orientation-vertical.mode-barcode .barcode-svg{height:15.5mm}.density-vertical-mini.orientation-vertical.mode-both .qr-wrap{grid-template-columns:12mm;grid-template-rows:12mm auto}.density-vertical-mini.orientation-vertical.mode-both .qr{width:12mm;height:12mm}.density-vertical-mini.orientation-vertical.mode-both .barcode-wrap{width:calc(var(--label-w) - 16.4mm);grid-template-rows:11.4mm auto}.density-vertical-mini.orientation-vertical.mode-both .barcode-svg{height:11.4mm}.density-vertical-mini.orientation-vertical.mode-qr .qr-wrap{grid-template-columns:20mm;grid-template-rows:20mm auto}.density-vertical-mini.orientation-vertical.mode-qr .qr{width:20mm;height:20mm}.density-vertical-mini.orientation-vertical.mode-qr .qr-wrap small{font-size:5px;font-weight:800}.template-brand_codes.density-vertical-mini.orientation-vertical{grid-template-rows:3.45mm 23.6mm!important;padding:1.35mm 1.15mm 1.1mm!important}.template-brand_codes.density-vertical-mini.orientation-vertical .codes{height:23.6mm!important}.template-brand_codes.density-vertical-mini.orientation-vertical.mode-both .qr-wrap{grid-template-columns:11.6mm!important;grid-template-rows:11.6mm auto!important}.template-brand_codes.density-vertical-mini.orientation-vertical.mode-both .qr{width:11.6mm!important;height:11.6mm!important}.template-brand_codes.density-vertical-mini.orientation-vertical.mode-both .barcode-wrap{grid-template-rows:11mm auto!important}.template-brand_codes.density-vertical-mini.orientation-vertical.mode-both .barcode-svg{height:11mm!important}
  .density-vertical-large.orientation-vertical{
    padding:2.75mm 2.95mm 2.35mm;
    grid-template-rows:5.1mm 18.6mm 18.7mm 32.4mm;
    gap:.62mm;
    border-radius:3.45mm;
  }
  .density-vertical-large.orientation-vertical .brand{min-height:5.1mm;font-size:8.45px}
  .density-vertical-large.orientation-vertical .brand img{width:4.1mm;height:4.1mm}
  .density-vertical-large.orientation-vertical .label-title{height:18.6mm}
  .density-vertical-large.orientation-vertical .name{height:14.6mm;max-height:14.6mm;font-size:13.55px;line-height:1.04}
  .density-vertical-large.orientation-vertical .name.name-medium{font-size:12.85px}
  .density-vertical-large.orientation-vertical .name.name-long{font-size:11.8px}
  .density-vertical-large.orientation-vertical .meta{height:2.65mm;font-size:7.2px}
  .density-vertical-large.orientation-vertical .price{min-height:18.7mm;border-radius:2.5mm}
  .density-vertical-large.orientation-vertical .price strong{font-size:29.5px}
  .density-vertical-large.orientation-vertical.price-max .price strong{font-size:33.5px!important}
  .density-vertical-large.orientation-vertical .price span{font-size:7px}
  .density-vertical-large.orientation-vertical .codes{height:32.4mm;align-items:center}
  .density-vertical-large.orientation-vertical.mode-barcode .barcode-wrap{grid-template-rows:21mm auto}
  .density-vertical-large.orientation-vertical.mode-barcode .barcode-svg{height:21mm}
  .density-vertical-large.orientation-vertical.mode-both .qr-wrap{grid-template-columns:19mm;grid-template-rows:19mm auto}
  .density-vertical-large.orientation-vertical.mode-both .qr{width:19mm;height:19mm}
  .density-vertical-large.orientation-vertical.mode-both .barcode-wrap{grid-template-rows:17.8mm auto}
  .density-vertical-large.orientation-vertical.mode-both .barcode-svg{height:17.8mm}
  .label.orientation-vertical.composition-two-column{
    grid-template-columns:21.5mm 1fr!important;
    grid-template-rows:3.9mm 13.7mm 25.2mm!important;
    grid-template-areas:"brand brand" "price title" "codes codes"!important;
    column-gap:1.1mm!important;
    row-gap:.42mm!important;
    text-align:left!important;
  }
  .label.orientation-vertical.composition-two-column .brand{grid-area:brand!important;justify-content:center!important}
  .label.orientation-vertical.composition-two-column .label-title{grid-area:title!important;height:13.7mm!important;text-align:left!important;padding-top:.2mm}
  .label.orientation-vertical.composition-two-column .name{height:10.15mm!important;max-height:10.15mm!important;font-size:10.35px!important;line-height:1.02!important}
  .label.orientation-vertical.composition-two-column .meta{font-size:6.2px!important;height:2.1mm!important}
  .label.orientation-vertical.composition-two-column .price{
    grid-area:price!important;
    min-height:13.7mm!important;
    display:flex!important;
    flex-direction:row!important;
    align-items:baseline!important;
    justify-content:center!important;
    gap:1mm!important;
    padding:1mm .8mm!important;
  }
  .label.orientation-vertical.composition-two-column .price span{font-size:5.9px!important}
  .label.orientation-vertical.composition-two-column .price strong{font-size:19.8px!important;white-space:nowrap!important}
  .label.orientation-vertical.composition-two-column.price-max .price strong{font-size:22.5px!important}
  .label.orientation-vertical.composition-two-column .codes{
    grid-area:codes!important;
    height:25.2mm!important;
    display:flex!important;
    align-items:center!important;
    justify-content:space-between!important;
    gap:1mm!important;
  }
  .label.orientation-vertical.composition-two-column .codes-both .barcode-wrap{width:auto!important;flex:1!important}

  .label.composition-price .price{min-height:8.4mm;border-radius:1.8mm;box-shadow:inset 0 -.25mm 0 rgba(255,255,255,.12)}.label.composition-price .price strong{font-size:17px}.label.composition-price .price span{font-size:5.7px}.label.composition-price .name{font-size:9px}.label.composition-qr .price{min-height:8.4mm}.label.composition-qr .price strong{font-size:17px}.label.composition-qr .codes{height:16.4mm}.label.composition-qr .qr-wrap{grid-template-columns:15.2mm;grid-template-rows:15.2mm auto}.label.composition-qr .qr{width:15.2mm;height:15.2mm}.label.composition-qr .qr-wrap small{display:block;font-size:5.3px;font-weight:900}.label.composition-barcode .barcode-wrap{width:100%;grid-template-rows:12mm auto}.label.composition-barcode .barcode-svg{height:12mm}
  @media screen{body{background:#f3f4f6;padding:12mm}.label-page{margin:0 auto 12mm;box-shadow:0 2mm 8mm rgba(15,23,42,.18)}}@media print{html,body{width:100%;height:auto;background:#fff}.label-page{box-shadow:none;margin:0}.label{border-color:#aeb5bf}.label-title{display:block!important;overflow:hidden!important}.name{display:block!important;-webkit-line-clamp:unset!important;-webkit-box-orient:initial!important;overflow:hidden!important}.density-compact.template-commercial{padding:1.2mm 1.2mm 1mm!important;grid-template-rows:2.75mm 6.05mm 5.05mm 10.05mm!important;gap:.22mm!important;align-content:space-between!important}.density-compact.template-commercial .brand{min-height:2.75mm!important;font-size:5.95px!important}.density-compact.template-commercial .brand img{width:2.45mm!important;height:2.45mm!important}.density-compact.template-commercial .label-title{height:6.05mm!important}.density-compact.template-commercial .name{min-height:4.72mm!important;max-height:4.72mm!important;font-size:6.95px!important;line-height:1!important}.density-compact.template-commercial .name.name-medium{font-size:6.62px!important}.density-compact.template-commercial .name.name-long{font-size:6.12px!important;letter-spacing:-.016em!important}.density-compact.template-commercial .meta{height:1.05mm!important;font-size:4.62px!important}.density-compact.template-commercial .price{min-height:5.05mm!important;padding:.46mm .76mm!important}.density-compact.template-commercial .price strong{font-size:9.28px!important}.density-compact.template-commercial .price span{font-size:4.02px!important}.density-compact.template-commercial .codes{height:10.05mm!important;gap:.84mm!important}.density-compact.template-commercial .qr-wrap{grid-template-columns:8.75mm!important;grid-template-rows:8.75mm auto!important}.density-compact.template-commercial .qr{width:8.75mm!important;height:8.75mm!important}.density-compact.template-commercial .barcode-wrap{grid-template-rows:7.05mm auto!important}.density-compact.template-commercial .barcode-svg{height:7.05mm!important}.density-compact.template-commercial .code-text{font-size:4.28px!important}.density-compact.template-commercial .codes-both .barcode-wrap{width:calc(var(--label-w) - 14.7mm)!important}}
@media print{.density-compact.template-commercial,.density-compact.template-showcase{padding:1.0mm 1.1mm .9mm!important;grid-template-rows:2.7mm 6.25mm 5.15mm 10.95mm!important;gap:.24mm!important;align-content:center!important}.density-compact.template-commercial .label-title,.density-compact.template-showcase .label-title{height:6.25mm!important}.density-compact.template-commercial .name,.density-compact.template-showcase .name{height:4.9mm!important;max-height:4.9mm!important;font-size:7.0px!important}.density-compact.template-commercial .name.name-medium,.density-compact.template-showcase .name.name-medium{font-size:6.65px!important}.density-compact.template-commercial .name.name-long,.density-compact.template-showcase .name.name-long{font-size:6.15px!important}.density-compact.template-commercial .price,.density-compact.template-showcase .price{min-height:5.15mm!important}.density-compact.template-commercial .price strong{font-size:9.45px!important}.density-compact.template-showcase .price strong{font-size:10.1px!important}.density-compact.template-commercial .codes,.density-compact.template-showcase .codes{height:10.95mm!important}.density-compact.template-commercial .qr-wrap,.density-compact.template-showcase .qr-wrap{grid-template-columns:8.9mm!important;grid-template-rows:8.9mm 1.15mm!important}.density-compact.template-commercial .qr,.density-compact.template-showcase .qr{width:8.9mm!important;height:8.9mm!important}.density-compact.template-commercial .qr-wrap small,.density-compact.template-showcase .qr-wrap small{display:block!important;font-size:4.1px!important}.density-compact.template-commercial .barcode-wrap,.density-compact.template-showcase .barcode-wrap{grid-template-rows:7mm auto!important}.density-compact.template-commercial .barcode-svg,.density-compact.template-showcase .barcode-svg{height:7mm!important}.density-compact.template-commercial .codes-both .barcode-wrap,.density-compact.template-showcase .codes-both .barcode-wrap{width:auto!important;flex:1!important}}

@media print{.price strong{white-space:nowrap!important}.composition-two-column .price strong{white-space:nowrap!important}.composition-two-column.price-max .price strong{font-size:13.6px!important}.label.orientation-vertical.vertical-compact{gap:.38mm!important;align-content:start!important}.label.orientation-vertical.vertical-compact .codes{align-items:center!important}.label.orientation-vertical.price-max .price strong{font-size:26px!important}.density-vertical-medium.orientation-vertical{grid-template-rows:3.95mm 11.9mm 14.3mm 25.4mm!important;gap:.38mm!important;padding:1.95mm 2.05mm 1.6mm!important}.density-vertical-medium.orientation-vertical.price-max .price strong{font-size:25.8px!important}.density-vertical-medium.orientation-vertical .price strong{font-size:22.8px!important}.density-vertical-medium.orientation-vertical.mode-barcode .barcode-svg{height:17.4mm!important}.density-vertical-compact.orientation-vertical{grid-template-rows:3.65mm 10.4mm 13.4mm 24.2mm!important;gap:.34mm!important;padding:1.7mm 1.7mm 1.45mm!important}.density-vertical-compact.orientation-vertical .price strong{font-size:20.2px!important}.density-vertical-compact.orientation-vertical.price-max .price strong{font-size:22.8px!important}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-svg{height:16.8mm!important}.density-vertical-large.orientation-vertical{grid-template-rows:5.1mm 18.6mm 18.7mm 32.4mm!important;gap:.58mm!important;padding:2.75mm 2.95mm 2.35mm!important}.density-vertical-large.orientation-vertical.price-max .price strong{font-size:33px!important}.density-vertical-large.orientation-vertical .price strong{font-size:29.2px!important}.label.orientation-vertical.composition-two-column .price{display:flex!important;flex-direction:row!important;align-items:baseline!important}.label.orientation-vertical.composition-two-column .price strong{white-space:nowrap!important;font-size:22px!important}.qr-large .qr-wrap{grid-template-columns:13.8mm!important;grid-template-rows:13.8mm auto!important}.qr-large .qr{width:13.8mm!important;height:13.8mm!important}.qr-max .qr-wrap{grid-template-columns:17mm!important;grid-template-rows:17mm auto!important}.qr-max .qr{width:17mm!important;height:17mm!important}}
</style></head><body>${pagesMarkup || '<section class="label-page"><p>No hay etiquetas seleccionadas.</p></section>'}
<script>
  const waitForImages = () => Promise.all(Array.from(document.images).map((image) => image.complete ? Promise.resolve() : new Promise((resolve) => { image.addEventListener('load', resolve, { once:true }); image.addEventListener('error', resolve, { once:true }); setTimeout(resolve, 2200); })));
  window.addEventListener('load', () => waitForImages().then(() => setTimeout(() => window.print(), 250)));
</script></body></html>`;
};

const openLabelsPrintWindow = (options) => {
  const popup = window.open('', '_blank', 'width=1100,height=820');
  if (!popup) {
    alert('El navegador bloqueó la ventana de impresión. Permite ventanas emergentes para este sitio y vuelve a intentar.');
    return false;
  }
  popup.document.open();
  popup.document.write(buildLabelsPrintHTML(options));
  popup.document.close();
  return true;
};
const productProfit = (product) => asNum(product?.price) - asNum(product?.cost);
const productMarkupPercent = (product) => asNum(product?.cost) > 0 ? (productProfit(product) / asNum(product.cost)) * 100 : 0;
const productMarginPercent = (product) => asNum(product?.price) > 0 ? (productProfit(product) / asNum(product.price)) * 100 : 0;
const suggestedPrice = (cost, marginTarget = 50) => asNum(cost) > 0 ? Math.round((asNum(cost) * (1 + asNum(marginTarget) / 100)) * 10) / 10 : 0;
const priceBadgeClass = (status) => `price-badge price-${normalizeText(status || 'Pendiente').replace(/[^a-z0-9]+/g, '-')}`;
const cleanFileName = (name = 'producto') => String(name).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/-+/g, '-').toLowerCase();

const normalizeText = (value = '') => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const normalizeHeader = (value = '') => normalizeText(value).replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
const categoryPrefix = (categoryName = '') => {
  const n = normalizeText(categoryName);
  if (n.includes('ropa hombre')) return 'RH';
  if (n.includes('ropa mujer')) return 'RM';
  if (n.includes('calzado')) return 'CAL';
  if (n.includes('accesorios')) return 'ACC';
  if (n.includes('belleza') || n.includes('cuidado')) return 'BEL';
  if (n.includes('hogar')) return 'HOG';
  if (n.includes('fiesta') || n.includes('pinateria')) return 'FIE';
  if (n.includes('bazar')) return 'BAZ';
  return 'GEN';
};
const canonicalProductField = (key) => ({
  codigo: 'code', code: 'code', cod: 'code', codigo_interno: 'code', sku: 'code',
  barcode: 'barcode', codigo_barras: 'barcode', codigo_de_barras: 'barcode', barra: 'barcode', ean: 'barcode', upc: 'barcode',
  nombre: 'name', producto: 'name', name: 'name', descripcion_producto: 'name',
  categoria: 'category', category: 'category',
  subcategoria: 'subcategory', sub_category: 'subcategory', subcategory: 'subcategory',
  marca: 'brand', brand: 'brand',
  talla: 'size', size: 'size', medida: 'size',
  color: 'color',
  descripcion: 'description', description: 'description', detalles: 'description',
  costo: 'cost', cost: 'cost', costo_compra: 'cost', compra: 'cost',
  precio: 'price', price: 'price', precio_venta: 'price', venta: 'price',
  estado_precio: 'price_status', price_status: 'price_status', precio_estado: 'price_status', validar_precio: 'price_status',
  margen_objetivo: 'margin_target', margin_target: 'margin_target', margen: 'margin_target',
  precio_minimo: 'min_price', min_price: 'min_price', precio_min: 'min_price',
  notas_precio: 'price_notes', price_notes: 'price_notes', observacion_precio: 'price_notes',
  stock: 'stock', cantidad: 'stock', unidades: 'stock',
  stock_min: 'stock_min', stock_minimo: 'stock_min', minimo: 'stock_min',
  imagen_url: 'image_url', image_url: 'image_url', url_imagen: 'image_url', imagen: 'image_url',
  activo: 'active', active: 'active', estado: 'active',
}[key] || key);
const parseMoneyLike = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/S\/?/gi, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
  return Number(cleaned || 0);
};
const boolActive = (value) => {
  if (value === false) return false;
  const n = normalizeText(value);
  return !['no', 'false', 'inactivo', '0', 'desactivado'].includes(n);
};
const ROLE_LABELS = {
  dueno: 'Dueño',
  admin: 'Administrador',
  cajero: 'Cajero',
  almacen: 'Almacén',
  lectura: 'Solo lectura',
};
const ROLE_HOME = {
  dueno: 'panel',
  admin: 'panel',
  cajero: 'ventas',
  almacen: 'productos',
  lectura: 'reportes',
};
const MODULE_PERMISSIONS = {
  panel: ['dueno', 'admin', 'lectura'],
  ia: ['dueno', 'admin'],
  ventas: ['dueno', 'admin', 'cajero'],
  comprobantes: ['dueno', 'admin', 'cajero'],
  creditos: ['dueno', 'admin', 'cajero'],
  caja: ['dueno', 'admin', 'cajero'],
  reportes: ['dueno', 'admin', 'lectura'],
  productos: ['dueno', 'admin', 'almacen'],
  catalogo: ['dueno', 'admin', 'almacen'],
  pedidos: ['dueno', 'admin', 'cajero'],
  precios: ['dueno', 'admin'],
  categorias: ['dueno', 'admin', 'almacen'],
  etiquetas: ['dueno', 'admin', 'almacen'],
  inventario: ['dueno', 'admin', 'almacen', 'lectura'],
  ingreso: ['dueno', 'admin', 'almacen'],
  clientes: ['dueno', 'admin', 'cajero'],
  usuarios: ['dueno', 'admin'],
  tienda: ['dueno', 'admin'],
  herramientas: ['dueno'],
};
const canAccess = (role, moduleKey) => MODULE_PERMISSIONS[moduleKey]?.includes(role || 'cajero');
const firstAllowedModule = (role) => ROLE_HOME[role] || 'ventas';


function useAuth() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!hasSupabaseConfig) {
      setLoading(false);
      return;
    }
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  return { session, loading };
}


function useUserProfile(session) {
  const [profile, setProfile] = useState(null);
  const [store, setStore] = useState(null);
  const [loading, setLoading] = useState(true);

  async function loadProfile() {
    if (!hasSupabaseConfig || !session?.user?.id) {
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data: profileData, error: profileError } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', session.user.id)
      .single();

    const fallbackProfile = {
      id: session.user.id,
      store_id: DEFAULT_STORE_ID,
      email: session.user.email,
      full_name: session.user.email,
      role: 'cajero',
      status: 'Activo',
    };
    const nextProfile = profileError ? fallbackProfile : (profileData || fallbackProfile);
    setProfile(nextProfile);

    const { data: storeData } = await supabase
      .from('stores')
      .select('*')
      .eq('id', nextProfile.store_id || DEFAULT_STORE_ID)
      .single();
    setStore(storeData || { id: DEFAULT_STORE_ID, name: 'Clomar Store Pro' });
    setLoading(false);
  }

  useEffect(() => { loadProfile(); }, [session?.user?.id]);
  return { profile, store, loading, reload: loadProfile };
}

function roleMeta(profile) {
  return ROLE_LABELS[profile?.role] || 'Usuario';
}

function AccessDenied({ profile, setCurrent }) {
  return (
    <div className="page">
      <div className="hero compact-hero">
        <h1>🔒 Acceso restringido</h1>
        <p>Tu rol actual es {roleMeta(profile)}. No tienes permiso para este módulo.</p>
      </div>
      <section className="card compact-card">
        <p className="muted">Regresa a un módulo permitido o solicita al dueño cambiar tus permisos.</p>
        <button className="primary-btn" onClick={() => setCurrent(firstAllowedModule(profile?.role))}>Ir a mi panel</button>
      </section>
    </div>
  );
}

function InactiveUser({ profile }) {
  return (
    <main className="login-page">
      <section className="login-card">
        <div className="brand-logo"><img src={APP_ICON} alt="Clomar Store" /></div>
        <h1>Usuario inactivo</h1>
        <p>La cuenta {profile?.email || ''} está inactiva. Contacta al dueño o administrador.</p>
        <button className="primary-btn" onClick={() => supabase?.auth.signOut()}>Cerrar sesión</button>
      </section>
    </main>
  );
}

function Login() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  async function signIn(e) {
    e.preventDefault();
    setError('');
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { error: signError } = await supabase.auth.signInWithPassword({ email, password });
    if (signError) setError('Correo o contraseña incorrectos.');
    setLoading(false);
  }

  return (
    <main className="login-page">
      <section className="login-card">
        <div className="brand-logo"><img src={APP_ICON} alt="Clomar Store" /></div>
        <h1>Clomar Store</h1>
        <p>POS rápido para ventas, inventario, caja y créditos.</p>
        {!hasSupabaseConfig && (
          <div className="warning-box">
            Falta configurar Supabase. Copia <strong>.env.example</strong> como <strong>.env</strong> y coloca tus claves.
          </div>
        )}
        <form onSubmit={signIn}>
          <label>Correo del usuario</label>
          <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="usuario@tienda.com" autoComplete="email" />
          <label>Contraseña</label>
          <input value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Tu contraseña" type="password" autoComplete="current-password" />
          {error && <div className="error-box">{error}</div>}
          <button disabled={loading || !hasSupabaseConfig} className="primary-btn">{loading ? 'Ingresando...' : 'Entrar'}</button>
        </form>
      </section>
    </main>
  );
}

function Sidebar({ current, setCurrent, open, setOpen, session, profile, store }) {
  const role = profile?.role || 'cajero';
  const mode = role === 'cajero' ? 'Modo vendedor' : role === 'almacen' ? 'Modo inventario' : role === 'lectura' ? 'Modo consulta' : 'Modo propietario';
  const ownerSections = [
    { title: 'Operación', items: [['ventas', '🧾', 'Nueva venta'], ['caja', '💰', 'Caja por turno'], ['pedidos', '💬', 'WhatsApp CRM'], ['comprobantes', '📄', 'Comprobantes'], ['creditos', '💳', 'Créditos']] },
    { title: 'Catálogo e inventario', items: [['productos', '📦', 'Productos'], ['inventario', '📘', 'Inventario'], ['ingreso', '📥', 'Ingreso mercadería'], ['precios', '🏷️', 'Precios'], ['etiquetas', '🔖', 'Etiquetas'], ['catalogo', '🛍️', 'Catálogo público'], ['categorias', '🗂️', 'Categorías']] },
    { title: 'Control', items: [['panel', '📊', 'Panel del dueño'], ['reportes', '📈', 'Reportes'], ['ia', '✦', 'Asistente IA'], ['clientes', '👥', 'Clientes']] },
    { title: 'Administración', items: [['usuarios', '🧑‍💼', 'Usuarios'], ['tienda', '🏪', 'Tienda'], ['herramientas', '🛠️', 'Herramientas']] },
  ];
  const operatorSections = role === 'almacen'
    ? [
      { title: 'Mi operación', items: [['productos', '📦', 'Productos'], ['inventario', '📘', 'Inventario'], ['ingreso', '📥', 'Ingreso mercadería'], ['etiquetas', '🔖', 'Etiquetas']] },
      { title: 'Catálogo', items: [['catalogo', '🛍️', 'Catálogo público'], ['categorias', '🗂️', 'Categorías']] },
    ]
    : role === 'lectura'
      ? [{ title: 'Consulta', items: [['panel', '📊', 'Panel comercial'], ['reportes', '📈', 'Reportes'], ['inventario', '📘', 'Inventario']] }]
      : [
        { title: 'Mi turno', items: [['ventas', '🧾', 'Nueva venta'], ['caja', '💰', 'Caja por turno'], ['pedidos', '💬', 'WhatsApp CRM'], ['comprobantes', '📄', 'Comprobantes'], ['creditos', '💳', 'Créditos']] },
        { title: 'Clientes', items: [['clientes', '👥', 'Clientes']] },
      ];
  const sections = (['dueno', 'admin'].includes(role) ? ownerSections : operatorSections)
    .map(section => ({ ...section, items: section.items.filter(([key]) => canAccess(role, key)) }))
    .filter(section => section.items.length);
  return (
    <aside className={`sidebar premium-sidebar ${open ? 'open' : ''}`}>
      <div className="sidebar-head premium-sidebar-head">
        <div className="mini-logo"><img src={logoSrc(store)} alt="Logo tienda" /></div>
        <div className="sidebar-store-copy">
          <strong>{store?.name || 'Clomar Store'}</strong>
          <small>{profile?.full_name || session?.user?.email || 'Usuario'}</small>
        </div>
        <button className="ghost mobile-only" type="button" onClick={() => setOpen(false)} aria-label="Cerrar menú"><X size={18}/></button>
      </div>
      <div className="sidebar-mode-card"><span className="status-dot" /><div><strong>{mode}</strong><small>{roleMeta(profile)} · Accesos según rol</small></div></div>
      {sections.map((section) => (
        <div key={section.title} className="menu-section premium-menu-section">
          <span>{section.title}</span>
          {section.items.map(([key, icon, label]) => (
            <button key={key} className={current === key ? 'active' : ''} onClick={() => { setCurrent(key); setOpen(false); }}>
              <span className="nav-icon premium-nav-icon" aria-hidden="true">{icon}</span><span>{label}</span><span className="nav-chevron" aria-hidden="true">›</span>
            </button>
          ))}
        </div>
      ))}
      <div className="sidebar-bottom-help"><span>Atajo útil</span><strong>Use el buscador para vender más rápido</strong></div>
      <button className="logout" type="button" onClick={() => supabase?.auth.signOut()}><LogOut size={16}/> Cerrar sesión</button>
    </aside>
  );
}
function Header({ setOpen, current, profile, store, setCurrent }) {
  const titleMap = {
    panel: 'Panel del dueño', ia: 'Asistente IA', ventas: 'Nueva venta', comprobantes: 'Comprobantes', creditos: 'Créditos', caja: 'Caja por turno', reportes: 'Reportes', productos: 'Productos', catalogo: 'Catálogo público', pedidos: 'WhatsApp CRM', precios: 'Control de precios', categorias: 'Categorías', etiquetas: 'Etiquetas', inventario: 'Inventario', ingreso: 'Compras y proveedores', clientes: 'Clientes', usuarios: 'Usuarios y roles', tienda: 'Configuración de tienda', herramientas: 'Herramientas'
  };
  const role = profile?.role || 'cajero';
  const mode = role === 'cajero' ? 'Modo vendedor' : role === 'almacen' ? 'Modo inventario' : role === 'lectura' ? 'Modo consulta' : 'Modo propietario';
  return (
    <header className="app-header app-header-pro premium-app-header">
      <button className="ghost mobile-only menu-toggle-pro" type="button" onClick={() => setOpen(true)} aria-label="Abrir menú"><Menu/></button>
      <div className="header-brand-mobile"><img src={logoSrc(store)} alt="Logo tienda" /></div>
      <div className="header-title-block"><h2>{titleMap[current]}</h2><p>{store?.name || 'Clomar Store'} · {mode}</p></div>
      <div className="header-actions-premium">
        {canAccess(role, 'ventas') && current !== 'ventas' && <button type="button" className="header-quick-sale" onClick={() => setCurrent?.('ventas')}>+ Nueva venta</button>}
        <div className="header-status-chip" title={APP_VERSION}><span className="status-dot" /><small>En línea</small></div>
      </div>
    </header>
  );
}
function useProducts(profile) {
  const [products, setProducts] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadProducts() {
    if (!hasSupabaseConfig) {
      setProducts(demoProducts);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('products')
      .select('id,code,name,category,subcategory,category_id,subcategory_id,price,cost,stock,stock_min,status,store_id,image_url,image_path,brand,size,color,description,barcode,active,price_status,margin_target,min_price,price_notes,price_updated_at,price_updated_by,public_visible,catalog_status,catalog_featured,catalog_description,catalog_position,catalog_updated_at')
      .eq('status', 'Activo')
      .eq('active', true)
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('name');
    if (!error) {
      setProducts(data || []);
    } else {
      console.error('No se pudieron cargar productos:', error);
      // Conserva la lista anterior para no vaciar el POS ante un fallo transitorio.
      // Esta versión usa únicamente columnas ya existentes en la base original.
    }
    setLoading(false);
  }

  useEffect(() => { loadProducts(); }, [profile?.store_id]);
  return { products, loading, reload: loadProducts };
}

function useCustomers(profile) {
  const [customers, setCustomers] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadCustomers() {
    if (!hasSupabaseConfig) {
      setCustomers([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('customers')
      .select('*')
      .eq('status', 'Activo')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('name');
    if (!error) setCustomers(data || []);
    setLoading(false);
  }

  useEffect(() => { loadCustomers(); }, [profile?.store_id]);
  return { customers, loading, reload: loadCustomers };
}


function useCategories(profile) {
  const [categories, setCategories] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadCategories() {
    if (!hasSupabaseConfig) {
      setCategories([
        { id: 'cat-demo-1', name: 'Ropa hombre', parent_id: null, sort_order: 1, active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
        { id: 'cat-demo-2', name: 'Ropa mujer', parent_id: null, sort_order: 2, active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
        { id: 'cat-demo-3', name: 'Calzado', parent_id: null, sort_order: 3, active: true, price_status: 'Pendiente', margin_target: 50, min_price: 0, price_notes: '' },
      ]);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('product_categories')
      .select('id,store_id,name,description,parent_id,sort_order,active')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .eq('active', true)
      .order('sort_order', { ascending: true })
      .order('name', { ascending: true });
    if (!error) setCategories(data || []);
    setLoading(false);
  }

  useEffect(() => { loadCategories(); }, [profile?.store_id]);
  return {
    categories: categories.filter(c => !c.parent_id),
    subcategories: categories.filter(c => c.parent_id),
    allCategories: categories,
    loading,
    reload: loadCategories,
  };
}

function useSales(profile) {
  const [sales, setSales] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState('');
  async function loadSales(limit = 50) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    setLoadError('');
    const { data, error } = await supabase
      .from('sales')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (error) {
      console.error('No se pudieron cargar ventas:', error);
      setLoadError(error.message || 'No se pudieron cargar las ventas.');
    } else {
      setSales(data || []);
    }
    setLoading(false);
  }
  useEffect(() => { loadSales(); }, [profile?.store_id]);
  return { sales, loading, loadError, reload: loadSales };
}

function useCashMovements(profile) {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState('');
  async function loadMovements(limit = 120) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    setLoadError('');
    const { data, error } = await supabase
      .from('cash_movements')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (error) {
      console.error('No se pudieron cargar movimientos de caja:', error);
      setLoadError(error.message || 'No se pudieron cargar los movimientos de caja.');
    } else {
      setMovements(data || []);
    }
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, [profile?.store_id]);
  return { movements, loading, loadError, reload: loadMovements };
}


function useCashSession(profile) {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState('');
  async function reload() {
    if (!hasSupabaseConfig || !profile?.store_id) { setSession(null); return; }
    setLoading(true);
    setLoadError('');
    const { data, error } = await supabase
      .from('cash_sessions')
      .select('id,store_id,opened_by,closed_by,status,opening_amount,expected_amount,counted_amount,difference_amount,opening_note,closing_note,opened_at,closed_at,created_at')
      .eq('store_id', profile.store_id || DEFAULT_STORE_ID)
      .eq('status', 'Abierta')
      .order('opened_at', { ascending: false })
      .limit(1)
      .maybeSingle();
    if (error) {
      console.error('No se pudo leer la caja activa:', error);
      setLoadError(error.message || 'No se pudo leer la caja activa.');
      setSession(null);
    } else {
      setSession(data || null);
    }
    setLoading(false);
  }
  useEffect(() => { reload(); }, [profile?.store_id]);
  return { session, loading, loadError, reload };
}

function useStockMovements(profile) {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadMovements(limit = 80) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('stock_movements')
      .select('*, products(name, code)')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setMovements(data || []);
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, [profile?.store_id]);
  return { movements, loading, reload: loadMovements };
}

function Kpi({ label, value, helper }) {
  return <div className="kpi"><span>{label}</span><strong>{value}</strong><small>{helper}</small></div>;
}

function fmtWhole(value) {
  return Number(value || 0).toLocaleString('es-PE', { maximumFractionDigits: 0 });
}

function managementSeverityClass(severity = '') {
  const value = String(severity || '').toLowerCase();
  if (value === 'critical' || value === 'alta') return 'critical';
  if (value === 'warning' || value === 'media') return 'warning';
  return 'info';
}

function managementAlertIcon(kind = '') {
  const key = String(kind || '').toLowerCase();
  if (key.includes('stock')) return '📦';
  if (key.includes('credit') || key.includes('cobran')) return '💳';
  if (key.includes('cash') || key.includes('caja')) return '💰';
  if (key.includes('margin') || key.includes('margen')) return '📉';
  if (key.includes('catalog')) return '🛍️';
  if (key.includes('rotation') || key.includes('rotacion')) return '🕒';
  return 'ℹ️';
}

function useManagementDashboard(profile, days) {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const storeId = profile?.store_id || DEFAULT_STORE_ID;

  async function reload() {
    if (!hasSupabaseConfig || !profile?.id) return;
    setLoading(true);
    setError('');
    const { data: result, error: rpcError } = await supabase.rpc('clomar_management_dashboard_v32', {
      p_store_id: storeId,
      p_days: Number(days || 30),
    });
    if (rpcError) {
      setError(rpcError.message || 'No se pudo cargar el control gerencial.');
      setData(null);
    } else {
      setData(result || null);
    }
    setLoading(false);
  }

  useEffect(() => { reload(); }, [profile?.id, storeId, days]);
  return { data, loading, error, reload };
}

function Panel({ products, profile, setCurrent }) {
  const [days, setDays] = useState(30);
  const { data, loading, error, reload } = useManagementDashboard(profile, days);
  const role = profile?.role || 'cajero';

  if (!['dueno', 'admin'].includes(role)) {
    return (
      <div className="page compact-page dashboard-owner-page">
        <div className="hero compact-hero owner-hero"><div><span className="eyebrow">Vista de lectura</span><h1>📊 Panel comercial</h1><p>El control gerencial detallado está reservado para dueño y administrador.</p></div></div>
        <section className="card compact-card"><h3>Acceso disponible</h3><p className="muted">Puede revisar los resultados desde Reportes e Inventario. Las alertas de rentabilidad, cobranza y control de caja requieren autorización de administrador.</p></section>
      </div>
    );
  }

  const metrics = data?.metrics || {};
  const alerts = Array.isArray(data?.alerts) ? data.alerts : [];
  const topProducts = Array.isArray(data?.top_products) ? data.top_products : [];
  const slowProducts = Array.isArray(data?.slow_products) ? data.slow_products : [];
  const stockCritical = Array.isArray(data?.stock_critical) ? data.stock_critical : [];
  const sellers = Array.isArray(data?.sellers) ? data.sellers : [];
  const paymentMethods = Array.isArray(data?.payment_methods) ? data.payment_methods : [];
  const crossSell = Array.isArray(data?.cross_sell) ? data.cross_sell : [];
  const maxPayment = Math.max(1, ...paymentMethods.map(row => asNum(row.amount)));
  const maxSeller = Math.max(1, ...sellers.map(row => asNum(row.amount)));
  const salesDelta = asNum(metrics.sales_delta_percent);

  return (
    <div className="page compact-page management-dashboard-page">
      <div className="hero compact-hero owner-hero management-hero">
        <div>
          <span className="eyebrow">Control del propietario</span>
          <h1>📊 Centro de control gerencial</h1>
          <p>Ventas, rentabilidad, caja, crédito, rotación y oportunidades comerciales con datos reales de Clomar Store.</p>
        </div>
        <div className="management-hero-actions">
          <label>Periodo<select value={days} onChange={e => setDays(Number(e.target.value))}><option value="7">Últimos 7 días</option><option value="30">Últimos 30 días</option><option value="60">Últimos 60 días</option><option value="90">Últimos 90 días</option></select></label>
          <button type="button" className="secondary-btn" onClick={reload}>{loading ? 'Actualizando...' : 'Actualizar'}</button>
        </div>
      </div>

      {error && <section className="data-error"><strong>No se pudo cargar V03.2:</strong> {error}. Verifique que ejecutó el SQL de esta entrega y recargue la página.</section>}
      {loading && !data && <div className="loader">Preparando indicadores gerenciales...</div>}

      {data && <>
        <div className="management-period-strip"><span>Periodo analizado</span><strong>{data.period_start || '—'} al {data.period_end || '—'}</strong><small>Actualizado: {data.generated_at ? fmtDate(data.generated_at) : 'ahora'}</small></div>
        <div className="management-kpi-grid">
          <Kpi label="Ventas del periodo" value={money(metrics.sales_total)} helper={`${fmtWhole(metrics.sales_count)} comprobantes · ${salesDelta >= 0 ? '+' : ''}${salesDelta.toFixed(1)}% vs periodo anterior`} />
          <Kpi label="Utilidad bruta" value={money(metrics.profit_total)} helper={`Margen ${asNum(metrics.margin_percent).toFixed(1)}%`} />
          <Kpi label="Ticket promedio" value={money(metrics.ticket_average)} helper={`${fmtWhole(metrics.units_sold)} unidades vendidas`} />
          <Kpi label="Crédito pendiente" value={money(metrics.credit_pending)} helper={`${fmtWhole(metrics.credit_overdue_count)} crédito(s) vencido(s)`} />
          <Kpi label="Stock crítico" value={fmtWhole(metrics.stock_critical_count)} helper={`${fmtWhole(metrics.slow_stock_count)} con baja rotación`} />
          <Kpi label="Caja última diferencia" value={money(metrics.last_cash_difference)} helper={metrics.last_cash_closed_at ? `Cierre ${fmtDate(metrics.last_cash_closed_at)}` : 'Sin caja cerrada'} />
        </div>
        <section className="owner-action-board">
          <div className="owner-action-board-head"><div><span className="eyebrow">Acciones sugeridas</span><h3>Resuelva lo importante primero</h3><p>Accesos directos para convertir alertas en tareas concretas.</p></div><span className="owner-action-live">Datos del periodo actual</span></div>
          <div className="owner-action-grid">
            <button type="button" className="owner-action-card critical" onClick={() => setCurrent?.('inventario')}><span>Stock</span><strong>{fmtWhole(metrics.stock_critical_count)} por reponer</strong><small>Revise mínimos y programe compra</small><b>Ver inventario →</b></button>
            <button type="button" className="owner-action-card warning" onClick={() => setCurrent?.('creditos')}><span>Cobranza</span><strong>{money(metrics.credit_pending)} pendiente</strong><small>{fmtWhole(metrics.credit_overdue_count)} crédito(s) vencido(s)</small><b>Gestionar cobros →</b></button>
            <button type="button" className="owner-action-card neutral" onClick={() => setCurrent?.('caja')}><span>Caja</span><strong>{money(metrics.last_cash_difference)} diferencia</strong><small>{metrics.last_cash_closed_at ? 'Revise el último cierre registrado' : 'Aún no hay cierre registrado'}</small><b>Ir a caja →</b></button>
            <button type="button" className="owner-action-card highlight" onClick={() => setCurrent?.('catalogo')}><span>Catálogo</span><strong>{fmtWhole(metrics.slow_stock_count)} con baja rotación</strong><small>Prepare oferta, foto o publicación destacada</small><b>Mejorar catálogo →</b></button>
          </div>
        </section>

        <section className="management-alerts card compact-card">
          <div className="section-head-inline"><div><span className="eyebrow">Prioridades del día</span><h3>Qué debe revisar primero</h3></div><span className="result-pill">{alerts.length} alerta(s)</span></div>
          <div className="management-alert-grid">
            {alerts.length ? alerts.map((alert, idx) => <article className={`management-alert ${managementSeverityClass(alert.severity)}`} key={`${alert.kind || 'alert'}-${idx}`}><span>{managementAlertIcon(alert.kind)}</span><div><strong>{alert.title || 'Alerta operativa'}</strong><p>{alert.message || 'Revise este indicador.'}</p></div></article>) : <article className="management-alert success"><span>✅</span><div><strong>Operación bajo control</strong><p>No hay alertas críticas para el periodo seleccionado.</p></div></article>}
          </div>
        </section>

        <div className="management-grid-primary">
          <section className="card compact-card"><div className="section-head-inline"><div><span className="eyebrow">Rentabilidad</span><h3>Productos más rentables</h3></div><span className="result-pill">Top {topProducts.length}</span></div><div className="management-table"><div className="management-table-head"><span>Producto</span><span>Unid.</span><span>Vendido</span><span>Utilidad</span></div>{topProducts.map((row, idx) => <div className="management-table-row" key={`${row.product_id || row.name}-${idx}`}><span><strong>{row.name || 'Producto'}</strong><small>{row.code || 'Sin código'} · margen {asNum(row.margin_percent).toFixed(1)}%</small></span><b>{fmtWhole(row.qty)}</b><b>{money(row.amount)}</b><b className={asNum(row.profit) < 0 ? 'danger-text' : ''}>{money(row.profit)}</b></div>)}{!topProducts.length && <p className="muted">Aún no hay ventas en el periodo.</p>}</div></section>
          <section className="card compact-card"><div className="section-head-inline"><div><span className="eyebrow">Reposición</span><h3>Stock crítico</h3></div><span className="result-pill danger">{stockCritical.length} producto(s)</span></div><div className="management-list">{stockCritical.map((row, idx) => <div className="list-row" key={`${row.product_id || row.name}-${idx}`}><span><strong>{row.name || 'Producto'}</strong><small>{row.code || 'Sin código'} · mínimo {fmtWhole(row.stock_min)}</small></span><b className="danger-text">{fmtWhole(row.stock)}</b></div>)}{!stockCritical.length && <p className="muted">No hay productos con stock crítico.</p>}</div></section>
          <section className="card compact-card"><div className="section-head-inline"><div><span className="eyebrow">Rotación</span><h3>Mercadería sin movimiento</h3></div><span className="result-pill">60 días</span></div><div className="management-list">{slowProducts.map((row, idx) => <div className="list-row" key={`${row.product_id || row.name}-${idx}`}><span><strong>{row.name || 'Producto'}</strong><small>{row.code || 'Sin código'} · stock actual {fmtWhole(row.stock)}</small></span><b>{money(row.stock_value)}</b></div>)}{!slowProducts.length && <p className="muted">No hay mercadería sin movimiento en los últimos 60 días.</p>}</div></section>
        </div>

        <div className="management-grid-secondary">
          <section className="card compact-card"><span className="eyebrow">Cobros</span><h3>Créditos por cobrar</h3><div className="credit-summary-rows"><div><span>Saldo total</span><strong>{money(metrics.credit_pending)}</strong></div><div><span>Vencido</span><strong className={asNum(metrics.credit_overdue) > 0 ? 'danger-text' : ''}>{money(metrics.credit_overdue)}</strong></div><div><span>Pedidos web pendientes</span><strong>{fmtWhole(metrics.pending_catalog_orders)}</strong></div></div><p className="muted">Priorice los créditos vencidos antes de otorgar nuevas excepciones.</p></section>
          <section className="card compact-card"><span className="eyebrow">Ingresos</span><h3>Por método de pago</h3><div className="bar-list">{paymentMethods.map((row, idx) => <div className="bar-row" key={`${row.method}-${idx}`}><div><strong>{row.method || 'Sin método'}</strong><small>{fmtWhole(row.operations)} operación(es)</small></div><div className="bar-track"><span style={{ width: `${Math.max(5, (asNum(row.amount) / maxPayment) * 100)}%` }} /></div><b>{money(row.amount)}</b></div>)}{!paymentMethods.length && <p className="muted">No hay ingresos registrados en el periodo.</p>}</div></section>
          <section className="card compact-card"><span className="eyebrow">Equipo</span><h3>Rendimiento por vendedor</h3><div className="bar-list">{sellers.map((row, idx) => <div className="bar-row" key={`${row.user_id || row.seller}-${idx}`}><div><strong>{row.seller || 'Sin vendedor'}</strong><small>{fmtWhole(row.sales_count)} venta(s) · ticket {money(row.ticket_average)}</small></div><div className="bar-track"><span style={{ width: `${Math.max(5, (asNum(row.amount) / maxSeller) * 100)}%` }} /></div><b>{money(row.amount)}</b></div>)}{!sellers.length && <p className="muted">No hay ventas por vendedor en el periodo.</p>}</div></section>
          <section className="card compact-card"><span className="eyebrow">Venta cruzada</span><h3>Productos que se compran juntos</h3><div className="management-list">{crossSell.map((row, idx) => <div className="list-row" key={`${row.product_a}-${row.product_b}-${idx}`}><span><strong>{row.product_a || 'Producto'} + {row.product_b || 'Producto'}</strong><small>Comprados juntos en {fmtWhole(row.times_together)} venta(s)</small></span><b>{fmtWhole(row.times_together)}</b></div>)}{!crossSell.length && <p className="muted">Se necesitan más ventas con varios productos para detectar combinaciones.</p>}</div></section>
        </div>
      </>}
    </div>
  );
}

const ASSISTANT_QUICK_QUESTIONS = [
  { intent: 'reponer', label: '¿Qué debo reponer?', text: '¿Qué productos debo reponer esta semana?' },
  { intent: 'rotacion', label: 'Productos lentos', text: '¿Qué productos no se venden hace tiempo?' },
  { intent: 'cobranza', label: 'Cobranza', text: '¿Cuánto tengo pendiente de cobrar?' },
  { intent: 'rentabilidad', label: 'Rentabilidad', text: '¿Qué productos me dejan más utilidad?' },
  { intent: 'vendedores', label: 'Vendedores', text: '¿Cómo va el rendimiento de los vendedores?' },
  { intent: 'pagos', label: 'Pagos', text: '¿Qué método de pago se usa más?' },
  { intent: 'resumen', label: 'Resumen', text: 'Dame un resumen del negocio.' },
];

function inferAssistantIntent(text = '') {
  const value = normalizeText(text);
  if (/(reponer|reposicion|comprar|stock|faltante)/.test(value)) return 'reponer';
  if (/(lento|rotacion|no se vende|sin vender|estancado)/.test(value)) return 'rotacion';
  if (/(cobrar|cobranza|deuda|credito|vencid)/.test(value)) return 'cobranza';
  if (/(rentab|utilidad|ganancia|margen)/.test(value)) return 'rentabilidad';
  if (/(vendedor|equipo|quien vende)/.test(value)) return 'vendedores';
  if (/(yape|plin|tarjeta|efectivo|pago|metodo)/.test(value)) return 'pagos';
  return 'resumen';
}

function copyTextToClipboard(value) {
  const text = String(value || '').trim();
  if (!text) return false;
  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(text).catch(() => {});
    return true;
  }
  const area = document.createElement('textarea');
  area.value = text;
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.select();
  try { document.execCommand('copy'); } catch (err) {}
  document.body.removeChild(area);
  return true;
}


function cleanAssistantLine(value = '') {
  return String(value || '')
    .replace(/\r/g, '')
    .replace(/\*\*/g, '')
    .replace(/`/g, '')
    .replace(/^#{1,6}\s*/, '')
    .replace(/^>\s*/, '')
    .trim();
}

function parseAssistantAnswer(value = '') {
  const lines = String(value || '').replace(/\r/g, '').split('\n');
  const intro = [];
  const sections = [];
  let current = null;
  const isHeading = (line) => {
    const plain = cleanAssistantLine(line);
    if (!plain || plain.length > 92) return false;
    return /^(?:\d+[.)]\s*)?(?:diagn[oó]stico|prioridades?|acciones?|datos\s+faltantes|resumen|riesgos?|plan|recomendaciones?|conclusi[oó]n|siguiente\s+paso)/i.test(plain);
  };
  for (const rawLine of lines) {
    const line = cleanAssistantLine(rawLine);
    if (!line) continue;
    if (isHeading(line)) {
      const title = line.replace(/^\d+[.)]\s*/, '').replace(/:$/, '').trim();
      current = { title, items: [] };
      sections.push(current);
      continue;
    }
    const bullet = line.replace(/^(?:[-•*]+|\d+[.)])\s*/, '').trim();
    if (current) current.items.push(bullet);
    else intro.push(bullet);
  }
  if (!sections.length && intro.length) {
    const text = intro.join(' ');
    return { intro: '', sections: [{ title: 'Análisis', items: [text] }] };
  }
  return { intro: intro.join(' '), sections };
}

function assistantActionItems(intent = '', question = '') {
  const value = normalizeText(`${intent} ${question}`);
  const items = [];
  const push = (key, label, target, focus = '') => {
    if (!items.some(item => item.key === key)) items.push({ key, label, target, focus });
  };
  if (/(reponer|reposicion|stock|faltante|inventario)/.test(value)) push('stock', 'Ver stock crítico', 'inventario', 'Bajo stock');
  if (/(lento|rotacion|sin vender|promocion)/.test(value)) push('products', 'Revisar productos', 'productos');
  if (/(cobranza|credito|deuda|vencid)/.test(value)) push('credits', 'Abrir créditos', 'creditos');
  if (/(venta|pago|rentab|utilidad|margen|vendedor|resumen)/.test(value)) push('reports', 'Ver reportes', 'reportes');
  if (!items.length) push('dashboard', 'Abrir panel dueño', 'panel');
  return items.slice(0, 3);
}

function assistantAnswerToText(answer, parsed) {
  const rows = [];
  if (answer?.question) rows.push(answer.question);
  if (parsed?.intro) rows.push(parsed.intro);
  (parsed?.sections || []).forEach(section => {
    rows.push(`${section.title}:`);
    (section.items || []).forEach(item => rows.push(`• ${item}`));
  });
  return rows.join('\n').trim();
}



function AssistantAI({ profile, products = [], store, onNavigate }) {
  const [days, setDays] = useState(30);
  const [question, setQuestion] = useState('');
  const [answer, setAnswer] = useState(null);
  const [asking, setAsking] = useState(false);
  const [notice, setNotice] = useState('');
  const [aiMode, setAiMode] = useState('erp');
  const [workspace, setWorkspace] = useState('decidir');
  const [selectedProductId, setSelectedProductId] = useState('');
  const [commercialTone, setCommercialTone] = useState('Cercano');
  const [commercialText, setCommercialText] = useState('');
  const [catalogText, setCatalogText] = useState('');
  const [collectionForm, setCollectionForm] = useState({ name: '', amount: '', phone: '', due_date: '' });
  const [collectionText, setCollectionText] = useState('');
  const [listening, setListening] = useState(false);
  const recognitionRef = useRef(null);
  const [history, setHistory] = useState(() => {
    try { return JSON.parse(localStorage.getItem('clomar_ai_history_v33') || localStorage.getItem('clomar_ai_history_v328') || '[]'); } catch (_) { return []; }
  });
  const [activity, setActivity] = useState(() => {
    try { return JSON.parse(localStorage.getItem('clomar_ai_activity_v33') || '[]'); } catch (_) { return []; }
  });

  useEffect(() => { try { localStorage.setItem('clomar_ai_history_v33', JSON.stringify(history.slice(0, 10))); } catch (_) {} }, [history]);
  useEffect(() => { try { localStorage.setItem('clomar_ai_activity_v33', JSON.stringify(activity.slice(0, 14))); } catch (_) {} }, [activity]);
  useEffect(() => () => { try { recognitionRef.current?.stop?.(); } catch (_) {} }, []);

  const selectedProduct = products.find(p => p.id === selectedProductId) || null;
  const lowStockProducts = useMemo(() => products.filter(p => p.active !== false && asNum(p.stock) <= asNum(p.stock_min || 0)).sort((a, b) => asNum(a.stock) - asNum(b.stock)).slice(0, 12), [products]);
  const currentRows = Array.isArray(answer?.data) ? answer.data : [];
  const parsed = answer ? parseAssistantAnswer(answer.answer) : null;
  const answerActions = answer ? assistantActionItems(answer.intent || inferAssistantIntent(answer.question || ''), answer.question || '') : [];
  const answerText = answer ? assistantAnswerToText(answer, parsed) : '';

  function recordActivity(label, type = 'consulta') {
    setActivity(prev => [{ id: `${Date.now()}-${type}`, label, type, created_at: new Date().toISOString() }, ...prev].slice(0, 14));
  }

  function productContext(product) {
    if (!product) return {};
    const availability = asNum(product.stock) <= 0 ? 'Agotado' : asNum(product.stock) <= asNum(product.stock_min || 2) ? 'Últimas unidades' : 'Disponible';
    return {
      name: product.name || '',
      code: product.code || product.barcode || '',
      price: asNum(product.price),
      category: product.category || '',
      brand: product.brand || '',
      color: product.color || '',
      size: product.size || '',
      availability,
      description: product.catalog_description || product.description || '',
    };
  }

  function fallbackCommercial(product, tone = 'Cercano') {
    if (!product) return '';
    const availability = asNum(product.stock) <= 0 ? 'En este momento figura agotado' : asNum(product.stock) <= asNum(product.stock_min || 2) ? 'Quedan últimas unidades' : 'Está disponible';
    const detail = [product.brand, product.color ? `color ${product.color}` : '', product.size ? `talla ${product.size}` : ''].filter(Boolean).join(' · ');
    const greeting = tone === 'Formal' ? 'Hola, gracias por escribir a Clomar Store.' : tone === 'Breve' ? 'Hola.' : 'Hola, gracias por comunicarte con Clomar Store.';
    const closing = tone === 'Formal' ? '¿Desea que verifiquemos la disponibilidad final o le ayudemos con otra talla o color?' : tone === 'Breve' ? '¿Desea reservarlo?' : '¿Desea que le ayudemos a reservarlo o revisar otra talla o color?';
    return `${greeting}\n\n${product.name}${detail ? ` (${detail})` : ''}\nPrecio: ${money(product.price)}\n${availability}.\nCódigo: ${product.code || product.barcode || '—'}\n\n${closing}`;
  }

  function fallbackCatalog(product) {
    if (!product) return '';
    const details = [product.brand, product.color, product.size].filter(Boolean).join(' · ');
    return `Nombre comercial\n${product.name}\n\nDescripción para catálogo\n${product.description || `Producto disponible en Clomar Store${details ? `: ${details}` : ''}. Consulte disponibilidad antes de confirmar su pedido.`}\n\nTexto para WhatsApp\n${fallbackCommercial(product, 'Cercano')}`;
  }

  function normalizeGeneratedText(value = '') {
    return String(value || '')
      .replace(/^\s*Asistente IA · Gemini \+ ERP\s*/im, '')
      .replace(/\*\*/g, '')
      .trim();
  }

  async function callAssistant({ prompt, intent = 'resumen', task = 'analysis', context = {}, saveAsAnswer = true, addHistory = true }) {
    const cleanPrompt = String(prompt || '').trim();
    if (!cleanPrompt) return null;
    setAsking(true);
    setNotice('');
    let nextAnswer = null;
    try {
      const { data, error } = await supabase.functions.invoke('clomar-ai', {
        body: {
          question: cleanPrompt,
          intent,
          task,
          days: Number(days || 30),
          store_id: profile?.store_id || DEFAULT_STORE_ID,
          context,
        },
      });
      if (error) throw error;
      if (!data?.answer) throw new Error('El asistente no devolvió una respuesta.');
      nextAnswer = { ...data, question: cleanPrompt, intent, task, title: data.title || (data.mode === 'gemini' ? 'Asistente IA · Gemini + ERP' : 'Análisis ERP verificado') };
      if (saveAsAnswer) {
        setAnswer(nextAnswer);
        setAiMode(data.mode === 'gemini' ? 'gemini' : 'erp');
      }
      if (data.notice) setNotice(data.notice);
    } catch (error) {
      if (task === 'analysis' || task === 'purchase_plan') {
        try {
          const { data, error: fallbackError } = await supabase.rpc('clomar_management_assistant_v32', {
            p_store_id: profile?.store_id || DEFAULT_STORE_ID,
            p_intent: intent,
            p_days: Number(days || 30),
          });
          if (fallbackError) throw fallbackError;
          nextAnswer = { ...data, question: cleanPrompt, intent, task, title: 'Análisis ERP verificado' };
          if (saveAsAnswer) {
            setAnswer(nextAnswer);
            setAiMode('erp');
          }
          setNotice('Gemini no respondió en este momento. Se muestra el análisis verificado del ERP.');
        } catch (fallbackError) {
          setNotice(fallbackError.message || error.message || 'No se pudo obtener la respuesta del asistente.');
        }
      } else {
        setNotice(error.message || 'No se pudo generar el borrador con IA. Se aplicó un formato local de respaldo cuando está disponible.');
      }
    } finally {
      if (addHistory) {
        setHistory(prev => [{ id: `${Date.now()}-${intent}`, question: cleanPrompt, intent, task, created_at: new Date().toISOString() }, ...prev.filter(item => item.question !== cleanPrompt)].slice(0, 10));
      }
      setAsking(false);
    }
    return nextAnswer;
  }

  async function askAssistant(rawQuestion, forcedIntent = null, task = 'analysis', context = {}) {
    const prompt = String(rawQuestion || question || '').trim();
    const intent = forcedIntent || inferAssistantIntent(prompt);
    const result = await callAssistant({ prompt, intent, task, context });
    if (result) recordActivity(`Consulta IA: ${prompt.slice(0, 72)}`, 'consulta');
    return result;
  }

  async function generateCommercialReply() {
    if (!selectedProduct) { setNotice('Seleccione un producto para preparar una respuesta comercial.'); return; }
    const result = await callAssistant({
      prompt: `Prepara un mensaje de WhatsApp ${commercialTone.toLowerCase()} para consultar o vender este producto. Incluye únicamente precio, código y disponibilidad verificados. No ofrezcas descuentos ni reservas automáticas.`,
      intent: 'resumen',
      task: 'whatsapp_message',
      context: { product: productContext(selectedProduct), tone: commercialTone },
      saveAsAnswer: false,
      addHistory: false,
    });
    const text = normalizeGeneratedText(result?.answer || fallbackCommercial(selectedProduct, commercialTone));
    setCommercialText(text);
    setWorkspace('comercial');
    recordActivity(`Mensaje comercial preparado: ${selectedProduct.name}`, 'whatsapp');
    if (!result) setNotice('Se preparó un mensaje local con datos reales del producto.');
  }

  async function generateCatalogCopy() {
    if (!selectedProduct) { setNotice('Seleccione un producto para generar la ficha comercial.'); return; }
    const result = await callAssistant({
      prompt: 'Crea una ficha comercial breve para catálogo con nombre comercial, descripción clara y texto corto para WhatsApp. Usa solo los datos del producto proporcionado. No inventes atributos, descuentos ni disponibilidad.',
      intent: 'resumen',
      task: 'catalog_copy',
      context: { product: productContext(selectedProduct) },
      saveAsAnswer: false,
      addHistory: false,
    });
    const text = normalizeGeneratedText(result?.answer || fallbackCatalog(selectedProduct));
    setCatalogText(text);
    setWorkspace('catalogo');
    recordActivity(`Ficha de catálogo preparada: ${selectedProduct.name}`, 'catalogo');
    if (!result) setNotice('Se preparó una ficha local con los datos verificados del producto.');
  }

  async function generateCollectionReply() {
    const name = String(collectionForm.name || '').trim();
    const amount = asNum(collectionForm.amount);
    if (!name || amount <= 0) { setNotice('Ingrese el nombre del cliente y el monto pendiente para preparar el mensaje.'); return; }
    const result = await callAssistant({
      prompt: 'Redacta un mensaje respetuoso de recordatorio de pago. No uses lenguaje de presión, amenaza o descuento. Incluye únicamente el nombre, saldo y fecha proporcionados.',
      intent: 'cobranza',
      task: 'collection_message',
      context: { customer: { name, amount, due_date: collectionForm.due_date || '' } },
      saveAsAnswer: false,
      addHistory: false,
    });
    const fallback = `Hola, ${name}. Le recordamos que tiene un saldo pendiente de ${money(amount)}${collectionForm.due_date ? ` con fecha de referencia ${collectionForm.due_date}` : ''}. ¿Podemos ayudarle a coordinar su pago? Gracias por su atención.`;
    setCollectionText(normalizeGeneratedText(result?.answer || fallback));
    setWorkspace('cobranza');
    recordActivity(`Mensaje de cobranza preparado: ${name}`, 'cobranza');
    if (!result) setNotice('Se preparó un recordatorio local. Revíselo antes de enviarlo.');
  }

  function startVoiceInput() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      setNotice('El dictado por voz no está disponible en este navegador. Use Chrome en Android o escriba la consulta.');
      return;
    }
    try { recognitionRef.current?.stop?.(); } catch (_) {}
    const recognition = new SpeechRecognition();
    recognition.lang = 'es-PE';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;
    recognition.onstart = () => setListening(true);
    recognition.onend = () => setListening(false);
    recognition.onerror = () => { setListening(false); setNotice('No se pudo captar el audio. Revise el permiso del micrófono e inténtelo otra vez.'); };
    recognition.onresult = (event) => {
      const spoken = event?.results?.[0]?.[0]?.transcript || '';
      setQuestion(prev => `${prev ? `${prev} ` : ''}${spoken}`.trim());
    };
    recognitionRef.current = recognition;
    recognition.start();
  }

  function goTo(target, focus = '') {
    try {
      if (target === 'inventario' && focus) sessionStorage.setItem('clomar_inventory_filter_v328', focus);
      if (target === 'productos' && focus) sessionStorage.setItem('clomar_products_focus_v328', focus);
    } catch (_) {}
    onNavigate?.(target);
  }

  function buildPurchasePlan() {
    const fromAnswer = currentRows.filter(row => row?.name).slice(0, 12);
    const source = fromAnswer.length ? fromAnswer : lowStockProducts;
    return source.map((row, index) => {
      const stock = asNum(row.stock);
      const stockMin = asNum(row.stock_min || row.minimum || row.min_stock || 0);
      const salesHint = asNum(row.qty || row.units_sold || row.sales_qty || 0);
      const suggested = Math.max(1, stockMin > 0 ? Math.max(stockMin * 3 - stock, stockMin - stock) : salesHint > 0 ? salesHint : 1);
      return {
        id: row.product_id || row.id || `${row.code || row.name || 'item'}-${index}`,
        name: row.name || 'Producto',
        code: row.code || 'Sin código',
        stock,
        stock_min: stockMin,
        suggested,
        reason: stockMin > 0 ? `Stock ${fmtWhole(stock)} / mínimo ${fmtWhole(stockMin)}` : 'Revisión manual recomendada',
      };
    });
  }

  const purchasePlan = useMemo(() => buildPurchasePlan(), [answer, products]);

  function exportPurchasePlan() {
    if (!purchasePlan.length) { setNotice('No hay productos para exportar en el borrador de compra.'); return; }
    const ws = XLSX.utils.json_to_sheet(purchasePlan.map((row, idx) => ({
      Prioridad: idx + 1,
      Producto: row.name,
      Código: row.code,
      'Stock actual': row.stock,
      'Stock mínimo': row.stock_min,
      'Cantidad sugerida': row.suggested,
      Motivo: row.reason,
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reposición');
    XLSX.writeFile(wb, `clomar_lista_reposicion_${todayISO()}.xlsx`);
    recordActivity('Lista de reposición exportada a Excel', 'reporte');
  }

  function printCommandReport() {
    const safeTitle = escapeHtml(answer?.title || 'Reporte IA de Clomar Store');
    const safeQuestion = escapeHtml(answer?.question || 'Resumen operativo');
    const summary = escapeHtml(answerText || 'Aún no se ha generado una respuesta.');
    const rows = purchasePlan.slice(0, 10).map((row, idx) => `<tr><td>${idx + 1}</td><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.code)}</td><td>${row.stock}</td><td>${row.suggested}</td></tr>`).join('') || '<tr><td colspan="5">No hay lista de reposición para este reporte.</td></tr>';
    const activityRows = activity.slice(0, 6).map(item => `<li><strong>${escapeHtml(item.label)}</strong><small>${escapeHtml(fmtDate(item.created_at))}</small></li>`).join('') || '<li>Sin actividad registrada todavía.</li>';
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>Reporte IA Clomar Store</title><style>
      *{box-sizing:border-box}body{margin:0;background:#f3f6fb;color:#17243a;font-family:Arial,Helvetica,sans-serif;font-size:12px}.report{width:190mm;margin:0 auto;background:#fff;padding:15mm}.head{display:flex;justify-content:space-between;gap:20px;border-bottom:2px solid #214b82;padding-bottom:12px}.brand{font-size:22px;font-weight:900;color:#17345e}.muted{color:#64748b}.badge{display:inline-block;padding:5px 8px;border-radius:999px;background:#ecfdf5;color:#166534;font-size:10px;font-weight:800}.block{margin-top:16px;border:1px solid #dce7f4;border-radius:12px;padding:12px}.block h2{margin:0 0 8px;font-size:14px}.analysis{white-space:pre-line;line-height:1.65;color:#334155}.note{font-size:10px;color:#64748b}.table{width:100%;border-collapse:collapse}.table th,.table td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left}.table th{font-size:10px;text-transform:uppercase;color:#64748b}.activity{margin:0;padding-left:16px}.activity li{margin:7px 0}.activity small{display:block;color:#64748b;margin-top:2px}@media print{@page{size:A4;margin:10mm}body{background:#fff}.report{width:auto;margin:0;padding:0}}
@media print{.price strong{white-space:nowrap!important}.composition-two-column .price strong{white-space:nowrap!important}.composition-two-column.price-max .price strong{font-size:13.6px!important}.label.orientation-vertical.vertical-compact{gap:.38mm!important;align-content:start!important}.label.orientation-vertical.vertical-compact .codes{align-items:center!important}.label.orientation-vertical.price-max .price strong{font-size:26px!important}.density-vertical-medium.orientation-vertical{grid-template-rows:3.95mm 11.9mm 14.3mm 25.4mm!important;gap:.38mm!important;padding:1.95mm 2.05mm 1.6mm!important}.density-vertical-medium.orientation-vertical.price-max .price strong{font-size:25.8px!important}.density-vertical-medium.orientation-vertical .price strong{font-size:22.8px!important}.density-vertical-medium.orientation-vertical.mode-barcode .barcode-svg{height:17.4mm!important}.density-vertical-compact.orientation-vertical{grid-template-rows:3.65mm 10.4mm 13.4mm 24.2mm!important;gap:.34mm!important;padding:1.7mm 1.7mm 1.45mm!important}.density-vertical-compact.orientation-vertical .price strong{font-size:20.2px!important}.density-vertical-compact.orientation-vertical.price-max .price strong{font-size:22.8px!important}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-svg{height:16.8mm!important}.density-vertical-large.orientation-vertical{grid-template-rows:5.1mm 18.6mm 18.7mm 32.4mm!important;gap:.58mm!important;padding:2.75mm 2.95mm 2.35mm!important}.density-vertical-large.orientation-vertical.price-max .price strong{font-size:33px!important}.density-vertical-large.orientation-vertical .price strong{font-size:29.2px!important}.label.orientation-vertical.composition-two-column .price{display:flex!important;flex-direction:row!important;align-items:baseline!important}.label.orientation-vertical.composition-two-column .price strong{white-space:nowrap!important;font-size:22px!important}.qr-large .qr-wrap{grid-template-columns:13.8mm!important;grid-template-rows:13.8mm auto!important}.qr-large .qr{width:13.8mm!important;height:13.8mm!important}.qr-max .qr-wrap{grid-template-columns:17mm!important;grid-template-rows:17mm auto!important}.qr-max .qr{width:17mm!important;height:17mm!important}}
</style></head><body><main class="report"><section class="head"><div><div class="brand">${escapeHtml(store?.name || 'Clomar Store')}</div><div class="muted">Reporte ejecutivo generado desde Clomar AI Command Center</div></div><div><div class="badge">Datos ERP + Gemini</div><div class="muted" style="margin-top:6px">${escapeHtml(fmtDate(new Date()))}</div></div></section><section class="block"><h2>${safeTitle}</h2><p class="muted"><strong>Consulta:</strong> ${safeQuestion}</p><div class="analysis">${summary}</div></section><section class="block"><h2>Borrador de reposición</h2><table class="table"><thead><tr><th>#</th><th>Producto</th><th>Código</th><th>Stock</th><th>Sugerido</th></tr></thead><tbody>${rows}</tbody></table><p class="note">La cantidad sugerida es un borrador. Debe revisarse antes de registrar una compra.</p></section><section class="block"><h2>Actividad del asistente</h2><ul class="activity">${activityRows}</ul></section><p class="note">Este informe interpreta datos del ERP. No reemplaza la revisión de caja, inventario ni decisiones del propietario.</p></main><script>window.onload=()=>setTimeout(()=>window.print(),250)</script></body></html>`;
    const win = window.open('', '_blank', 'width=900,height=760');
    if (!win) { setNotice('El navegador bloqueó la ventana de impresión. Habilite ventanas emergentes para generar el PDF.'); return; }
    win.document.open(); win.document.write(html); win.document.close();
    recordActivity('Reporte ejecutivo preparado para PDF', 'reporte');
  }

  const modeText = aiMode === 'gemini' ? 'Gemini + ERP conectado' : 'ERP verificado';
  const commercialWhatsapp = `https://wa.me/${String(store?.whatsapp_number || '51931709871').replace(/\D/g,'')}?text=${encodeURIComponent(commercialText || '')}`;
  const collectionWhatsapp = `https://wa.me/${String(collectionForm.phone || '').replace(/\D/g,'')}?text=${encodeURIComponent(collectionText || '')}`;

  return (
    <div className="page ai-assistant-page ai-v33-page">
      <section className="ai-v33-commandbar">
        <div className="ai-v33-brandcopy"><span className="eyebrow">Clomar AI Command Center</span><h1>Copiloto operativo</h1><p>Analiza datos reales, prepara borradores y abre el módulo correcto. Toda acción que cambia el negocio sigue requiriendo confirmación humana.</p></div>
        <div className={`ai-engine-badge ${aiMode === 'gemini' ? 'online' : ''}`}><span></span>{modeText}</div>
        <label className="ai-period-compact">Periodo<select value={days} onChange={e => setDays(Number(e.target.value))}><option value="7">7 días</option><option value="30">30 días</option><option value="60">60 días</option><option value="90">90 días</option></select></label>
      </section>
      <section className="ai-safety-strip ai-v33-safety"><strong>Control seguro:</strong> la IA puede analizar, buscar, redactar, exportar y crear borradores. No cambia precios, stock, caja, pagos, créditos ni ventas sin una confirmación posterior del usuario.</section>
      {notice && <div className="catalog-toast ai-toast">{notice}</div>}

      <section className="ai-v33-command-grid">
        <button type="button" className="ai-v33-command-card stock" onClick={() => { setWorkspace('compra'); askAssistant('Prioriza los productos que debo reponer esta semana. Indica riesgo, motivo y qué revisar antes de comprar.', 'reponer', 'purchase_plan'); }}><span>01</span><div><strong>Reponer stock</strong><small>{lowStockProducts.length} producto(s) requieren revisión</small></div><b>Preparar compra →</b></button>
        <button type="button" className="ai-v33-command-card sales" onClick={() => { setWorkspace('decidir'); askAssistant('Analiza ventas, utilidad y productos lentos. Dame las tres decisiones comerciales más importantes para esta semana.', 'resumen'); }}><span>02</span><div><strong>Decidir qué vender</strong><small>Ventas, margen y rotación</small></div><b>Analizar →</b></button>
        <button type="button" className="ai-v33-command-card collections" onClick={() => { setWorkspace('cobranza'); askAssistant('Resume la cobranza pendiente y señala qué riesgo debo atender primero.', 'cobranza'); }}><span>03</span><div><strong>Gestionar cobros</strong><small>Créditos pendientes y vencidos</small></div><b>Ver cobranza →</b></button>
        <button type="button" className="ai-v33-command-card catalog" onClick={() => setWorkspace('catalogo')}><span>04</span><div><strong>Mejorar catálogo</strong><small>Ficha, descripción y WhatsApp</small></div><b>Crear contenido →</b></button>
      </section>

      <div className="ai-v33-workspace">
        <section className="card compact-card ai-v33-chat-card">
          <header className="assistant-section-head ai-v33-section-head"><div><span className="eyebrow">Consulta estratégica</span><h3>¿Qué necesita resolver ahora?</h3></div><button type="button" className="ai-daily-brief" onClick={() => askAssistant('Genera mi resumen operativo de hoy con ventas, caja, stock crítico, créditos y una acción prioritaria.', 'resumen')}>Resumen del día</button></header>
          <div className="ai-v33-composer"><textarea value={question} onChange={e => setQuestion(e.target.value)} rows="3" placeholder="Ej.: ¿Qué productos debo reponer primero y cuáles debo promocionar?" /><div className="ai-v33-composer-actions"><button type="button" className={`secondary-btn ai-voice-btn ${listening ? 'listening' : ''}`} onClick={startVoiceInput}>{listening ? 'Escuchando…' : 'Dictar'}</button><button type="button" className="primary-btn" disabled={asking} onClick={() => askAssistant()}>{asking ? 'Analizando…' : 'Preguntar'}</button></div></div>
          <div className="ai-v33-quick-grid">{ASSISTANT_QUICK_QUESTIONS.map(item => <button type="button" key={item.intent} onClick={() => { setQuestion(item.text); askAssistant(item.text, item.intent); }}>{item.label}</button>)}</div>

          {answer ? <article className="ai-v33-answer-card">
            <div className="ai-answer-head"><div><span className="eyebrow">{answer.title || 'Respuesta verificada'}</span><h3>{answer.question}</h3></div><button type="button" className="icon-btn ai-copy-btn" title="Copiar respuesta" onClick={() => { copyTextToClipboard(answerText); setNotice('Respuesta copiada.'); recordActivity('Respuesta IA copiada', 'accion'); }}>⧉</button></div>
            {parsed?.intro && <p className="ai-v33-answer-intro">{parsed.intro}</p>}
            <div className="ai-v33-sections">{(parsed?.sections || []).map((section, idx) => <section className="ai-v33-section" key={`${section.title}-${idx}`}><header><span>{String(idx + 1).padStart(2, '0')}</span><h4>{section.title}</h4></header><div>{(section.items || []).map((item, itemIdx) => <p key={itemIdx}>{item}</p>)}</div></section>)}</div>
            {Array.isArray(answer.data) && answer.data.length > 0 && <details className="ai-v33-evidence"><summary>Ver datos del ERP utilizados</summary><div className="ai-result-list">{answer.data.slice(0, 8).map((row, idx) => <div className="list-row" key={idx}><span><strong>{row.name || row.customer_name || row.seller || row.product_a || row.method || row.title || 'Dato'}</strong><small>{row.code || row.product_b || row.message || row.due_date || ''}</small></span><b>{row.amount !== undefined ? money(row.amount) : row.balance !== undefined ? money(row.balance) : row.stock !== undefined ? fmtWhole(row.stock) : row.qty !== undefined ? fmtWhole(row.qty) : row.times_together !== undefined ? fmtWhole(row.times_together) : ''}</b></div>)}</div></details>}
            <footer className="ai-v33-actionbar"><span>Acciones disponibles</span><div>{answerActions.map(action => <button type="button" className="secondary-btn" key={action.key} onClick={() => { recordActivity(`Abrir módulo: ${action.label}`, 'navegacion'); goTo(action.target, action.focus); }}>{action.label} →</button>)}<button type="button" className="secondary-btn" onClick={() => { setWorkspace('reportes'); }}>Reporte PDF →</button></div></footer>
          </article> : <div className="assistant-empty-state ai-v33-empty"><span>IA</span><div><strong>Su centro de decisiones está listo</strong><small>Pregunte por reposición, ventas, utilidad, cobranza, catálogo o un plan comercial.</small></div></div>}
          {history.length > 0 && <div className="assistant-history ai-v33-history"><div><span>Consultas recientes</span><button type="button" onClick={() => setHistory([])}>Limpiar</button></div><section>{history.map(item => <button type="button" key={item.id} onClick={() => { setQuestion(item.question); askAssistant(item.question, item.intent, item.task || 'analysis'); }}><span>{item.question}</span><small>{fmtDate(item.created_at)}</small></button>)}</section></div>}
        </section>

        <aside className="card compact-card ai-v33-command-desk">
          <div className="ai-v33-desk-tabs"><button type="button" className={workspace === 'compra' ? 'active' : ''} onClick={() => setWorkspace('compra')}>Compra</button><button type="button" className={workspace === 'comercial' ? 'active' : ''} onClick={() => setWorkspace('comercial')}>WhatsApp</button><button type="button" className={workspace === 'catalogo' ? 'active' : ''} onClick={() => setWorkspace('catalogo')}>Catálogo</button><button type="button" className={workspace === 'cobranza' ? 'active' : ''} onClick={() => setWorkspace('cobranza')}>Cobranza</button><button type="button" className={workspace === 'reportes' ? 'active' : ''} onClick={() => setWorkspace('reportes')}>Reportes</button></div>

          {(workspace === 'comercial' || workspace === 'catalogo') && <section className="ai-v33-desk-section ai-v33-product-tools"><header><span className="eyebrow">Producto real</span><h3>{workspace === 'catalogo' ? 'Contenido para catálogo' : 'Respuesta comercial'}</h3><p>La IA recibe solo nombre, precio, código y características verificadas del producto seleccionado.</p></header><label>Producto<select value={selectedProductId} onChange={e => setSelectedProductId(e.target.value)}><option value="">Seleccione un producto</option>{products.filter(p => p.active !== false && asNum(p.price) > 0 && productPriceStatus(p) === 'Validado').slice(0, 500).map(p => <option value={p.id} key={p.id}>{p.name} · {money(p.price)}{p.color ? ` · ${p.color}` : ''}{p.size ? ` · ${p.size}` : ''}</option>)}</select></label>{workspace === 'comercial' && <label>Tono<select value={commercialTone} onChange={e => setCommercialTone(e.target.value)}><option>Cercano</option><option>Formal</option><option>Breve</option></select></label>}<button type="button" className="primary-btn" disabled={asking} onClick={workspace === 'catalogo' ? generateCatalogCopy : generateCommercialReply}>{workspace === 'catalogo' ? 'Generar ficha comercial' : 'Preparar mensaje WhatsApp'}</button>{workspace === 'catalogo' && catalogText && <div className="ai-v33-output"><textarea value={catalogText} onChange={e => setCatalogText(e.target.value)} rows="10" /><div className="button-row"><button type="button" className="secondary-btn" onClick={() => { copyTextToClipboard(catalogText); setNotice('Ficha de catálogo copiada.'); }}>Copiar</button><button type="button" className="secondary-btn" onClick={() => goTo('catalogo')}>Abrir catálogo</button></div></div>}{workspace === 'comercial' && commercialText && <div className="ai-v33-output"><textarea value={commercialText} onChange={e => setCommercialText(e.target.value)} rows="9" /><div className="button-row"><button type="button" className="secondary-btn" onClick={() => { copyTextToClipboard(commercialText); setNotice('Mensaje comercial copiado.'); }}>Copiar</button><a className="primary-btn" href={commercialWhatsapp} target="_blank" rel="noreferrer" onClick={() => recordActivity('Mensaje comercial abierto en WhatsApp', 'whatsapp')}>Abrir WhatsApp</a></div></div>}</section>}

          {workspace === 'compra' && <section className="ai-v33-desk-section"><header><span className="eyebrow">Borrador de compra</span><h3>Reposición inteligente</h3><p>La lista propone cantidades orientativas. Revise proveedor, costo y demanda antes de registrar la compra.</p></header><div className="ai-v33-plan-list">{purchasePlan.length ? purchasePlan.map((row, idx) => <div className="ai-v33-plan-row" key={row.id}><span className="ai-v33-plan-number">{idx + 1}</span><div><strong>{row.name}</strong><small>{row.code} · {row.reason}</small></div><b>{row.suggested} u.</b></div>) : <div className="commercial-empty-state"><span>—</span><div><strong>No hay productos para el borrador</strong><small>Ejecute “Reponer stock” o cargue productos con stock mínimo.</small></div></div>}</div><div className="ai-v33-desk-actions"><button type="button" className="secondary-btn" onClick={() => { setWorkspace('compra'); askAssistant('Prioriza los productos que debo reponer esta semana. Indica riesgo, motivo y qué revisar antes de comprar.', 'reponer', 'purchase_plan'); }}>Actualizar análisis</button><button type="button" className="primary-btn" onClick={exportPurchasePlan}>Exportar Excel</button></div><button type="button" className="link-action" onClick={() => goTo('ingresos')}>Registrar compra manualmente →</button></section>}

          {workspace === 'cobranza' && <section className="ai-v33-desk-section"><header><span className="eyebrow">Cobranza responsable</span><h3>Preparar recordatorio</h3><p>La IA redacta un mensaje cordial. No confirma pagos ni modifica la cuenta del cliente.</p></header><div className="ai-v33-form-grid"><label>Cliente<input value={collectionForm.name} onChange={e => setCollectionForm({ ...collectionForm, name: e.target.value })} placeholder="Nombre del cliente" /></label><label>Saldo pendiente<input value={collectionForm.amount} onChange={e => setCollectionForm({ ...collectionForm, amount: e.target.value })} inputMode="decimal" placeholder="0.00" /></label><label>WhatsApp<input value={collectionForm.phone} onChange={e => setCollectionForm({ ...collectionForm, phone: e.target.value })} inputMode="tel" placeholder="519XXXXXXXX" /></label><label>Fecha de referencia<input value={collectionForm.due_date} onChange={e => setCollectionForm({ ...collectionForm, due_date: e.target.value })} type="date" /></label></div><button type="button" className="primary-btn" disabled={asking} onClick={generateCollectionReply}>Generar mensaje de cobranza</button>{collectionText && <div className="ai-v33-output"><textarea value={collectionText} onChange={e => setCollectionText(e.target.value)} rows="8" /><div className="button-row"><button type="button" className="secondary-btn" onClick={() => { copyTextToClipboard(collectionText); setNotice('Mensaje de cobranza copiado.'); }}>Copiar</button>{collectionForm.phone ? <a className="primary-btn" href={collectionWhatsapp} target="_blank" rel="noreferrer" onClick={() => recordActivity(`Recordatorio de cobranza abierto: ${collectionForm.name}`, 'cobranza')}>Abrir WhatsApp</a> : <button type="button" className="secondary-btn" onClick={() => goTo('creditos')}>Abrir créditos</button>}</div></div>}<button type="button" className="link-action" onClick={() => goTo('creditos')}>Ver cuentas por cobrar →</button></section>}

          {workspace === 'reportes' && <section className="ai-v33-desk-section"><header><span className="eyebrow">Reporte ejecutivo</span><h3>PDF y Excel bajo demanda</h3><p>El ERP aporta las cifras; la IA redacta el diagnóstico. Revise el contenido antes de compartirlo.</p></header><div className="ai-v33-report-preview"><div><span>Respuesta actual</span><strong>{answer ? answer.title : 'Aún no hay análisis'}</strong></div><div><span>Lista de reposición</span><strong>{purchasePlan.length} producto(s)</strong></div><div><span>Actividad local</span><strong>{activity.length} acción(es)</strong></div></div><div className="ai-v33-desk-actions"><button type="button" className="primary-btn" disabled={!answer} onClick={printCommandReport}>Generar PDF</button><button type="button" className="secondary-btn" onClick={exportPurchasePlan}>Excel reposición</button></div><small className="muted">El PDF se abre en una ventana de impresión. Seleccione “Guardar como PDF”. El historial de actividad se guarda localmente en este navegador.</small></section>}

          {workspace === 'decidir' && <section className="ai-v33-desk-section"><header><span className="eyebrow">Acciones seguras</span><h3>Cómo usar el copiloto</h3><p>Empiece por una consulta o elija una orden rápida. La IA analiza; usted decide y confirma.</p></header><div className="ai-v33-guide-list"><div><b>1</b><span><strong>Analizar</strong><small>Ventas, stock, caja, créditos y rotación.</small></span></div><div><b>2</b><span><strong>Preparar</strong><small>Listas de compra, mensajes y contenido comercial.</small></span></div><div><b>3</b><span><strong>Confirmar</strong><small>Las acciones reales se registran en los módulos del ERP.</small></span></div></div><button type="button" className="primary-btn" onClick={() => askAssistant('Con los datos disponibles, dime las tres decisiones más importantes que debo tomar esta semana y qué módulo debo abrir para cada una.', 'resumen')}>Generar prioridades</button></section>}
        </aside>
      </div>
    </div>
  );
}


function POS({ products, reloadProducts, customers, profile, store, onGoReceipts, cashSession, menuOpen = false }) {
  const [query, setQuery] = useState('');
  const [cart, setCart] = useState([]);
  const [method, setMethod] = useState('Efectivo');
  const [documentType, setDocumentType] = useState('Interno');
  const [customerDocType, setCustomerDocType] = useState('DNI');
  const [customerDocNumber, setCustomerDocNumber] = useState('');
  const [customer, setCustomer] = useState('Consumidor final');
  const [selectedCustomerId, setSelectedCustomerId] = useState('');
  const [checkoutCustomers, setCheckoutCustomers] = useState(customers || []);
  const [customerQuery, setCustomerQuery] = useState('');
  const [customerQuickOpen, setCustomerQuickOpen] = useState(false);
  const [customerPickerOpen, setCustomerPickerOpen] = useState(false);
  const [quickCustomerSaving, setQuickCustomerSaving] = useState(false);
  const [quickCustomer, setQuickCustomer] = useState({ name: '', document_type: 'DNI', document: '', phone: '', address: '' });
  const [saving, setSaving] = useState(false);
  const [lastTicket, setLastTicket] = useState(null);
  const [saleModal, setSaleModal] = useState(null);
  const [dismissedTicketId, setDismissedTicketId] = useState(null);
  const [scanOpen, setScanOpen] = useState(false);
  const [scanStatus, setScanStatus] = useState('');
  const [notice, setNotice] = useState(null);
  const [globalDiscount, setGlobalDiscount] = useState('0');
  const [showItemDiscounts, setShowItemDiscounts] = useState(false);
  const [showGlobalDiscount, setShowGlobalDiscount] = useState(false);
  const [showMorePaymentMethods, setShowMorePaymentMethods] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [mobileCartOpen, setMobileCartOpen] = useState(false);
  const [mobileCheckoutStep, setMobileCheckoutStep] = useState('items');
  const [mixedPayments, setMixedPayments] = useState({ Efectivo: '', Yape: '', Plin: '', Transferencia: '', Tarjeta: '' });
  const [productView, setProductView] = useState('Todos');
  const [categoryFilter, setCategoryFilter] = useState('Todas');
  const [favoriteIds, setFavoriteIds] = useState(() => { try { return JSON.parse(localStorage.getItem('clomar_pos_favorites_v322') || '[]'); } catch (_) { return []; } });
  const [recentIds, setRecentIds] = useState(() => { try { return JSON.parse(localStorage.getItem('clomar_pos_recent_v322') || '[]'); } catch (_) { return []; } });
  useEffect(() => {
    try {
      const requested = sessionStorage.getItem('clomar_pos_assistant_search');
      if (requested) {
        setQuery(requested);
        sessionStorage.removeItem('clomar_pos_assistant_search');
        setTimeout(() => searchInputRef.current?.focus(), 120);
      }
    } catch (_) {}
  }, []);
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const searchInputRef = useRef(null);
  const normalized = query.trim().toLowerCase();
  const activeProducts = useMemo(() => products.filter(p => p.active !== false), [products]);
  const productCategories = useMemo(() => [...new Set(activeProducts.map(p => String(p.category || '').trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'es')), [activeProducts]);
  const fiscalMeta = documentMeta(documentType);
  const isDefaultCheckoutCustomer = !customer || ['Cliente', 'Cliente general', 'Consumidor final'].includes(customer);
  const matches = useMemo(() => {
    let base = activeProducts.filter(p => categoryFilter === 'Todas' || String(p.category || '') === categoryFilter);
    if (productView === 'Favoritos') base = base.filter(p => favoriteIds.includes(p.id));
    if (productView === 'Recientes') {
      const order = new Map(recentIds.map((id, idx) => [id, idx]));
      base = base.filter(p => order.has(p.id)).sort((a,b) => order.get(a.id) - order.get(b.id));
    }
    if (normalized) base = base.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(normalized));
    if (!normalized && productView === 'Todos' && categoryFilter === 'Todas') base = base.slice(0, 24);
    return base.slice(0, 40);
  }, [activeProducts, normalized, productView, categoryFilter, favoriteIds, recentIds]);
  const customerMatches = useMemo(() => {
    const needle = customerQuery.trim().toLowerCase();
    if (!needle) return checkoutCustomers.slice(0, 5);
    return checkoutCustomers.filter(c => `${c.name || ''} ${c.document || ''} ${c.phone || ''}`.toLowerCase().includes(needle)).slice(0, 6);
  }, [checkoutCustomers, customerQuery]);
  const lineBase = (item) => asNum(item.price) * asNum(item.qty);
  const lineDiscount = (item) => Math.min(lineBase(item), Math.max(0, asNum(item.discount)));
  const lineSubtotal = (item) => Math.max(0, lineBase(item) - lineDiscount(item));
  const subtotal = cart.reduce((sum, item) => sum + lineBase(item), 0);
  const itemDiscountTotal = cart.reduce((sum, item) => sum + lineDiscount(item), 0);
  const afterItemDiscount = Math.max(0, subtotal - itemDiscountTotal);
  const saleDiscount = Math.min(afterItemDiscount, Math.max(0, asNum(globalDiscount)));
  const total = Math.max(0, afterItemDiscount - saleDiscount);
  const mixedTotal = Object.values(mixedPayments).reduce((sum, v) => sum + asNum(v), 0);
  const mixedBalance = total - mixedTotal;
  const paymentOk = method !== 'Mixto' || Math.abs(mixedBalance) < 0.01;
  const checkoutButtonLabel = !cart.length
    ? 'Agregar productos'
    : saving
      ? 'Guardando...'
      : method === 'Crédito'
        ? 'Registrar crédito'
        : documentType === 'Interno'
          ? 'Cobrar'
          : documentType === 'Boleta'
            ? 'Registrar boleta pendiente'
            : 'Registrar factura pendiente';

  useEffect(() => { setCheckoutCustomers(customers || []); }, [customers]);
  useEffect(() => { try { localStorage.setItem('clomar_pos_favorites_v322', JSON.stringify(favoriteIds)); } catch (_) {} }, [favoriteIds]);
  useEffect(() => { try { localStorage.setItem('clomar_pos_recent_v322', JSON.stringify(recentIds)); } catch (_) {} }, [recentIds]);
  useEffect(() => { if (menuOpen) setMobileCartOpen(false); }, [menuOpen]);
  useEffect(() => { if (!mobileCartOpen) setMobileCheckoutStep('items'); }, [mobileCartOpen]);
  useEffect(() => {
    const shouldLock = mobileCartOpen || customerQuickOpen || customerPickerOpen || confirmOpen || Boolean(saleModal);
    const previousBodyOverflow = document.body.style.overflow;
    const previousHtmlOverflow = document.documentElement.style.overflow;
    if (shouldLock) {
      document.body.style.overflow = 'hidden';
      document.documentElement.style.overflow = 'hidden';
      document.body.classList.add('clomar-modal-open');
    }
    return () => {
      document.body.style.overflow = previousBodyOverflow;
      document.documentElement.style.overflow = previousHtmlOverflow;
      document.body.classList.remove('clomar-modal-open');
    };
  }, [mobileCartOpen, customerQuickOpen, customerPickerOpen, confirmOpen, saleModal]);

  function clearLastTicketBackup() {
    try { sessionStorage.removeItem('clomar_last_completed_sale'); } catch (err) { /* no-op */ }
  }

  function openCompletedSale(ticket) {
    if (!ticket?.sale) return;
    setLastTicket(ticket);
    setDismissedTicketId(null);
    try { sessionStorage.setItem('clomar_last_completed_sale', JSON.stringify({ ...ticket, saved_at: Date.now() })); } catch (err) {}
    setSaleModal(null);
    setTimeout(() => setSaleModal(ticket), 80);
  }

  useEffect(() => {
    try {
      const raw = sessionStorage.getItem('clomar_last_completed_sale');
      if (!raw) return;
      const saved = JSON.parse(raw);
      const fresh = Date.now() - asNum(saved.saved_at) < 15 * 60 * 1000;
      if (saved?.sale && fresh) { setLastTicket(saved); setDismissedTicketId(null); setTimeout(() => setSaleModal(saved), 120); }
      else sessionStorage.removeItem('clomar_last_completed_sale');
    } catch (err) { clearLastTicketBackup(); }
  }, []);

  useEffect(() => {
    const id = lastTicket?.sale?.id;
    if (!id || saleModal?.sale || dismissedTicketId === id) return;
    const timer = setTimeout(() => setSaleModal(lastTicket), 150);
    return () => clearTimeout(timer);
  }, [lastTicket, saleModal, dismissedTicketId]);

  function findProductByBarcode(value) {
    const clean = String(value || '').trim().toLowerCase();
    if (!clean) return null;
    return activeProducts.find(p => String(p.barcode || '').trim().toLowerCase() === clean || String(p.code || '').trim().toLowerCase() === clean) || null;
  }

  function toggleFavorite(productId) {
    setFavoriteIds(prev => prev.includes(productId) ? prev.filter(id => id !== productId) : [productId, ...prev].slice(0, 60));
  }

  function rememberProduct(productId) {
    setRecentIds(prev => [productId, ...prev.filter(id => id !== productId)].slice(0, 24));
  }

  function addProduct(product) {
    if (asNum(product.stock) <= 0) return setNotice({ type: 'warning', icon: '📦', title: 'Producto sin stock disponible', message: `${product.name} no tiene stock para vender. Ingresa a Inventario o Ingreso de mercadería para actualizar existencias.` });
    if (asNum(product.price) <= 0) return setNotice({ type: 'warning', icon: '💰', title: 'Producto sin precio de venta', message: `${product.name} tiene precio 0. Valida costo y precio antes de vender.` });
    if (productPriceStatus(product) !== 'Validado') return setNotice({ type: 'warning', icon: '🔎', title: 'Precio pendiente de validar', message: `${product.name} todavía está marcado como ${productPriceStatus(product)}. Ingresa a Precios, confirma costo y precio, y marca el producto como Validado.` });
    setLastTicket(null);
    rememberProduct(product.id);
    setCart(prev => {
      const found = prev.find(x => x.id === product.id);
      if (found) {
        if (asNum(found.qty) + 1 > asNum(product.stock)) return prev;
        return prev.map(x => x.id === product.id ? { ...x, qty: asNum(x.qty) + 1 } : x);
      }
      return [...prev, { ...product, qty: 1, discount: 0 }];
    });
  }

  function processBarcode(value, source = 'manual') {
    const clean = String(value || '').trim();
    if (!clean) return;
    const product = findProductByBarcode(clean);
    if (product) { addProduct(product); setQuery(''); setScanStatus(`Agregado al carrito: ${product.name}`); if (source === 'camera') setScanOpen(false); return; }
    setQuery(clean);
    setScanStatus(`Código no encontrado: ${clean}`);
    if (source !== 'camera') setNotice({ type: 'info', icon: '🔍', title: 'Código no encontrado', message: `No existe un producto con el código ${clean}. Puedes asignarlo desde Productos o importarlo desde Excel.` });
  }

  function handleSearchKeyDown(e) {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const clean = query.trim();
    if (!clean) return;
    const exact = findProductByBarcode(clean);
    if (exact) { processBarcode(clean, 'lector'); return; }
    if (matches.length === 1) { addProduct(matches[0]); setQuery(''); setScanStatus(`Agregado al carrito: ${matches[0].name}`); return; }
    setNotice({ type: 'info', icon: '🔍', title: 'Sin coincidencia exacta', message: 'No se encontró un producto exacto. Escanea el código de barras, escribe el código interno o busca por nombre.' });
  }

  function stopScanner() {
    if (streamRef.current) { streamRef.current.getTracks().forEach(track => track.stop()); streamRef.current = null; }
    setScanOpen(false);
  }

  useEffect(() => {
    if (!scanOpen) return;
    let cancelled = false;
    let raf = 0;
    async function startScanner() {
      try {
        if (!navigator.mediaDevices?.getUserMedia) { setScanStatus('Este navegador no permite usar la cámara. Usa lector físico o escribe el código.'); return; }
        if (!('BarcodeDetector' in window)) { setScanStatus('Tu navegador no tiene lector de código por cámara. Usa Chrome/Android, lector físico USB/Bluetooth o escribe el código manualmente.'); return; }
        setScanStatus('Solicitando permiso de cámara...');
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } }, audio: false });
        if (cancelled) { stream.getTracks().forEach(track => track.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) { videoRef.current.srcObject = stream; await videoRef.current.play(); }
        const supportedFormats = window.BarcodeDetector.getSupportedFormats ? await window.BarcodeDetector.getSupportedFormats() : [];
        const preferred = ['ean_13','ean_8','code_128','code_39','upc_a','upc_e','itf','qr_code'];
        const formats = supportedFormats.length ? preferred.filter(f => supportedFormats.includes(f)) : preferred;
        const detector = new window.BarcodeDetector({ formats: formats.length ? formats : undefined });
        setScanStatus('Apunta la cámara al código de barras.');
        async function loop() {
          if (cancelled || !videoRef.current) return;
          try { const codes = await detector.detect(videoRef.current); if (codes?.length) { const raw = codes[0].rawValue || codes[0].rawData; if (raw) { processBarcode(raw, 'camera'); return; } } } catch (err) {}
          raf = requestAnimationFrame(loop);
        }
        loop();
      } catch (err) { setScanStatus(`No se pudo abrir la cámara: ${err.message || err}`); }
    }
    startScanner();
    return () => { cancelled = true; if (raf) cancelAnimationFrame(raf); if (streamRef.current) { streamRef.current.getTracks().forEach(track => track.stop()); streamRef.current = null; } };
  }, [scanOpen]);

  function updateQty(id, qty) { setCart(prev => prev.map(x => x.id === id ? { ...x, qty: Math.max(1, Math.min(asNum(x.stock), asNum(qty || 1))) } : x)); }
  function updateItemDiscount(id, value) { setCart(prev => prev.map(x => x.id === id ? { ...x, discount: Math.max(0, Math.min(lineBase(x), asNum(value))) } : x)); }
  function removeItem(id) { setCart(prev => prev.filter(x => x.id !== id)); }
  function clearCart() {
    if (!cart.length) return;
    const accepted = window.confirm('¿Vaciar el carrito actual?');
    if (!accepted) return;
    setCart([]);
    setGlobalDiscount('0');
    setShowItemDiscounts(false);
    setShowGlobalDiscount(false);
    setMobileCartOpen(false);
    setNotice({ type:'info', icon:'🛒', title:'Carrito vaciado', message:'Puedes seleccionar nuevos productos para iniciar otra venta.' });
  }
  function setMixed(methodName, value) { setMixedPayments(prev => ({ ...prev, [methodName]: value })); }
  function fillMixed(methodName) { setMixed(methodName, Math.max(0, total - (mixedTotal - asNum(mixedPayments[methodName]))).toFixed(2)); }

  function changeDocumentType(type) {
    setDocumentType(type);
    if (type === 'Factura') setCustomerDocType('RUC');
    if (type === 'Interno') setCustomerDocNumber('');
  }

  function selectCheckoutCustomer(item) {
    const doc = String(item?.document || '');
    const docType = item?.document_type || inferDocumentType(doc);
    setCustomer(item?.name || 'Consumidor final');
    setSelectedCustomerId(item?.id || '');
    if (documentType !== 'Interno' || doc) {
      setCustomerDocType(documentType === 'Factura' ? 'RUC' : docType);
      setCustomerDocNumber(doc);
    }
    setCustomerQuery('');
    setCustomerPickerOpen(false);
  }

  function openQuickCustomer() {
    setCustomerPickerOpen(false);
    setQuickCustomer({
      name: isDefaultCheckoutCustomer ? '' : customer,
      document_type: documentType === 'Factura' ? 'RUC' : customerDocType || 'DNI',
      document: customerDocNumber || '',
      phone: '',
      address: '',
    });
    setCustomerQuickOpen(true);
  }

  async function saveQuickCustomer(e) {
    e?.preventDefault?.();
    const name = String(quickCustomer.name || '').trim();
    const doc = cleanDocument(quickCustomer.document);
    const docType = documentType === 'Factura' ? 'RUC' : quickCustomer.document_type;
    if (!name) { setNotice({ type: 'warning', icon: '👤', title: 'Falta el cliente', message: docType === 'RUC' ? 'Ingresa la razón social del cliente.' : 'Ingresa el nombre del cliente.' }); return; }
    if (docType === 'RUC' && doc.length !== 11) { setNotice({ type: 'warning', icon: '🧾', title: 'RUC inválido', message: 'El RUC debe tener 11 dígitos para registrar una factura o cliente fiscal.' }); return; }
    if (docType === 'DNI' && doc && doc.length !== 8) { setNotice({ type: 'warning', icon: '🪪', title: 'DNI inválido', message: 'El DNI debe tener 8 dígitos o puedes dejarlo vacío en una venta interna.' }); return; }
    const existing = doc ? checkoutCustomers.find(c => String(c.document || '').replace(/\D/g, '') === doc.replace(/\D/g, '')) : null;
    if (existing) { selectCheckoutCustomer(existing); setCustomerQuickOpen(false); setNotice({ type: 'info', icon: '👤', title: 'Cliente ya registrado', message: 'Se seleccionó el cliente existente con ese documento.' }); return; }
    const localCustomer = { id: `quick-${Date.now()}`, name, document: doc, document_type: docType, phone: quickCustomer.phone || '', address: quickCustomer.address || '' };
    if (!hasSupabaseConfig) {
      setCheckoutCustomers(prev => [localCustomer, ...prev]);
      selectCheckoutCustomer(localCustomer);
      setCustomerQuickOpen(false);
      return;
    }
    setQuickCustomerSaving(true);
    const payload = { name, phone: quickCustomer.phone || '', document: doc, address: quickCustomer.address || '', credit_limit: 0, status: 'Activo', store_id: profile?.store_id || DEFAULT_STORE_ID, created_by: profile?.id || null };
    const { data, error } = await supabase.from('customers').insert(payload).select().single();
    setQuickCustomerSaving(false);
    if (error) { setNotice({ type: 'warning', icon: '⚠️', title: 'No se pudo guardar el cliente', message: error.message }); return; }
    const saved = { ...data, document_type: docType };
    setCheckoutCustomers(prev => [saved, ...prev]);
    selectCheckoutCustomer(saved);
    setCustomerQuickOpen(false);
    setNotice({ type: 'success', icon: '✓', title: 'Cliente agregado', message: `${name} quedó disponible para esta venta y futuras compras.` });
  }

  function validateSaleBeforeConfirm() {
    if (!cart.length || saving) return false;
    const invalidPrice = cart.find(item => asNum(item.price) <= 0 || productPriceStatus(item) !== 'Validado');
    if (invalidPrice) { setNotice({ type: 'warning', icon: '💰', title: 'No se puede cobrar todavía', message: `Revisa y valida el precio de ${invalidPrice.name} antes de finalizar la venta.` }); return false; }
    if (total <= 0) { setNotice({ type: 'warning', icon: '📉', title: 'Total inválido', message: 'El total de la venta debe ser mayor a cero.' }); return false; }
    if (method === 'Mixto' && !paymentOk) { setNotice({ type: 'warning', icon: '💳', title: 'Pago mixto incompleto', message: `Falta cuadrar ${money(Math.abs(mixedBalance))}. El total de pagos debe coincidir con el total de la venta.` }); return false; }
    const doc = cleanDocument(customerDocNumber);
    if (documentType === 'Factura') {
      if (customerDocType !== 'RUC' || doc.length !== 11 || isDefaultCheckoutCustomer) {
        setNotice({ type: 'warning', icon: '🧾', title: 'Factura incompleta', message: 'Para factura registra RUC de 11 dígitos y razón social del cliente desde “Nuevo cliente”.' });
        return false;
      }
    }
    if (documentType === 'Boleta' && doc && ![8, 11].includes(doc.replace(/\D/g, '').length)) { setNotice({ type: 'warning', icon: '🧾', title: 'Documento incompleto', message: 'Revisa el DNI/RUC del cliente o deja la boleta sin documento.' }); return false; }
    return true;
  }

  async function checkout() {
    if (!validateSaleBeforeConfirm()) return;
    if (!cashSession?.session?.id) {
      setNotice({
        type: 'warning', icon: '🔒', title: 'Caja cerrada',
        message: 'Antes de vender debes abrir la caja desde el módulo Caja. La venta no se registró ni modificó el stock.'
      });
      return;
    }
    document.activeElement?.blur?.();
    setConfirmOpen(false);
    setMobileCartOpen(false);
    if (!hasSupabaseConfig) { alert('Venta demo registrada. Configura Supabase para guardar.'); setCart([]); return; }
    setSaving(true);

    const lineItems = cart.map(item => {
      const qty = asNum(item.qty);
      const afterLine = lineSubtotal(item);
      const globalShare = afterItemDiscount > 0 ? saleDiscount * (afterLine / afterItemDiscount) : 0;
      const finalSubtotal = Math.max(0, afterLine - globalShare);
      const finalUnitPrice = qty > 0 ? finalSubtotal / qty : 0;
      return { product_id: item.id, qty, final_unit_price: finalUnitPrice };
    });
    const mixed = method === 'Mixto'
      ? Object.entries(mixedPayments).filter(([, amount]) => asNum(amount) > 0).map(([payment_method, amount]) => ({ payment_method, amount: asNum(amount) }))
      : [];
    const saleMethod = method === 'Mixto' ? 'Mixto' : method;
    const rpcPayload = {
      p_store_id: profile?.store_id || DEFAULT_STORE_ID,
      p_customer_id: selectedCustomerId && !String(selectedCustomerId).startsWith('quick-') ? selectedCustomerId : null,
      p_customer_name: customer || 'Consumidor final',
      p_payment_method: saleMethod,
      p_document_type: documentType,
      p_items: lineItems,
      p_mixed_payments: mixed,
      p_due_date: null,
      p_customer_doc_type: documentType === 'Interno' ? null : customerDocType,
      p_customer_doc_number: documentType === 'Interno' ? null : cleanDocument(customerDocNumber),
      p_discount_total: itemDiscountTotal + saleDiscount,
      p_note: null,
    };

    const { data, error } = await supabase.rpc('clomar_register_sale_r3', rpcPayload);
    if (error) {
      console.error('Venta transaccional R3 rechazada:', error);
      setNotice({
        type: 'warning', icon: '⚠️', title: 'No se completó la venta',
        message: `${error.message || 'La operación fue rechazada.'} No se registró una venta parcial: revisa caja abierta, stock, precio y permisos.`,
      });
      setSaving(false);
      return;
    }

    const sale = data?.sale || data;
    const savedItems = Array.isArray(data?.items) && data.items.length
      ? data.items
      : cart.map(item => ({ product_id: item.id, product_name: item.name, qty: asNum(item.qty), price: asNum(item.price), subtotal: lineSubtotal(item), unit_cost: asNum(item.cost) }));
    const completedSale = {
      sale: {
        ...sale,
        payment_method: sale?.payment_method || saleMethod,
        total: asNum(sale?.total || total),
        document_type: sale?.document_type || documentType,
        sunat_status: sale?.sunat_status || (documentType === 'Interno' ? 'Interno' : 'Pre-emisión'),
        fiscal_series: sale?.fiscal_series || fiscalMeta.series,
        customer_doc_type: sale?.customer_doc_type || customerDocType,
        customer_doc_number: sale?.customer_doc_number || cleanDocument(customerDocNumber),
      },
      items: savedItems,
    };
    openCompletedSale(completedSale);
    setNotice({
      type: 'success', icon: '✅', title: `Venta ${receiptNumber(completedSale.sale)} registrada`,
      message: `Caja, inventario, reporte y auditoría se actualizaron en una sola operación por ${money(completedSale.sale.total)}.`,
    });
    setCart([]); setCustomer('Consumidor final'); setSelectedCustomerId(''); setCustomerDocNumber(''); setCustomerDocType('DNI'); setMethod('Efectivo'); setDocumentType('Interno'); setGlobalDiscount('0'); setShowItemDiscounts(false); setShowGlobalDiscount(false); setShowMorePaymentMethods(false); setMixedPayments({ Efectivo: '', Yape: '', Plin: '', Transferencia: '', Tarjeta: '' });
    await Promise.all([reloadProducts(), cashSession.reload?.()]);
    setSaving(false);
    setTimeout(() => searchInputRef.current?.focus(), 250);
  }

  const submitCheckout = () => {
    if (!cart.length) { setNotice({ type: 'info', icon: '🛒', title: 'Carrito vacío', message: 'Agrega al menos un producto para continuar con la venta.' }); return; }
    if (validateSaleBeforeConfirm()) { document.activeElement?.blur?.(); setMobileCartOpen(false); setConfirmOpen(true); }
  };

  const handleCheckoutPrimary = () => {
    if (!cart.length) return submitCheckout();
    if (typeof window !== 'undefined' && window.innerWidth <= 720 && mobileCheckoutStep === 'items') {
      setMobileCheckoutStep('payment');
      return;
    }
    submitCheckout();
  };

  return (
    <div className="page pos-page pos-pro-page">
      <FriendlyNotice notice={notice} onClose={()=>setNotice(null)} />
      <CustomerQuickModal
        open={customerQuickOpen}
        onClose={() => setCustomerQuickOpen(false)}
        form={quickCustomer}
        setForm={setQuickCustomer}
        saving={quickCustomerSaving}
        forceRuc={documentType === 'Factura'}
        onSave={saveQuickCustomer}
      />
      <CheckoutCustomerPickerModal
        open={customerPickerOpen}
        onClose={() => setCustomerPickerOpen(false)}
        query={customerQuery}
        setQuery={setCustomerQuery}
        matches={customerMatches}
        onSelect={selectCheckoutCustomer}
        onCreate={openQuickCustomer}
      />
      <SaleConfirmModal open={confirmOpen} onClose={()=>setConfirmOpen(false)} onConfirm={checkout} saving={saving} subtotal={subtotal} itemDiscountTotal={itemDiscountTotal} saleDiscount={saleDiscount} total={total} method={method} mixedPayments={mixedPayments} mixedTotal={mixedTotal} customer={customer} cart={cart} documentType={documentType} customerDocType={customerDocType} customerDocNumber={customerDocNumber} />
      <SaleCompleteModal ticket={saleModal} store={store} profile={profile} onClose={() => { setDismissedTicketId(saleModal?.sale?.id || null); clearLastTicketBackup(); setSaleModal(null); setTimeout(() => searchInputRef.current?.focus(), 100); }} onNewSale={() => { setDismissedTicketId(saleModal?.sale?.id || null); clearLastTicketBackup(); setSaleModal(null); setQuery(''); setScanStatus(''); setTimeout(() => searchInputRef.current?.focus(), 100); }} onGoReceipts={() => { setDismissedTicketId(saleModal?.sale?.id || null); clearLastTicketBackup(); setSaleModal(null); onGoReceipts?.(); }} />
      <div className="hero compact-hero pos-operation-hero"><div><span className="eyebrow">Punto de venta</span><h1>Nueva venta</h1><p>Busque, escanee y cobre. El inventario, caja y reportes se actualizan al confirmar.</p></div><span className={cashSession?.session ? 'pos-cash-status ready' : 'pos-cash-status blocked'}>{cashSession?.session ? 'Caja abierta' : 'Abra caja para vender'}</span></div>
      {lastTicket && <LastReceiptBanner ticket={lastTicket} store={store} profile={profile} onOpen={() => { setDismissedTicketId(null); setSaleModal(lastTicket); }} onGoReceipts={onGoReceipts} onDismiss={() => { setDismissedTicketId(lastTicket?.sale?.id || null); clearLastTicketBackup(); setLastTicket(null); setSaleModal(null); }} />}
      <div className="pos-layout">
        <section className="card compact-card">
          <div className="barcode-tools"><div className="search-box barcode-search"><Search size={18}/><input ref={searchInputRef} value={query} onChange={(e)=>setQuery(e.target.value)} onKeyDown={handleSearchKeyDown} placeholder="Buscar o escanear código de barras..." autoFocus /></div><button className="secondary-btn scan-btn" type="button" onClick={()=>setScanOpen(true)}>📷 Escanear con celular</button></div>
          <div className="scanner-help">Lector físico: enfoca el buscador y escanea. Cámara: abre el escáner y apunta al código.</div>
          {scanStatus && <div className="scan-status">{scanStatus}</div>}
          {scanOpen && <div className="scanner-panel"><div className="scanner-head"><strong>Escáner con cámara</strong><button className="icon-btn" type="button" onClick={stopScanner}>×</button></div><div className="scanner-frame"><video ref={videoRef} muted playsInline /></div><p className="muted">Usa la cámara trasera del celular. Si no detecta, escribe el código manualmente en el buscador.</p></div>}
          <div className="pos-discovery-toolbar">
            <div className="pos-view-tabs" aria-label="Vista de productos"><button type="button" className={productView === 'Todos' ? 'active' : ''} onClick={() => setProductView('Todos')}>Todos</button><button type="button" className={productView === 'Favoritos' ? 'active' : ''} onClick={() => setProductView('Favoritos')}>Favoritos</button><button type="button" className={productView === 'Recientes' ? 'active' : ''} onClick={() => setProductView('Recientes')}>Recientes</button></div>
            <div className="pos-category-pills"><button type="button" className={categoryFilter === 'Todas' ? 'active' : ''} onClick={() => setCategoryFilter('Todas')}>Todas</button>{productCategories.map(category => <button type="button" key={category} className={categoryFilter === category ? 'active' : ''} onClick={() => setCategoryFilter(category)}>{category}</button>)}</div>
          </div>
          <div className="pos-results-meta"><span>{matches.length} producto(s)</span><small>{productView === 'Favoritos' ? 'Accesos guardados por este usuario' : productView === 'Recientes' ? 'Últimos productos agregados' : 'Toque un producto para agregarlo al carrito'}</small></div>
          <div className="product-list premium-product-list">{matches.map(product => {
            const cartItem = cart.find(item => item.id === product.id);
            const stockState = asNum(product.stock) <= 0 ? 'agotado' : asNum(product.stock) <= asNum(product.stock_min || 1) ? 'bajo' : 'disponible';
            const stockLabel = stockState === 'agotado' ? 'Agotado' : stockState === 'bajo' ? 'Últimas unidades' : 'Disponible';
            const favorite = favoriteIds.includes(product.id);
            return <button key={product.id} type="button" className={`product-row product-row-media premium-product-row ${cartItem ? 'product-in-cart' : ''}`} onClick={() => addProduct(product)} aria-label={`Agregar ${product.name} al carrito`}>
              <img className="product-thumb" src={productImageSrc(product)} alt={product.name} />
              <div className="product-row-info"><div className="product-row-title"><strong>{product.name}</strong><span role="button" tabIndex="0" className={`favorite-toggle ${favorite ? 'active' : ''}`} onClick={(event) => { event.stopPropagation(); toggleFavorite(product.id); }} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); toggleFavorite(product.id); } }} aria-label={favorite ? 'Quitar de favoritos' : 'Agregar a favoritos'} title={favorite ? 'Quitar de favoritos' : 'Agregar a favoritos'}>{favorite ? '★' : '☆'}</span></div><small>{product.code || 'Sin código'} · {product.category || 'General'}{product.color ? ` · ${product.color}` : ''}{product.size ? ` · Talla ${product.size}` : ''}</small><div className="product-row-pills"><span className={`pos-stock-pill ${stockState}`}>{stockLabel}</span><span className={priceBadgeClass(productPriceStatus(product))}>{productPriceStatus(product)}</span>{cartItem && <em className="product-cart-badge">En carrito · {cartItem.qty}</em>}</div></div>
              <div className="product-price-column"><b>{money(product.price)}</b><small>Agregar +</small></div>
            </button>;
          })}{!matches.length && <div className="pos-empty-products"><strong>{productView === 'Favoritos' ? 'Aún no tiene favoritos' : productView === 'Recientes' ? 'Aún no hay productos recientes' : 'No se encontraron productos'}</strong><span>{productView === 'Favoritos' ? 'Use la estrella de un producto para crear su acceso rápido.' : productView === 'Recientes' ? 'Los productos que agregue aparecerán aquí.' : 'Cambie el filtro o use otro término de búsqueda.'}</span></div>}</div>
        </section>
        <aside className={`card compact-card cart-card pro-cart-card cart-mobile-sheet checkout-step-${mobileCheckoutStep} ${mobileCartOpen ? 'mobile-sheet-open' : ''}`}>
          <button className="sheet-close-btn cart-sheet-close" type="button" aria-label="Cerrar carrito" title="Cerrar carrito" onClick={() => setMobileCartOpen(false)}>×</button>
          <h3><ShoppingCart size={20}/> Carrito</h3>
          <div className="mobile-checkout-steps" aria-label="Pasos de venta">
            <button type="button" className={mobileCheckoutStep === 'items' ? 'active' : ''} onClick={() => setMobileCheckoutStep('items')}><span>1</span> Productos <b>{cart.length}</b></button>
            <button type="button" className={mobileCheckoutStep === 'payment' ? 'active' : ''} onClick={() => cart.length && setMobileCheckoutStep('payment')} disabled={!cart.length}><span>2</span> Cliente y pago</button>
          </div>
          <div className="checkout-sheet-scroll">
            {cart.length === 0 ? <div className="empty-checkout-state"><strong>Aún no hay productos</strong><span>Busca, escanea o toca un producto para armar la venta.</span><button type="button" className="secondary-btn" onClick={() => setMobileCartOpen(false)}>Agregar productos</button></div> : cart.map(item => (
              <article className="cart-item-premium" key={item.id}>
                <div className="cart-item-head"><div><strong>{item.name}</strong><small>{money(item.price)} c/u · Stock {asNum(item.stock)}</small></div><button type="button" className="cart-remove-btn" aria-label={`Quitar ${item.name}`} title="Quitar producto" onClick={()=>removeItem(item.id)}>×</button></div>
                <div className="cart-item-controls"><div className="cart-control-field cart-qty-field"><span>Cant.</span><div className="quantity-stepper"><button type="button" aria-label="Restar unidad" onClick={()=>updateQty(item.id, asNum(item.qty)-1)}>−</button><input type="number" value={item.qty} min="1" max={asNum(item.stock)} onChange={(e)=>updateQty(item.id, e.target.value)} /><button type="button" aria-label="Sumar unidad" onClick={()=>updateQty(item.id, asNum(item.qty)+1)}>+</button></div></div>{showItemDiscounts || asNum(item.discount) > 0 ? <label className="cart-control-field cart-discount-field">Desc.<input value={item.discount || ''} inputMode="decimal" onChange={(e)=>updateItemDiscount(item.id, e.target.value)} placeholder="0.00" /></label> : <button type="button" className="add-line-discount" onClick={()=>setShowItemDiscounts(true)}>Descuento ítem</button>}<div className="cart-item-total"><span>Importe</span><strong>{money(lineSubtotal(item))}</strong></div></div>
              </article>
            ))}
            <div className="sale-total-panel"><div><span>Subtotal</span><strong>{money(subtotal)}</strong></div><div><span>Desc. productos</span><strong>{money(itemDiscountTotal)}</strong></div>{showGlobalDiscount || saleDiscount > 0 ? <div className="global-discount-row"><div className="global-discount-copy"><span>Desc. general</span><button className="discount-clear-btn" type="button" onClick={()=>{ setGlobalDiscount('0'); setShowGlobalDiscount(false); }}>Quitar</button></div><input aria-label="Descuento general" value={globalDiscount} inputMode="decimal" onChange={e=>setGlobalDiscount(e.target.value)} /></div> : <button type="button" className="add-global-discount" onClick={()=>setShowGlobalDiscount(true)}>Descuento general</button>}<div className="final-total"><span>Total a cobrar</span><strong>{money(total)}</strong></div></div>
            <div className="sunat-ready-card checkout-fiscal-card checkout-fiscal-compact">
              <div className="sunat-card-head"><div><span className="eyebrow">Comprobante</span><strong>{fiscalMeta.label}</strong></div><span className={`sunat-status-pill ${sunatStatusClass(fiscalMeta.status)}`}>{fiscalMeta.status}</span></div>
              <div className="document-type-tabs">{DOCUMENT_TYPES.map(type => <button key={type} type="button" className={documentType === type ? 'active' : ''} onClick={()=>changeDocumentType(type)}>{type}</button>)}</div>
              <div className="fiscal-compact-line"><span>{documentType === 'Interno' ? 'No se envía a SUNAT' : `Serie ${fiscalMeta.series} · pre-emisión`}</span><button type="button" className="fiscal-info-link" aria-label={`Información sobre ${fiscalMeta.label}`} title="Ver información" onClick={()=>setNotice({ type:'info', icon:'🧾', title:fiscalMeta.label, message:fiscalMeta.note })}>ⓘ</button></div>
            </div>
            <section className="checkout-customer-card customer-compact-card checkout-customer-summary">
              <div className="checkout-customer-head"><div><span className="eyebrow">Cliente</span><strong>{isDefaultCheckoutCustomer ? 'Consumidor final' : customer}</strong><small>{customerDocNumber ? `${customerDocType}: ${customerDocNumber}` : 'Venta rápida sin documento'}</small></div><div className="customer-compact-actions"><button type="button" className="secondary-btn mini-change-customer" onClick={()=>setCustomerPickerOpen(true)}>Cambiar</button><button type="button" className="mini-add-customer" onClick={openQuickCustomer}>+ Nuevo</button></div></div>
            </section>
            {documentType !== 'Interno' && <section className="fiscal-client-panel fiscal-client-compact"><div className="fiscal-client-title"><div><span className="eyebrow">Datos fiscales</span><strong>{documentType === 'Factura' ? 'RUC y razón social' : 'Documento opcional para boleta'}</strong></div><button type="button" className="fiscal-edit-btn" onClick={openQuickCustomer}>Editar</button></div><div className="sunat-client-grid"><select value={documentType === 'Factura' ? 'RUC' : customerDocType} disabled={documentType === 'Factura'} onChange={e=>setCustomerDocType(e.target.value)}><option>DNI</option><option>RUC</option><option>CE</option><option>Sin documento</option></select><input value={customerDocNumber} inputMode="numeric" onChange={e=>setCustomerDocNumber(cleanDocument(e.target.value))} placeholder={documentType === 'Factura' ? 'RUC de 11 dígitos' : 'Número opcional'} /></div></section>}
            <div className="payment-section"><span className="eyebrow payment-label">Método de pago</span><div className="payment-fast-row"><button type="button" className={method==='Efectivo'?'active':''} onClick={()=>{setMethod('Efectivo');setShowMorePaymentMethods(false);}}>Efectivo</button><button type="button" className={method==='Yape'?'active':''} onClick={()=>{setMethod('Yape');setShowMorePaymentMethods(false);}}>Yape</button><button type="button" className={method==='Plin'?'active':''} onClick={()=>{setMethod('Plin');setShowMorePaymentMethods(false);}}>Plin</button><button type="button" className={method==='Mixto'?'active':''} onClick={()=>{setMethod('Mixto');setShowMorePaymentMethods(false);}}>Mixto</button></div><button type="button" className="other-payment-toggle" onClick={()=>setShowMorePaymentMethods(v=>!v)}>{showMorePaymentMethods ? 'Ocultar otros métodos' : 'Otros métodos'}</button>{showMorePaymentMethods && <div className="additional-payment-row"><button type="button" className={method==='Tarjeta'?'active':''} onClick={()=>setMethod('Tarjeta')}>Tarjeta</button><button type="button" className={method==='Transferencia'?'active':''} onClick={()=>setMethod('Transferencia')}>Transferencia</button><button type="button" className={method==='Crédito'?'active':''} onClick={()=>setMethod('Crédito')}>Crédito</button></div>}</div>
            {method === 'Mixto' && <div className="mixed-payment-box"><h4>Pago mixto</h4>{Object.keys(mixedPayments).map(pay => <div className="mixed-row" key={pay}><span>{pay}</span><input value={mixedPayments[pay]} inputMode="decimal" onChange={e=>setMixed(pay, e.target.value)} placeholder="0.00" /><button type="button" onClick={()=>fillMixed(pay)}>Completar</button></div>)}<div className={paymentOk ? 'mixed-ok' : 'mixed-pending'}>{paymentOk ? 'Pagos cuadrados' : `Falta/cuadra: ${money(Math.abs(mixedBalance))}`}</div></div>}
          </div>
          <footer className="checkout-footer" aria-label="Resumen final de la venta">
            <div className="checkout-footer-action">
              <div className="checkout-total-dock" aria-label={`Total ${money(total)}`}>
                <span>Total</span>
                <strong>{money(total)}</strong>
              </div>
              <button className="primary-btn checkout-submit-btn" disabled={saving} onClick={handleCheckoutPrimary}><span className="checkout-label-desktop">{checkoutButtonLabel}</span><span className="checkout-label-mobile">{mobileCheckoutStep === 'items' ? 'Continuar' : checkoutButtonLabel}</span></button>
            </div>
          </footer>
        </aside>
      </div>
      {cart.length > 0 && !mobileCartOpen && !menuOpen && !customerQuickOpen && !customerPickerOpen && !confirmOpen && !saleModal && <div className="mobile-checkout-bar" role="status" aria-label={`${cart.length} producto(s) en el carrito por ${money(total)}`}>
        <div className="mobile-cart-summary"><small>Carrito activo</small><strong>{cart.length} producto(s) · {money(total)}</strong></div>
        <button type="button" className="mobile-cart-clear" onClick={clearCart}>Vaciar</button>
        <button type="button" className="primary-btn mobile-cart-open-btn" onClick={() => { setMobileCheckoutStep('items'); setMobileCartOpen(true); }}>Ver carrito</button>
      </div>}
    </div>
  );
}

function CheckoutCustomerPickerModal({ open, onClose, query, setQuery, matches = [], onSelect, onCreate }) {
  if (!open) return null;
  return (
    <div className="customer-modal-backdrop customer-picker-backdrop" role="dialog" aria-modal="true" onMouseDown={(event)=>{ if (event.target === event.currentTarget) onClose(); }}>
      <section className="customer-modal-card customer-picker-card">
        <div className="customer-modal-handle" />
        <div className="customer-modal-head"><div><span className="eyebrow">Cliente de la venta</span><h3>Buscar o seleccionar</h3><p>Busca por nombre, DNI, RUC o teléfono. Selecciona un cliente para usarlo en esta venta.</p></div><button type="button" className="sheet-x-btn" onClick={onClose}>×</button></div>
        <div className="customer-search-picker customer-picker-search"><Search size={18}/><input autoFocus value={query} onChange={event=>setQuery(event.target.value)} placeholder="Nombre, DNI, RUC o teléfono" />{query && <button type="button" className="clear-customer-query" onClick={()=>setQuery('')}>×</button>}</div>
        <div className="customer-picker-results">
          {matches.length ? matches.map(item => <button type="button" key={item.id} onClick={()=>onSelect(item)}><span><strong>{item.name}</strong><small>{item.document || 'Sin documento'} · {item.phone || 'Sin teléfono'}</small></span><b>Usar</b></button>) : <div className="empty-customer-picker">No se encontraron clientes con esa búsqueda.</div>}
        </div>
        <button type="button" className="create-customer-inline customer-picker-create" onClick={onCreate}>+ Registrar nuevo cliente</button>
      </section>
    </div>
  );
}

function CustomerQuickModal({ open, onClose, form, setForm, saving, forceRuc = false, onSave }) {
  if (!open) return null;
  const type = forceRuc ? 'RUC' : form.document_type;
  return (
    <div className="customer-modal-backdrop" role="dialog" aria-modal="true">
      <form className="customer-modal-card" onSubmit={onSave}>
        <div className="customer-modal-handle" />
        <div className="customer-modal-head"><div><span className="eyebrow">Cliente desde venta</span><h3>{forceRuc ? 'Nueva razón social' : 'Nuevo cliente'}</h3><p>{forceRuc ? 'Para factura se requiere RUC y razón social.' : 'Registra el cliente sin salir de la venta.'}</p></div><button type="button" className="sheet-x-btn" onClick={onClose}>×</button></div>
        <div className="quick-customer-grid">
          <label>Tipo de documento<select value={type} disabled={forceRuc} onChange={e=>setForm({...form, document_type:e.target.value})}><option>DNI</option><option>RUC</option><option>CE</option><option>Sin documento</option></select></label>
          <label>{type === 'RUC' ? 'RUC' : 'Documento'}<input value={form.document} inputMode="numeric" onChange={e=>setForm({...form, document:cleanDocument(e.target.value)})} placeholder={type === 'RUC' ? '11 dígitos' : type === 'DNI' ? '8 dígitos' : 'Opcional'} /></label>
          <label className="full-row">{type === 'RUC' ? 'Razón social' : 'Nombre completo'}<input autoFocus value={form.name} onChange={e=>setForm({...form, name:e.target.value})} placeholder={type === 'RUC' ? 'Razón social del cliente' : 'Nombres y apellidos'} /></label>
          <label>Teléfono<input value={form.phone} onChange={e=>setForm({...form, phone:e.target.value})} placeholder="Celular o WhatsApp" /></label>
          <label>Dirección<input value={form.address} onChange={e=>setForm({...form, address:e.target.value})} placeholder="Opcional" /></label>
        </div>
        <div className="customer-api-note"><strong>Consulta automática DNI/RUC</strong><span>Se conecta después desde backend seguro. Por ahora los datos se registran manualmente.</span></div>
        <div className="customer-modal-actions"><button type="button" className="secondary-btn" onClick={onClose}>Cancelar</button><button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar y usar cliente'}</button></div>
      </form>
    </div>
  );
}

function SaleConfirmModal({ open, onClose, onConfirm, saving, subtotal, itemDiscountTotal, saleDiscount, total, method, mixedPayments, mixedTotal, customer, cart, documentType = 'Interno', customerDocType = 'DNI', customerDocNumber = '' }) {
  const fiscalMeta = documentMeta(documentType);
  if (!open) return null;
  return (
    <div className="notice-backdrop" role="dialog" aria-modal="true">
      <div className="notice-card confirm-sale-card premium-confirm-sale-card">
        <div className="notice-icon">🧾</div>
        <div className="notice-content">
          <h3>{documentType === 'Interno' ? 'Confirmar venta interna' : `Confirmar ${fiscalMeta.label}`}</h3>
          <p>{documentType === 'Interno' ? 'Revisa el total y método de pago antes de registrar el comprobante interno.' : 'Esta venta se registrará como pre-emisión. Aún no se enviará a SUNAT hasta conectar un backend seguro y un PSE/OSE.'}</p>
          <div className="sunat-confirm-strip"><span>{fiscalMeta.label}</span><strong>{fiscalMeta.series} · {fiscalMeta.status}</strong></div>
          <div className="confirm-sale-summary">
            <div><span>Cliente</span><strong>{customer || 'Cliente'}</strong></div>
            <div><span>Documento</span><strong>{documentType === 'Interno' ? 'Interno' : `${customerDocType} ${customerDocNumber || 'pendiente'}`}</strong></div>
            <div><span>Productos</span><strong>{cart.length}</strong></div>
            <div><span>Subtotal</span><strong>{money(subtotal)}</strong></div>
            <div><span>Descuentos</span><strong>{money(itemDiscountTotal + saleDiscount)}</strong></div>
            <div><span>Método</span><strong>{method}</strong></div>
            <div><span>Total final</span><strong>{money(total)}</strong></div>
          </div>
          {method === 'Mixto' && <div className="info-box">Pagos mixtos: {Object.entries(mixedPayments).filter(([,v])=>asNum(v)>0).map(([k,v])=>`${k} ${money(v)}`).join(' · ')} · Total pagos {money(mixedTotal)}</div>}
          <div className="notice-actions"><button type="button" className="secondary-btn" onClick={onClose}>Cancelar</button><button type="button" className="primary-btn" disabled={saving} onClick={onConfirm}>{saving ? 'Guardando...' : documentType === 'Interno' ? 'Registrar venta' : 'Registrar pre-emisión'}</button></div>
        </div>
      </div>
    </div>
  );
}

function LastReceiptBanner({ ticket, store = {}, profile = {}, onOpen, onGoReceipts, onDismiss }) {
  if (!ticket?.sale) return null;
  const sale = ticket.sale;
  const items = ticket.items || [];
  const totalItems = items.reduce((sum, item) => sum + asNum(item.qty), 0);
  return (
    <section className="post-sale-banner">
      <div className="post-sale-main">
        <div className="post-sale-icon">✓</div>
        <div>
          <span className="eyebrow">Último comprobante generado</span>
          <h3>{receiptNumber(sale)} · {money(sale.total)}</h3>
          <p>{sale.customer_name || 'Cliente'} · {sale.payment_method || 'Efectivo'} · {totalItems} producto(s). Si la ventana automática no aparece, usa estos botones de respaldo.</p>
        </div>
      </div>
      <div className="post-sale-actions">
        <button className="primary-btn" type="button" onClick={onOpen}>Abrir comprobante</button>
        <button className="secondary-btn" type="button" onClick={() => printReceipt({ sale, items, store, profile, format: '80mm' })}>Ticket 80mm</button>
        <button className="secondary-btn" type="button" onClick={() => printReceipt({ sale, items, store, profile, format: '58mm' })}>Ticket 58mm</button>
        <button className="secondary-btn" type="button" onClick={() => printReceipt({ sale, items, store, profile, format: 'a4' })}>PDF A4</button>
        <button className="secondary-btn" type="button" onClick={onGoReceipts}>Ver historial</button>
        <button className="icon-btn" type="button" onClick={onDismiss}>×</button>
      </div>
    </section>
  );
}



function SaleCompleteModal({ ticket, store = {}, profile = {}, onClose, onNewSale, onGoReceipts }) {
  const [showPreview, setShowPreview] = useState(false);
  if (!ticket?.sale) return null;
  const sale = ticket.sale;
  const items = ticket.items || [];
  const totalItems = items.reduce((sum, item) => sum + asNum(item.qty), 0);
  const documentType = sale?.document_type || 'Interno';
  const fiscalMeta = documentMeta(documentType);
  const fiscalStatus = sale?.sunat_status || fiscalMeta.status;
  const formatPrint = (format) => printReceipt({ sale, items, store, profile, format });
  return (
    <div className="sale-modal-backdrop" role="dialog" aria-modal="true">
      <div className="sale-modal-card premium-sale-modal-card">
        <div className="sale-modal-head">
          <div className="sale-success-icon">✓</div>
          <div>
            <span className="eyebrow">{documentType === 'Interno' ? 'Comprobante generado' : 'Registro SUNAT-ready'}</span>
            <h2>{documentType === 'Interno' ? 'Venta registrada correctamente' : `${fiscalMeta.label} registrada`}</h2>
            <p>{documentType === 'Interno' ? 'El comprobante interno quedó listo para imprimir, guardar o reimprimir desde el módulo Comprobantes.' : 'El documento quedó guardado como pre-emisión. Aún no fue enviado a SUNAT ni a un proveedor electrónico.'}</p>
          </div>
          <button className="icon-btn sale-modal-close neutral-close" type="button" onClick={onClose}>Cerrar</button>
        </div>
        <div className="sunat-result-strip"><span className={`sunat-status-pill ${sunatStatusClass(fiscalStatus)}`}>{fiscalStatus}</span><strong>{documentType === 'Interno' ? 'Control interno' : `${sale?.fiscal_series || fiscalMeta.series} · Pre-emisión · Envío no conectado`}</strong></div>
        <div className="sale-summary-grid">
          <div><span>Tipo</span><strong>{documentType}</strong></div>
          <div><span>N°</span><strong>{receiptNumber(sale)}</strong></div>
          <div><span>Total</span><strong>{money(sale.total)}</strong></div>
          <div><span>Método</span><strong>{sale.payment_method || 'Efectivo'}</strong></div>
          <div><span>Productos</span><strong>{totalItems}</strong></div>
        </div>
        <div className="sale-modal-body premium-sale-modal-body">
          <section className="sale-modal-actions-card premium-actions-card">
            <h3>Acción rápida</h3>
            {documentType !== 'Interno' && <button className="secondary-btn sunat-disabled-action" type="button" disabled>Enviar a SUNAT · Próximamente</button>}
            <div className="receipt-action-grid premium-action-grid">
              <button className="primary-btn" type="button" onClick={() => formatPrint('80mm')}>Imprimir 80mm</button>
              <button className="secondary-btn" type="button" onClick={() => formatPrint('58mm')}>Ticket 58mm</button>
              <button className="secondary-btn" type="button" onClick={() => formatPrint('a4')}>PDF A4</button>
              <button className="secondary-btn" type="button" onClick={() => downloadText(`comprobante-${receiptNumber(sale)}.txt`, ticketText(sale, items))}>TXT</button>
            </div>
            <div className="sale-modal-next-actions premium-next-actions">
              <button className="secondary-btn" type="button" onClick={() => setShowPreview(v => !v)}>{showPreview ? 'Ocultar comprobante' : 'Ver comprobante'}</button>
              <button className="secondary-btn" type="button" onClick={onGoReceipts}>Historial</button>
              <button className="primary-btn" type="button" onClick={onNewSale}>Nueva venta</button>
            </div>
          </section>
          {showPreview && (
            <section className="sale-modal-preview-card premium-preview-card">
              <ReceiptMiniPreview sale={sale} items={items} store={store} profile={profile} format="80mm" />
            </section>
          )}
        </div>
      </div>
    </div>
  );
}

function ReceiptMiniPreview({ sale, items = [], store = {}, profile = {}, format = '80mm' }) {
  const totals = receiptTotals(sale, items);
  return (
    <div className={`receipt-preview receipt-preview-${format}`}>
      <div className="receipt-preview-head">
        <img src={store?.logo_url || APP_ICON} alt="Logo" />
        <strong>{store?.name || 'Clomar Store'}</strong>
        {store?.ruc && <small>RUC: {store.ruc}</small>}
        {store?.address && <small>{store.address}</small>}
        {store?.phone && <small>Tel: {store.phone}</small>}
      </div>
      <div className="receipt-sep" />
      <div className="receipt-title">{sale?.payment_method === 'Crédito' ? 'COMPROBANTE DE CRÉDITO' : 'COMPROBANTE INTERNO'}</div>
      <div className="receipt-meta-row"><span>N°</span><b>{receiptNumber(sale)}</b></div>
      <div className="receipt-meta-row"><span>Fecha</span><b>{fmtDate(sale?.created_at)}</b></div>
      <div className="receipt-meta-row"><span>Cliente</span><b>{sale?.customer_name || 'Cliente'}</b></div>
      <div className="receipt-meta-row"><span>Vendedor</span><b>{profile?.full_name || sale?.seller_email || profile?.email || 'Usuario'}</b></div>
      <div className="receipt-sep" />
      <div className="receipt-items-preview">
        {items.map((it) => <div key={it.id || it.product_name} className="receipt-item-line"><span>{asNum(it.qty)} x {it.product_name}</span><b>{money(it.subtotal || asNum(it.qty) * asNum(it.price))}</b></div>)}
        {!items.length && <p className="muted">Selecciona un comprobante para ver su detalle.</p>}
      </div>
      <div className="receipt-sep" />
      <div className="receipt-total-line"><span>Total</span><strong>{money(totals.total)}</strong></div>
      <img className="receipt-qr" src={qrUrl(receiptQrData(sale, store))} alt="QR comprobante" />
      <small className="muted center-block">Comprobante interno. No reemplaza comprobante electrónico SUNAT.</small>
    </div>
  );
}

function ReceiptsPage({ profile, store }) {
  const { sales, loading, reload } = useSales(profile);
  const [selectedSale, setSelectedSale] = useState(null);
  const [items, setItems] = useState([]);
  const [format, setFormat] = useState('80mm');
  const [query, setQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('Todos');
  const [loadingItems, setLoadingItems] = useState(false);

  const filteredSales = useMemo(() => sales.filter(s => {
    const text = `${receiptNumber(s)} ${s.customer_name || ''} ${s.payment_method || ''} ${s.status || ''} ${s.document_type || ''} ${s.sunat_status || ''}`.toLowerCase();
    const passText = !query.trim() || text.includes(query.trim().toLowerCase());
    const passStatus = statusFilter === 'Todos' || s.status === statusFilter || s.payment_method === statusFilter || s.document_type === statusFilter || s.sunat_status === statusFilter;
    return passText && passStatus;
  }), [sales, query, statusFilter]);

  async function selectSale(sale) {
    setSelectedSale(sale);
    setItems([]);
    if (!hasSupabaseConfig || !sale?.id) return;
    setLoadingItems(true);
    const { data, error } = await supabase
      .from('sale_items')
      .select('*')
      .eq('sale_id', sale.id)
      .order('created_at', { ascending: true });
    if (!error) setItems(data || []);
    setLoadingItems(false);
  }

  useEffect(() => {
    if (!selectedSale && filteredSales.length) selectSale(filteredSales[0]);
  }, [filteredSales.length]);

  const today = todayISO();
  const salesToday = sales.filter(s => String(s.created_at || '').slice(0,10) === today);
  const totalToday = salesToday.reduce((sum, sale) => sum + asNum(sale.total), 0);
  const creditToday = salesToday.filter(s => s.payment_method === 'Crédito' || s.status === 'Crédito').reduce((sum, sale) => sum + asNum(sale.total), 0);

  return (
    <div className="page receipts-page">
      <div className="hero compact-hero"><h1>🧾 Comprobantes</h1><p>Reimprime tickets internos, vouchers y PDF A4 con logo, tienda, vendedor, cliente y detalle de venta.</p></div>
      <div className="kpi-grid">
        <Kpi label="Comprobantes hoy" value={salesToday.length} helper="ventas registradas" />
        <Kpi label="Total hoy" value={money(totalToday)} helper="monto vendido" />
        <Kpi label="Crédito hoy" value={money(creditToday)} helper="por cobrar" />
        <Kpi label="Historial cargado" value={sales.length} helper="últimas ventas" />
      </div>
      <div className="receipt-workspace">
        <section className="card compact-card receipt-list-card">
          <div className="card-head-line"><h3>Historial de comprobantes</h3><button className="secondary-btn" onClick={reload} disabled={loading}>{loading ? 'Cargando...' : 'Actualizar'}</button></div>
          <div className="receipt-filters">
            <label>Buscar<input value={query} onChange={e=>setQuery(e.target.value)} placeholder="B123, cliente, método..." /></label>
            <label>Estado<select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}><option>Todos</option><option>Pagado</option><option>Crédito</option><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option><option>Interno</option><option>Pre-emisión</option></select></label>
          </div>
          <div className="receipt-list">
            {filteredSales.map(sale => (
              <button key={sale.id} className={`receipt-row ${selectedSale?.id === sale.id ? 'active' : ''}`} onClick={()=>selectSale(sale)}>
                <div><strong>{receiptNumber(sale)} · {sale.customer_name || 'Cliente'}</strong><small>{fmtDate(sale.created_at)} · {sale.payment_method || 'Efectivo'} · {sale.document_type || 'Interno'} · {sale.sunat_status || sale.status || 'Pagado'}</small></div>
                <b>{money(sale.total)}</b>
              </button>
            ))}
            {!filteredSales.length && <p className="muted">No hay comprobantes con esos filtros.</p>}
          </div>
        </section>
        <section className="card compact-card receipt-detail-card">
          <div className="card-head-line"><h3>Vista e impresión</h3><span className="result-pill">{selectedSale ? receiptNumber(selectedSale) : 'Sin selección'}</span></div>
          <div className="receipt-print-controls">
            <label>Formato<select value={format} onChange={e=>setFormat(e.target.value)}><option value="80mm">Ticket 80mm</option><option value="58mm">Ticket 58mm</option><option value="a4">PDF A4</option></select></label>
            <button className="primary-btn" disabled={!selectedSale || loadingItems} onClick={()=>printReceipt({ sale: selectedSale, items, store, profile, format })}>Imprimir / Guardar PDF</button>
            <button className="secondary-btn" disabled={!selectedSale} onClick={()=>downloadText(`comprobante-${receiptNumber(selectedSale)}.txt`, ticketText(selectedSale, items))}>Descargar TXT</button>
          </div>
          {loadingItems ? <div className="loader">Cargando detalle...</div> : <ReceiptMiniPreview sale={selectedSale || { total: 0 }} items={items} store={store} profile={profile} format={format} />}
        </section>
      </div>
    </div>
  );
}


function CategoriesAdmin({ profile, categories = [], subcategories = [], products = [], reloadCategories }) {
  const [form, setForm] = useState({ name: '', type: 'principal', parent_id: '', description: '', sort_order: '100' });
  const [saving, setSaving] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [expanded, setExpanded] = useState({});
  const [detailCategory, setDetailCategory] = useState(null);

  const productCountFor = (cat) => products.filter(p => p.active !== false && (p.category_id === cat.id || p.subcategory_id === cat.id || p.category === cat.name || p.subcategory === cat.name)).length;
  const visibleCategories = categories.filter(cat => {
    const children = subcategories.filter(s => s.parent_id === cat.id);
    const text = `${cat.name} ${children.map(c => c.name).join(' ')}`.toLowerCase();
    return text.includes(query.toLowerCase());
  });

  function resetForm() {
    setForm({ name: '', type: 'principal', parent_id: '', description: '', sort_order: '100' });
  }

  async function save(e) {
    e.preventDefault();
    if (!form.name.trim()) return alert('Coloca el nombre de la categoría o subcategoría.');
    if (form.type === 'subcategoria' && !form.parent_id) return alert('Selecciona la categoría principal de esta subcategoría.');
    setSaving(true);
    const payload = {
      store_id: profile?.store_id || DEFAULT_STORE_ID,
      name: form.name.trim(),
      parent_id: form.type === 'subcategoria' ? form.parent_id : null,
      description: form.description.trim(),
      sort_order: asNum(form.sort_order || 100),
      active: true,
    };
    const { error } = await supabase.from('product_categories').insert(payload);
    setSaving(false);
    if (error) return alert(error.message || 'No se pudo guardar la categoría.');
    resetForm();
    setFormOpen(false);
    reloadCategories?.();
  }

  async function renameCategory(cat) {
    const nextName = prompt('Nuevo nombre de categoría/subcategoría:', cat.name);
    if (!nextName || !nextName.trim() || nextName.trim() === cat.name) return;
    const { error } = await supabase.from('product_categories').update({ name: nextName.trim() }).eq('id', cat.id);
    if (error) alert(error.message); else reloadCategories?.();
  }

  async function deactivateCategory(cat) {
    if (!confirm(`¿Desactivar ${cat.name}? No se borrarán los productos ya registrados.`)) return;
    const { error } = await supabase.from('product_categories').update({ active: false }).eq('id', cat.id);
    if (error) alert(error.message); else reloadCategories?.();
  }

  function toggle(id) {
    setExpanded(prev => ({ ...prev, [id]: !prev[id] }));
  }

  return (
    <div className="page categories-page">
      <div className="hero compact-hero"><h1>🏷️ Categorías</h1><p>Panel profesional para organizar productos por categoría principal y subcategoría.</p></div>
      <div className="category-toolbar card compact-card premium-summary-card">
        <div>
          <h3>Resumen de categorías</h3>
          <p className="muted">{categories.length} categorías principales · {subcategories.length} subcategorías activas · {products.filter(p=>p.active!==false).length} productos activos</p>
        </div>
        <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar categoría o subcategoría..." /></div>
      </div>
      <div className="mobile-page-actions premium-mobile-actions">
        <button type="button" className="primary-btn" onClick={() => setFormOpen(true)}>+ Nueva categoría</button>
        <span>{visibleCategories.length} visibles</span>
      </div>
      <div className="two-col category-admin-layout polished-categories">
        <form className={`card form-grid category-form-card category-form-sheet ${formOpen ? 'form-open' : ''}`} onSubmit={save}>
          <button className="sheet-close-btn form-sheet-close" type="button" onClick={() => setFormOpen(false)}>Cerrar ×</button>
          <h3>Nueva categoría</h3>
          <label>Tipo de registro
            <select value={form.type} onChange={e=>setForm({...form,type:e.target.value,parent_id:''})}>
              <option value="principal">Categoría principal</option>
              <option value="subcategoria">Subcategoría</option>
            </select>
          </label>
          {form.type === 'subcategoria' && (
            <label>Categoría padre
              <select value={form.parent_id} onChange={e=>setForm({...form,parent_id:e.target.value})}>
                <option value="">Selecciona categoría principal</option>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
          )}
          <label>{form.type === 'principal' ? 'Nombre de categoría principal' : 'Nombre de subcategoría'}
            <input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder={form.type === 'principal' ? 'Ejemplo: Tecnología' : 'Ejemplo: Correas hombre'} />
          </label>
          <label>Descripción interna
            <input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="Opcional para control interno" />
          </label>
          <label>Orden
            <input value={form.sort_order} onChange={e=>setForm({...form,sort_order:e.target.value})} inputMode="numeric" />
          </label>
          <div className="form-actions-row">
            <button type="button" className="secondary-btn" onClick={resetForm}>Limpiar</button>
            <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar'}</button>
          </div>
        </form>
        <section className="card compact-card category-panel-card">
          <div className="category-panel-head">
            <div>
              <h3>Mapa de categorías</h3>
              <p className="muted">Haz clic en “Ver” para desplegar subcategorías.</p>
            </div>
          </div>
          <div className="category-card-grid">
            {visibleCategories.map(cat => {
              const children = subcategories.filter(s => s.parent_id === cat.id);
              const catProducts = productCountFor(cat);
              const isOpen = false;
              return (
                <article className="category-card-pro" key={cat.id}>
                  <div className="category-card-main">
                    <div className="category-icon-badge">🏷️</div>
                    <div>
                      <strong>{cat.name}</strong>
                      <small>{children.length} subcategorías · {catProducts} productos</small>
                    </div>
                  </div>
                  <div className="category-card-actions">
                    <button type="button" className="secondary-btn" onClick={()=>setDetailCategory(cat)}>Ver</button>
                    <button type="button" className="secondary-btn" onClick={()=>renameCategory(cat)}>Editar</button>
                    <button type="button" className="danger-mini-btn" onClick={()=>deactivateCategory(cat)}>Desactivar</button>
                  </div>
                </article>
              );
            })}
            {!visibleCategories.length && <p className="muted">No se encontraron categorías con ese criterio.</p>}
          </div>
        </section>
      </div>
      <CategoryDetailSheet category={detailCategory} subcategories={subcategories.filter(s => s.parent_id === detailCategory?.id)} productCountFor={productCountFor} onClose={() => setDetailCategory(null)} onEdit={renameCategory} onDeactivate={deactivateCategory} />
    </div>
  );
}

function CategoryDetailSheet({ category, subcategories = [], productCountFor, onClose, onEdit, onDeactivate }) {
  if (!category) return null;
  return (
    <div className="ux-sheet-backdrop" role="dialog" aria-modal="true" onClick={onClose}>
      <section className="category-detail-sheet" onClick={e => e.stopPropagation()}>
        <div className="sheet-handle" />
        <div className="category-detail-head">
          <div className="category-icon-badge">🏷️</div>
          <div>
            <h3>{category.name}</h3>
            <p>{subcategories.length} subcategorías · {productCountFor(category)} productos</p>
          </div>
          <button type="button" className="sheet-x-btn" onClick={onClose}>×</button>
        </div>
        <div className="category-detail-actions">
          <button type="button" className="secondary-btn" onClick={() => onEdit(category)}>Editar categoría</button>
          <button type="button" className="danger-mini-btn" onClick={() => onDeactivate(category)}>Desactivar</button>
        </div>
        <h4>Subcategorías</h4>
        <div className="subcategory-sheet-list">
          {subcategories.length ? subcategories.map(sub => (
            <article className="subcategory-sheet-card" key={sub.id}>
              <div>
                <strong>{sub.name}</strong>
                <small>{productCountFor(sub)} productos</small>
              </div>
              <button type="button" className="secondary-btn" onClick={() => onEdit(sub)}>Editar</button>
              <button type="button" className="danger-mini-btn" onClick={() => onDeactivate(sub)}>Desactivar</button>
            </article>
          )) : <p className="muted empty-subcategory">Sin subcategorías registradas.</p>}
        </div>
      </section>
    </div>
  );
}

function Products({ products, reload, profile, categories = [], subcategories = [], reloadCategories }) {
  const emptyForm = { code:'', barcode:'', name:'', category_id:'', subcategory_id:'', category:'', subcategory:'', brand:'', size:'', color:'', description:'', price:'', cost:'', stock:'0', stock_min:'2', image_url:'', price_status:'Pendiente', margin_target:'50', min_price:'0', price_notes:'', public_visible:false, catalog_status:'Borrador', catalog_featured:false, catalog_description:'', catalog_position:'999' };
  const [form, setForm] = useState(emptyForm);
  const [imageFile, setImageFile] = useState(null);
  const [previewUrl, setPreviewUrl] = useState('');
  const [saving, setSaving] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [productPage, setProductPage] = useState(1);
  useEffect(() => {
    if (!formOpen) return undefined;
    const onKeyDown = (event) => { if (event.key === 'Escape') setFormOpen(false); };
    document.body.classList.add('clomar-modal-open');
    window.addEventListener('keydown', onKeyDown);
    return () => { document.body.classList.remove('clomar-modal-open'); window.removeEventListener('keydown', onKeyDown); };
  }, [formOpen]);
  const filteredProducts = products.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(query.toLowerCase()));
  const PRODUCT_PAGE_SIZE = 18;
  const productPageCount = Math.max(1, Math.ceil(filteredProducts.length / PRODUCT_PAGE_SIZE));
  const safeProductPage = Math.min(productPage, productPageCount);
  const visibleProducts = filteredProducts.slice((safeProductPage - 1) * PRODUCT_PAGE_SIZE, safeProductPage * PRODUCT_PAGE_SIZE);
  useEffect(() => { setProductPage(1); }, [query]);
  const selectedSubcategories = subcategories.filter(c => c.parent_id === form.category_id);
  function setCategoryById(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    setForm({ ...form, category_id: categoryId, category: cat?.name || '', subcategory_id: '', subcategory: '' });
  }
  function setSubcategoryById(subcategoryId) {
    const sub = subcategories.find(c => c.id === subcategoryId);
    setForm({ ...form, subcategory_id: subcategoryId, subcategory: sub?.name || '' });
  }

  useEffect(() => {
    if (!imageFile) {
      setPreviewUrl(form.image_url || '');
      return;
    }
    const localUrl = URL.createObjectURL(imageFile);
    setPreviewUrl(localUrl);
    return () => URL.revokeObjectURL(localUrl);
  }, [imageFile, form.image_url]);

  async function uploadProductImage() {
    if (!imageFile) return { image_url: form.image_url || '', image_path: '' };
    const storeId = profile?.store_id || DEFAULT_STORE_ID;
    const filePath = `${storeId}/${Date.now()}-${cleanFileName(imageFile.name)}`;
    const { error: uploadError } = await supabase.storage
      .from('product-images')
      .upload(filePath, imageFile, { cacheControl: '3600', upsert: false });
    if (uploadError) throw uploadError;
    const { data } = supabase.storage.from('product-images').getPublicUrl(filePath);
    return { image_url: data?.publicUrl || '', image_path: filePath };
  }

  async function save(e) {
    e.preventDefault();
    if (!hasSupabaseConfig) return alert('Configura Supabase para guardar productos.');
    if (!form.name.trim()) return alert('Coloca el nombre del producto.');
    if (form.price_status === 'Validado' && asNum(form.price) <= 0) return alert('No puedes validar un producto con precio 0.');
    if (form.price_status === 'Validado' && asNum(form.min_price) > 0 && asNum(form.price) < asNum(form.min_price)) return alert('El precio validado no puede ser menor al precio mínimo.');
    setSaving(true);
    try {
      const imageData = await uploadProductImage();
      const payload = {
        code: form.code.trim(),
        barcode: form.barcode.trim(),
        name: form.name.trim(),
        category: form.category.trim() || 'General',
        subcategory: form.subcategory.trim(),
        category_id: form.category_id || null,
        subcategory_id: form.subcategory_id || null,
        brand: form.brand.trim(),
        size: form.size.trim(),
        color: form.color.trim(),
        description: form.description.trim(),
        price: asNum(form.price),
        cost: asNum(form.cost),
        price_status: form.price_status || 'Pendiente',
        margin_target: asNum(form.margin_target || 50),
        min_price: asNum(form.min_price || 0),
        price_notes: form.price_notes || '',
        public_visible: Boolean(form.public_visible),
        catalog_status: form.public_visible ? (form.catalog_status || 'Borrador') : 'Borrador',
        catalog_featured: Boolean(form.catalog_featured),
        catalog_description: form.catalog_description || '',
        catalog_position: Math.max(0, Math.trunc(asNum(form.catalog_position || 999))),
        catalog_updated_at: new Date().toISOString(),
        price_updated_at: new Date().toISOString(),
        price_updated_by: profile?.id || null,
        stock: asNum(form.stock),
        stock_min: asNum(form.stock_min),
        status: 'Activo',
        active: true,
        image_url: imageData.image_url,
        image_path: imageData.image_path,
        store_id: profile?.store_id || DEFAULT_STORE_ID,
        created_by: profile?.id || null,
      };
      const { error } = await supabase.from('products').insert(payload);
      if (error) throw error;
      setForm(emptyForm);
      setImageFile(null);
      setPreviewUrl('');
      reload();
      alert('Producto guardado correctamente.');
      setFormOpen(false);
    } catch (error) {
      alert(error.message || 'No se pudo guardar el producto.');
    } finally {
      setSaving(false);
    }
  }

  async function deactivateProduct(product) {
    if (!confirm(`¿Desactivar ${product.name}? No se borrará el historial de ventas.`)) return;
    const { error } = await supabase.from('products').update({ active: false, status: 'Inactivo' }).eq('id', product.id);
    if (error) alert(error.message); else reload();
  }

  return (
    <div className="page products-page">
      <div className="hero compact-hero"><h1>📦 Productos con imágenes</h1><p>Crea artículos visuales con marca, talla, color, código de barras y foto.</p></div>
      <div className="mobile-page-actions">
        <button type="button" className="primary-btn" onClick={() => setFormOpen(true)}>+ Nuevo producto</button>
        <span>{filteredProducts.length} producto(s)</span>
      </div>
      <div className="two-col product-admin-layout">
        {formOpen && <div className="product-modal-backdrop" role="dialog" aria-modal="true" aria-label="Nuevo producto" onMouseDown={(event) => { if (event.target === event.currentTarget) setFormOpen(false); }}>
        <form className="card form-grid product-form-sheet form-open" onSubmit={save} onMouseDown={(event) => event.stopPropagation()}>
          <header className="product-modal-head">
            <div><span className="eyebrow">Registro de catálogo e inventario</span><h3>Nuevo producto</h3><p>Complete lo esencial primero. Las opciones avanzadas quedan al final.</p></div>
            <button className="product-modal-close" type="button" aria-label="Cerrar ventana de nuevo producto" title="Cerrar" onClick={() => setFormOpen(false)}>×</button>
          </header>
          <div className="image-uploader">
            <div className="image-preview-box">
              <img src={previewUrl || APP_ICON} alt="Vista previa" />
              <small>{previewUrl ? 'Vista previa de imagen' : 'Sin imagen cargada'}</small>
            </div>
            <label>Subir imagen del producto<input type="file" accept="image/png,image/jpeg,image/webp" onChange={e=>setImageFile(e.target.files?.[0] || null)} /></label>
            <label>O pegar URL pública de imagen<input value={form.image_url} onChange={e=>{ setImageFile(null); setForm({...form,image_url:e.target.value}); }} placeholder="https://..." /></label>
          </div>
          <div className="form-split">
            <label>Código interno<input value={form.code} onChange={e=>setForm({...form,code:e.target.value})} placeholder="0004" /></label>
            <label>Código de barras<input value={form.barcode} onChange={e=>setForm({...form,barcode:e.target.value})} placeholder="Escanea o escribe el código" /></label>
          </div>
          <label>Nombre del producto<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Nombre comercial del producto" /></label>
          <div className="form-split">
            <label>Categoría
              <select value={form.category_id} onChange={e=>setCategoryById(e.target.value)}>
                <option value="">Selecciona categoría</option>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
            <label>Subcategoría
              <select value={form.subcategory_id} onChange={e=>setSubcategoryById(e.target.value)} disabled={!form.category_id}>
                <option value="">Selecciona subcategoría</option>
                {selectedSubcategories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
          </div>
          <label>Marca<input value={form.brand} onChange={e=>setForm({...form,brand:e.target.value})} placeholder="Marca" /></label>
          <div className="form-split">
            <label>Talla<input value={form.size} onChange={e=>setForm({...form,size:e.target.value})} placeholder="S, M, L, 38, 40" /></label>
            <label>Color<input value={form.color} onChange={e=>setForm({...form,color:e.target.value})} placeholder="Negro, rosa, azul" /></label>
          </div>
          <div className="form-split product-core-price-row">
            <label>Precio de venta<input value={form.price} inputMode="decimal" onChange={e=>setForm({...form,price:e.target.value})} placeholder="0.00" /></label>
            <label>Costo de compra<input value={form.cost} inputMode="decimal" onChange={e=>setForm({...form,cost:e.target.value})} placeholder="0.00" /></label>
          </div>
          <div className="form-split product-core-stock-row">
            <label>Stock inicial<input value={form.stock} inputMode="decimal" onChange={e=>setForm({...form,stock:e.target.value})} /></label>
            <label>Stock mínimo<input value={form.stock_min} inputMode="decimal" onChange={e=>setForm({...form,stock_min:e.target.value})} /></label>
          </div>
          <details className="product-advanced-section">
            <summary>Opciones avanzadas: catálogo, margen y notas</summary>
          <label>Descripción interna<input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="Material, modelo y detalles internos" /></label>
          <section className="catalog-inline-config">
            <strong>Catálogo público</strong>
            <label className="check-row"><input type="checkbox" checked={Boolean(form.public_visible)} onChange={e=>setForm({...form,public_visible:e.target.checked})} /> Mostrar este producto en el catálogo</label>
            <div className="form-split">
              <label>Estado web<select value={form.catalog_status} onChange={e=>setForm({...form,catalog_status:e.target.value})}><option>Borrador</option><option>Publicado</option><option>Oculto</option></select></label>
              <label>Orden de aparición<input value={form.catalog_position} inputMode="numeric" onChange={e=>setForm({...form,catalog_position:e.target.value})} placeholder="999" /></label>
            </div>
            <label className="check-row"><input type="checkbox" checked={Boolean(form.catalog_featured)} onChange={e=>setForm({...form,catalog_featured:e.target.checked})} /> Destacar en portada</label>
            <label>Descripción para clientes<input value={form.catalog_description} onChange={e=>setForm({...form,catalog_description:e.target.value})} placeholder="Texto corto que verá el cliente" /></label>
          </section>
          <div className="form-split">
            <label>Estado del precio<select value={form.price_status} onChange={e=>setForm({...form,price_status:e.target.value})}>{PRICE_STATUS_OPTIONS.map(o=><option key={o}>{o}</option>)}</select></label>
            <label>Margen objetivo sobre costo %<input value={form.margin_target} inputMode="decimal" onChange={e=>setForm({...form,margin_target:e.target.value})} placeholder="50" /></label>
          </div>
          <div className="form-split">
            <label>Precio mínimo permitido<input value={form.min_price} inputMode="decimal" onChange={e=>setForm({...form,min_price:e.target.value})} placeholder="0.00" /></label>
            <label>Precio sugerido<input value={money(suggestedPrice(form.cost, form.margin_target))} readOnly /></label>
          </div>
          <label>Nota de precio<input value={form.price_notes} onChange={e=>setForm({...form,price_notes:e.target.value})} placeholder="Ej.: falta confirmar precio real con proveedor" /></label>
          </details>
          <footer className="product-modal-actions"><button type="button" className="secondary-btn" onClick={() => setFormOpen(false)}>Cancelar</button><button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar producto'}</button></footer>
        </form>
        </div>}
        <section className="card compact-card product-list-panel">
          <h3>Lista de productos</h3>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar por nombre, marca, código o barcode..." /></div>
          <div className="product-card-list">
            {visibleProducts.map(p=>(
              <div className="product-card-item" key={p.id}>
                <img src={productImageSrc(p)} alt={p.name} />
                <div>
                  <strong>{p.name}</strong>
                  <small>{p.code || 'Sin código'} · {p.barcode || 'Sin barcode'} · {p.category || 'General'}{p.subcategory ? ` / ${p.subcategory}` : ''}</small>
                  <small>{p.brand || 'Sin marca'} · {p.size || 'Sin talla'} · {p.color || 'Sin color'}</small>
                  <small><span className={priceBadgeClass(productPriceStatus(p))}>{productPriceStatus(p)}</span> · Ganancia {money(productProfit(p))} · Margen precio {productMarginPercent(p).toFixed(1)}%</small>{p.description && <small>{p.description}</small>}
                </div>
                <div className="product-card-actions">
                  <b>{money(p.price)}</b>
                  <small>Stock {asNum(p.stock)}</small>
                  <button type="button" className="secondary-btn" onClick={()=>deactivateProduct(p)}>Desactivar</button>
                </div>
              </div>
            ))}
            {!filteredProducts.length && <p className="muted">No hay productos activos.</p>}
          </div>
          {filteredProducts.length > PRODUCT_PAGE_SIZE && <div className="compact-pagination"><span>Mostrando {(safeProductPage - 1) * PRODUCT_PAGE_SIZE + 1}–{Math.min(safeProductPage * PRODUCT_PAGE_SIZE, filteredProducts.length)} de {filteredProducts.length}</span><div><button type="button" className="secondary-btn" disabled={safeProductPage === 1} onClick={() => setProductPage(p => Math.max(1, p - 1))}>← Anterior</button><span className="pagination-page">Página {safeProductPage} de {productPageCount}</span><button type="button" className="secondary-btn" disabled={safeProductPage === productPageCount} onClick={() => setProductPage(p => Math.min(productPageCount, p + 1))}>Siguiente →</button></div></div>}
        </section>
      </div>
    </div>
  );
}

function Customers({ customers, reload, profile }) {
  const [form, setForm] = useState({ name:'', phone:'', document:'', address:'', credit_limit:'0' });
  const [query, setQuery] = useState('');
  const filtered = customers.filter(c => `${c.name} ${c.phone} ${c.document}`.toLowerCase().includes(query.toLowerCase()));
  async function save(e) {
    e.preventDefault();
    if (!form.name.trim()) return alert('Coloca el nombre del cliente.');
    const payload = { ...form, credit_limit: asNum(form.credit_limit), status: 'Activo', store_id: profile?.store_id || DEFAULT_STORE_ID, created_by: profile?.id || null };
    const { error } = await supabase.from('customers').insert(payload);
    if (error) alert(error.message); else { setForm({ name:'', phone:'', document:'', address:'', credit_limit:'0' }); reload(); }
  }
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>👥 Clientes</h1><p>Registra contactos para ventas, créditos y seguimiento.</p></div>
      <div className="two-col">
        <form className="card form-grid" onSubmit={save}>
          <label>Nombre<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Nombre del cliente" /></label>
          <label>Teléfono<input value={form.phone} onChange={e=>setForm({...form,phone:e.target.value})} placeholder="Celular o WhatsApp" /></label>
          <label>Documento<input value={form.document} onChange={e=>setForm({...form,document:e.target.value})} placeholder="DNI / RUC / documento" /></label>
          <label>Dirección<input value={form.address} onChange={e=>setForm({...form,address:e.target.value})} placeholder="Dirección o referencia" /></label>
          <label>Límite de crédito<input value={form.credit_limit} onChange={e=>setForm({...form,credit_limit:e.target.value})} inputMode="decimal" placeholder="0.00" /></label>
          <button className="primary-btn">Guardar cliente</button>
        </form>
        <section className="card compact-card">
          <h3>Lista de clientes</h3>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar cliente..." /></div>
          {filtered.map(c => <div className="list-row" key={c.id}><span>{c.name}<small>{c.phone || 'Sin teléfono'} · {c.document || 'Sin documento'}</small></span><strong>{money(c.credit_limit)}</strong></div>)}
          {!filtered.length && <p className="muted">No hay clientes registrados.</p>}
        </section>
      </div>
    </div>
  );
}

function Inventory({ products }) {
  const [query, setQuery] = useState('');
  const [stockFilter, setStockFilter] = useState(() => { try { return sessionStorage.getItem('clomar_inventory_filter_v328') || 'Todos'; } catch (_) { return 'Todos'; } });
  const [activeCategory, setActiveCategory] = useState('');
  useEffect(() => { try { sessionStorage.removeItem('clomar_inventory_filter_v328'); } catch (_) {} }, []);
  const active = products.filter(p => p.active !== false);
  const lowStock = active.filter(p => asNum(p.stock) <= asNum(p.stock_min));
  const noStock = active.filter(p => asNum(p.stock) <= 0);
  const searched = active.filter(p => `${p.code || ''} ${p.barcode || ''} ${p.name || ''} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''}`.toLowerCase().includes(query.toLowerCase()));
  const filtered = searched.filter(p => stockFilter === 'Todos' || (stockFilter === 'Bajo stock' && asNum(p.stock) <= asNum(p.stock_min) && asNum(p.stock) > 0) || (stockFilter === 'Sin stock' && asNum(p.stock) <= 0));
  const byCat = filtered.reduce((acc, p) => { (acc[p.category || 'General'] ||= []).push(p); return acc; }, {});
  const inventoryCategories = Object.keys(byCat).sort((a, b) => a.localeCompare(b, 'es'));
  const visibleCategory = inventoryCategories.includes(activeCategory) ? activeCategory : (inventoryCategories[0] || '');
  const visibleCategoryItems = byCat[visibleCategory] || [];
  const categoryRisk = (cat) => (byCat[cat] || []).filter(p => asNum(p.stock) <= asNum(p.stock_min)).length;
  return (
    <div className="page inventory-page inventory-v327-page">
      <section className="inventory-v327-commandbar">
        <div className="inventory-v327-copy"><span className="eyebrow">Operación de almacén</span><h1>Inventario</h1><p><strong>{active.length}</strong> productos · <strong className="risk">{lowStock.length}</strong> por reponer · <strong>{noStock.length}</strong> agotados · <strong>{inventoryCategories.length}</strong> categorías</p></div>
        <div className="inventory-v327-tools"><div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar producto, código o categoría..." /></div><div className="inventory-stock-filter" aria-label="Filtro de inventario">{['Todos','Bajo stock','Sin stock'].map(option => <button key={option} type="button" className={stockFilter===option ? 'active' : ''} onClick={()=>setStockFilter(option)}>{option}</button>)}</div></div>
      </section>
      {inventoryCategories.length > 0 && <section className="inventory-workspace-v326 inventory-v327-workspace">
        <aside className="card compact-card inventory-side-categories inventory-v327-side">
          <div className="inventory-side-head"><div><span className="eyebrow">Navegación</span><h3>Categorías</h3></div><span>{inventoryCategories.length}</span></div>
          <div className="inventory-category-menu" role="navigation" aria-label="Categorías de inventario">{inventoryCategories.map(cat => {
            const risk = categoryRisk(cat);
            return <button type="button" key={cat} className={visibleCategory === cat ? 'active' : ''} onClick={() => setActiveCategory(cat)}><span><strong>{cat}</strong><small>{byCat[cat].length} producto(s){risk ? ` · ${risk} por reponer` : ''}</small></span><b>{byCat[cat].length}</b></button>;
          })}</div>
        </aside>
        <section className="card compact-card inventory-products-pane inventory-v327-pane">
          <div className="inventory-category-head inventory-v327-head"><div><span className="eyebrow">Categoría activa</span><h3>{visibleCategory}</h3><p>{visibleCategoryItems.length} producto(s) · revise primero los destacados en ámbar</p></div><span className="inventory-guide-badge">Foto · detalle · stock</span></div>
          <div className="inventory-card-grid inventory-v327-grid">{visibleCategoryItems.map(p => {
            const stock = asNum(p.stock); const min = asNum(p.stock_min); const out = stock <= 0; const low = stock <= min && !out; const state = out ? 'out' : low ? 'low' : 'ready'; const stateLabel = out ? 'Agotado' : low ? 'Reponer' : 'Disponible';
            return <article className={`inventory-card-pro inventory-v327-card inventory-${state}`} key={p.id}><img className="product-thumb small" src={productImageSrc(p)} alt={p.name}/><div className="inventory-card-info"><strong>{p.name}</strong><small>{p.code || 'Sin código'} · {p.brand || 'Sin marca'}{p.color ? ` · ${p.color}` : ''}{p.size ? ` · ${p.size}` : ''}</small></div><div className={`stock-pill ${out || low ? 'stock-low' : ''}`}><span>{stateLabel}</span><b>{stock}</b><small>mín. {min}</small></div></article>;
          })}</div>
        </section>
      </section>}
      {!filtered.length && <section className="card compact-card"><p className="muted">No se encontraron productos con ese criterio.</p></section>}
    </div>
  );
}


function StockEntry({ products, reloadProducts, profile, cashSession }) {
  const { movements, reload: reloadMovements } = useStockMovements(profile);
  const [form, setForm] = useState({ product_id:'', provider:'', qty:'1', cost:'0', method:'Efectivo', paid:'0', invoice:'', note:'' });
  const [query, setQuery] = useState('');
  const [historyOpen, setHistoryOpen] = useState(false);
  const selected = products.find(p => p.id === form.product_id) || products[0];
  const total = asNum(form.qty) * asNum(form.cost);
  const entries = movements.filter(m => m.type === 'Entrada');
  const filteredEntries = entries.filter(m => `${m.products?.name || ''} ${m.products?.code || ''} ${m.note || ''}`.toLowerCase().includes(query.toLowerCase())).slice(0, 20);
  const today = todayISO();
  const entriesToday = entries.filter(m => String(m.created_at || '').slice(0,10) === today);
  const unitsToday = entriesToday.reduce((s,m)=>s+asNum(m.qty),0);
  const lastProviders = Array.from(new Set(entries.map(m => String(m.note || '').match(/Proveedor:\s*([^·]+)/)?.[1]?.trim()).filter(Boolean))).slice(0, 8);
  useEffect(() => { if (!form.product_id && products[0]) setForm(f => ({ ...f, product_id: products[0].id, cost: String(products[0].cost || 0) })); }, [products]);
  useEffect(() => { if (form.method !== 'Crédito') setForm(f => ({ ...f, paid: String(total || 0) })); }, [form.qty, form.cost, form.method]);
  async function save(e) {
    e.preventDefault();
    if (!selected) return alert('Selecciona un producto.');
    if (!cashSession?.session?.id) return alert('Primero abre una caja activa. La compra no se registró.');
    if (asNum(form.qty) <= 0) return alert('La cantidad debe ser mayor a cero.');
    if (asNum(form.cost) < 0) return alert('El costo no puede ser negativo.');
    const { data, error } = await supabase.rpc('clomar_register_purchase_r3', {
      p_store_id: profile?.store_id || DEFAULT_STORE_ID,
      p_product_id: selected.id,
      p_provider: form.provider || 'Sin proveedor',
      p_qty: asNum(form.qty),
      p_unit_cost: asNum(form.cost),
      p_payment_method: form.method,
      p_amount_paid: form.method === 'Crédito' ? 0 : asNum(form.paid || total),
      p_document_number: form.invoice || null,
      p_note: form.note || null,
    });
    if (error) return alert(`No se pudo registrar la compra: ${error.message}. No se modificó el inventario ni la caja.`);
    const stockAfter = asNum(data?.new_stock);
    alert(`Compra registrada. Stock actualizado: ${stockAfter}.`);
    setForm({ product_id: selected.id, provider:'', qty:'1', cost:String(form.cost || selected.cost || 0), method:'Efectivo', paid:'0', invoice:'', note:'' });
    await Promise.all([reloadProducts(), reloadMovements(), cashSession.reload?.()]);
  }

  return (
    <div className="page purchases-page">
      <div className="hero compact-hero purchase-command-hero"><div><span className="eyebrow">Abastecimiento</span><h1>Compras e ingreso de mercadería</h1><p>Registre la compra y actualice stock, costo y caja en una sola operación.</p></div><button type="button" className="secondary-btn" onClick={()=>setHistoryOpen(true)}>Ver historial</button></div>
      <section className="purchase-summary-strip"><div><span>Ingresos hoy</span><strong>{entriesToday.length}</strong></div><div><span>Unidades hoy</span><strong>{unitsToday}</strong></div><div><span>Proveedores</span><strong>{lastProviders.length}</strong></div><div><span>Stock actual</span><strong>{selected ? asNum(selected.stock) : 0}</strong></div></section>
      <div className="two-col purchase-workspace">
        <form className="card form-grid purchase-form purchase-form-compact" onSubmit={save}>
          <h3>Nueva compra / ingreso</h3>
          <label>Producto<select value={form.product_id} onChange={e=>{const p=products.find(x=>x.id===e.target.value); setForm({...form,product_id:e.target.value,cost:String(p?.cost || 0)})}}>{products.map(p=><option key={p.id} value={p.id}>{p.code} · {p.name} · Stock {asNum(p.stock)}</option>)}</select></label>
          <div className="form-split"><label>Proveedor<input value={form.provider} onChange={e=>setForm({...form,provider:e.target.value})} placeholder="Nombre del proveedor" /></label><label>Documento<input value={form.invoice} onChange={e=>setForm({...form,invoice:e.target.value})} placeholder="Factura, boleta, guía" /></label></div>
          <div className="form-split"><label>Cantidad ingresada<input value={form.qty} onChange={e=>setForm({...form,qty:e.target.value})} inputMode="decimal" /></label><label>Costo unitario<input value={form.cost} onChange={e=>setForm({...form,cost:e.target.value})} inputMode="decimal" /></label></div>
          <div className="form-split"><label>Método de pago<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option><option>Crédito</option></select></label><label>Monto pagado<input value={form.paid} onChange={e=>setForm({...form,paid:e.target.value})} inputMode="decimal" /></label></div>
          <label>Observación<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Detalle de compra, cambio de costo, referencia..." /></label>
          <div className="total-box"><span>Total compra</span><strong>{money(total)}</strong><small>Stock después: {selected ? asNum(selected.stock) + asNum(form.qty) : 0}</small></div>
          <div className="purchase-submit-row"><div><span>Total de compra</span><strong>{money(total)}</strong></div><button className="primary-btn">Registrar ingreso</button></div>
        </form>
        <section className="card compact-card purchase-preview-card"><h3>Resumen del ingreso</h3>{selected ? <><div className="list-row inventory-product-row"><img className="product-thumb small" src={productImageSrc(selected)} alt={selected.name}/><span>{selected.name}<small>{selected.code} · {selected.category}{selected.subcategory ? ` / ${selected.subcategory}` : ''} · {selected.brand || 'Sin marca'}</small></span><strong>{money(selected.price)}</strong></div><div className="list-row"><span>Stock actual</span><strong>{asNum(selected.stock)}</strong></div><div className="list-row"><span>Stock después</span><strong>{asNum(selected.stock) + asNum(form.qty)}</strong></div><div className="list-row"><span>Costo anterior</span><strong>{money(selected.cost)}</strong></div><div className="list-row"><span>Nuevo costo</span><strong>{money(form.cost)}</strong></div></> : <p className="muted">No hay productos.</p>}<h4>Proveedores recientes</h4>{lastProviders.map(p => <button key={p} type="button" className="provider-chip" onClick={()=>setForm({...form, provider:p})}>{p}</button>)}</section>
      </div>
      {historyOpen && <div className="purchase-history-backdrop" role="dialog" aria-modal="true" onMouseDown={(event)=>{ if (event.target === event.currentTarget) setHistoryOpen(false); }}><section className="card purchase-history-modal" onMouseDown={(event)=>event.stopPropagation()}><header><div><span className="eyebrow">Movimientos recientes</span><h3>Historial de compras</h3><p>Busque por producto, código, proveedor o documento.</p></div><button type="button" className="product-modal-close" onClick={()=>setHistoryOpen(false)}>×</button></header><div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar compra..." /></div><div className="purchase-history-list">{filteredEntries.map(m=><div className="list-row" key={m.id}><span>{m.products?.name || 'Producto'}<small>{fmtDate(m.created_at)} · {m.note || 'Sin nota'}</small></span><strong>+{asNum(m.qty)}</strong></div>)}{!filteredEntries.length && <p className="muted">No hay compras registradas.</p>}</div></section></div>}
    </div>
  );
}

function CashPage({ profile, cashSession }) {
  const { movements, reload: reloadMovements, loading: movementsLoading, loadError: movementsError } = useCashMovements(profile);
  const [opening, setOpening] = useState({ amount: '0', note: '' });
  const [manual, setManual] = useState({ type: 'Ingreso', method: 'Efectivo', amount: '', note: '' });
  const [closing, setClosing] = useState({ counted: '', note: '' });
  const [working, setWorking] = useState('');
  const session = cashSession?.session || null;
  const storeId = profile?.store_id || DEFAULT_STORE_ID;
  const sessionMovements = useMemo(() => {
    if (!session?.id) return [];
    return (movements || []).filter(m => m.cash_session_id === session.id);
  }, [movements, session?.id]);
  const expectedCash = useMemo(() => sessionMovements
    .filter(m => (m.payment_method || 'Efectivo') === 'Efectivo')
    .reduce((sum, m) => {
      if (['Apertura', 'Ingreso'].includes(m.type)) return sum + asNum(m.amount);
      if (['Egreso', 'Retiro', 'Compra'].includes(m.type)) return sum - asNum(m.amount);
      return sum;
    }, 0), [sessionMovements]);
  const byMethod = useMemo(() => sessionMovements.reduce((acc, m) => {
    const key = m.payment_method || 'Sin método';
    const signed = ['Egreso', 'Retiro', 'Compra'].includes(m.type) ? -asNum(m.amount) : asNum(m.amount);
    acc[key] = (acc[key] || 0) + signed;
    return acc;
  }, {}), [sessionMovements]);
  const totalIncome = sessionMovements.filter(m => ['Apertura', 'Ingreso'].includes(m.type)).reduce((s, m) => s + asNum(m.amount), 0);
  const totalOut = sessionMovements.filter(m => ['Egreso', 'Retiro', 'Compra'].includes(m.type)).reduce((s, m) => s + asNum(m.amount), 0);
  const canClose = session && (session.opened_by === profile?.id || ['dueno', 'admin'].includes(String(profile?.role || '').toLowerCase()));

  async function openCash(e) {
    e.preventDefault();
    if (asNum(opening.amount) < 0) return alert('El monto inicial no puede ser negativo.');
    setWorking('open');
    const { error } = await supabase.rpc('clomar_open_cash_session_r3', {
      p_store_id: storeId,
      p_opening_amount: asNum(opening.amount),
      p_note: opening.note || null,
    });
    setWorking('');
    if (error) return alert(`No se pudo abrir caja: ${error.message}`);
    setOpening({ amount: '0', note: '' });
    await Promise.all([cashSession.reload?.(), reloadMovements(300)]);
  }

  async function registerManual(e) {
    e.preventDefault();
    if (!session?.id) return alert('No hay caja abierta.');
    if (asNum(manual.amount) <= 0) return alert('El monto debe ser mayor a cero.');
    if (!manual.note.trim()) return alert('El motivo del movimiento es obligatorio.');
    setWorking('manual');
    const { error } = await supabase.rpc('clomar_register_cash_movement_r3', {
      p_store_id: storeId,
      p_type: manual.type,
      p_payment_method: manual.method,
      p_amount: asNum(manual.amount),
      p_note: manual.note,
    });
    setWorking('');
    if (error) return alert(`No se pudo registrar el movimiento: ${error.message}`);
    setManual({ type: 'Ingreso', method: 'Efectivo', amount: '', note: '' });
    await Promise.all([cashSession.reload?.(), reloadMovements(300)]);
  }

  async function closeCash(e) {
    e.preventDefault();
    if (!session?.id) return;
    if (!canClose) return alert('Solo quien abrió la caja o un administrador puede cerrarla.');
    if (asNum(closing.counted) < 0) return alert('El efectivo contado no puede ser negativo.');
    const difference = asNum(closing.counted) - expectedCash;
    if (Math.abs(difference) >= 0.01 && !closing.note.trim()) return alert('Explica el motivo de la diferencia de caja antes de cerrar.');
    if (!confirm(`Cerrar caja con efectivo esperado ${money(expectedCash)} y contado ${money(closing.counted)}?`)) return;
    setWorking('close');
    const { error } = await supabase.rpc('clomar_close_cash_session_r3', {
      p_cash_session_id: session.id,
      p_counted_amount: asNum(closing.counted),
      p_note: closing.note || null,
    });
    setWorking('');
    if (error) return alert(`No se pudo cerrar caja: ${error.message}`);
    setClosing({ counted: '', note: '' });
    await Promise.all([cashSession.reload?.(), reloadMovements(300)]);
    alert('Caja cerrada correctamente. Las ventas siguientes requerirán abrir un nuevo turno.');
  }

  if (cashSession?.loadError) {
    return <div className="page cash-pro-page"><div className="hero compact-hero"><h1>💰 Caja por turno</h1><p>Control automático de apertura, ventas, movimientos y cierre.</p></div><div className="data-error"><strong>No se pudo cargar la caja R3:</strong> {cashSession.loadError}. Ejecuta primero el SQL de V03.0-R3 y actualiza la página.</div></div>;
  }

  return (
    <div className="page cash-pro-page cash-r3-page">
      <div className="hero compact-hero cash-premium-hero">
        <div><span className="eyebrow">Operación guiada</span><h1>Caja por turno</h1><p>Las ventas nuevas alimentan caja, inventario y reportes automáticamente.</p></div>
        <span className={session ? 'cash-session-badge open' : 'cash-session-badge closed'}>{session ? 'Caja abierta' : 'Caja cerrada'}</span>
      </div>
      <section className="cash-flow-stepper" aria-label="Flujo de caja"><div className={!session ? 'active' : 'done'}><b>1</b><span>Abrir caja</span></div><i>→</i><div className={session ? 'active' : ''}><b>2</b><span>Vender</span></div><i>→</i><div className={session ? '' : ''}><b>3</b><span>Registrar movimiento</span></div><i>→</i><div><b>4</b><span>Arqueo y cierre</span></div></section>
      {movementsError && <div className="data-error"><strong>No se pudieron leer algunos movimientos:</strong> {movementsError}</div>}
      {!session ? (
        <section className="card compact-card cash-open-card">
          <div className="cash-state-copy"><span className="eyebrow">Inicio de turno</span><h3>Inicie su turno de venta</h3><p className="muted">Registre el fondo inicial. Desde este momento, cada venta quedará asociada automáticamente a este turno.</p></div>
          <form className="form-grid cash-open-form" onSubmit={openCash}>
            <label>Fondo inicial en efectivo<input value={opening.amount} onChange={e => setOpening({ ...opening, amount: e.target.value })} inputMode="decimal" placeholder="S/ 0.00" /></label>
            <label>Nota de apertura<input value={opening.note} onChange={e => setOpening({ ...opening, note: e.target.value })} placeholder="Ej. Fondo de cambio" /></label>
            <button className="primary-btn" disabled={working === 'open'}>{working === 'open' ? 'Abriendo...' : 'Abrir caja'}</button>
          </form>
        </section>
      ) : (
        <>
          <section className="card compact-card cash-session-card">
            <div className="cash-session-head"><div><span className="eyebrow">Turno activo</span><h3>Abierta {fmtDate(session.opened_at)}</h3><p className="muted">Fondo inicial: {money(session.opening_amount)}{session.opening_note ? ` · ${session.opening_note}` : ''}</p></div><button type="button" className="secondary-btn" onClick={() => Promise.all([cashSession.reload?.(), reloadMovements(300)])}>{movementsLoading ? 'Actualizando...' : 'Actualizar'}</button></div>
            <div className="cash-kpi-grid"><Kpi label="Efectivo esperado" value={money(expectedCash)} helper="para arqueo" /><Kpi label="Ingresos" value={money(totalIncome)} helper="incluye ventas" /><Kpi label="Egresos" value={money(totalOut)} helper="compras y retiros" /><Kpi label="Movimientos" value={sessionMovements.length} helper="en este turno" /></div>
          </section>
          <div className="two-col cash-r3-columns">
            <form className="card form-grid" onSubmit={registerManual}>
              <h3>Movimiento manual</h3>
              <p className="muted">Use solo para ingresos, egresos o retiros que no provengan de una venta.</p>
              <label>Tipo<select value={manual.type} onChange={e => setManual({ ...manual, type: e.target.value })}><option>Ingreso</option><option>Egreso</option><option>Retiro</option></select></label>
              <label>Método<select value={manual.method} onChange={e => setManual({ ...manual, method: e.target.value })}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option></select></label>
              <label>Monto<input value={manual.amount} onChange={e => setManual({ ...manual, amount: e.target.value })} inputMode="decimal" placeholder="0.00" /></label>
              <label>Motivo obligatorio<input value={manual.note} onChange={e => setManual({ ...manual, note: e.target.value })} placeholder="Ej. Retiro para pago de proveedor" /></label>
              <button className="secondary-btn" disabled={working === 'manual'}>{working === 'manual' ? 'Registrando...' : 'Registrar movimiento'}</button>
            </form>
            <section className="card compact-card"><h3>Resumen por método</h3>{Object.entries(byMethod).map(([method, amount]) => <div className="list-row" key={method}><span>{method}</span><strong>{money(amount)}</strong></div>)}{!sessionMovements.length && <p className="muted">Aún no hay movimientos en el turno.</p>}</section>
          </div>
          <section className="card compact-card extra-row"><div className="report-filter-head"><div><h3>Movimientos del turno</h3><p className="muted">Ventas y compras se incorporan aquí sin sincronización manual.</p></div></div>{sessionMovements.slice(0, 60).map(m => <div className="list-row" key={m.id}><span><strong>{m.type} · {m.payment_method || 'Sin método'}</strong><small>{fmtDate(m.created_at)} · {m.note || 'Sin nota'}</small></span><strong className={['Egreso','Retiro','Compra'].includes(m.type) ? 'danger-text' : ''}>{['Egreso','Retiro','Compra'].includes(m.type) ? '−' : '+'}{money(m.amount)}</strong></div>)}{!sessionMovements.length && <p className="muted">Aún no hay movimientos registrados.</p>}</section>
          <form className="card compact-card extra-row close-cash-r3" onSubmit={closeCash}>
            <div className="report-filter-head"><div><span className="eyebrow">Arqueo y cierre</span><h3>Cerrar turno</h3><p className="muted">Efectivo esperado: {money(expectedCash)}. Si existe diferencia, explique el motivo.</p></div></div>
            <div className="form-split"><label>Efectivo contado<input value={closing.counted} onChange={e => setClosing({ ...closing, counted: e.target.value })} inputMode="decimal" placeholder="0.00" disabled={!canClose} /></label><label>Motivo / observación<input value={closing.note} onChange={e => setClosing({ ...closing, note: e.target.value })} placeholder="Obligatorio si hay diferencia" disabled={!canClose} /></label></div>
            {!canClose && <p className="muted">Solo quien abrió este turno o un administrador puede cerrarlo.</p>}
            <button className="primary-btn" disabled={!canClose || working === 'close'}>{working === 'close' ? 'Cerrando...' : 'Cerrar caja y guardar arqueo'}</button>
          </form>
        </>
      )}
    </div>
  );
}

function Credits({ profile, cashSession }) {
  const [sales, setSales] = useState([]);
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(false);
  const [form, setForm] = useState({ sale_id:'', amount:'0', method:'Efectivo', note:'' });
  const [paymentOpen, setPaymentOpen] = useState(false);
  async function loadCredits() {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data: creditSales } = await supabase
      .from('sales')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .or('status.eq.Crédito,payment_method.eq.Crédito')
      .order('created_at', { ascending: false });
    setSales(creditSales || []);
    const { data: pays, error: payError } = await supabase
      .from('credit_payments')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false });
    if (!payError) setPayments(pays || []);
    setLoading(false);
  }
  useEffect(() => { loadCredits(); }, []);
  const paidBySale = payments.reduce((acc, p) => { acc[p.sale_id] = (acc[p.sale_id] || 0) + asNum(p.amount); return acc; }, {});
  const rows = sales.map(s => ({ ...s, paid: paidBySale[s.id] || 0, balance: Math.max(0, asNum(s.total) - (paidBySale[s.id] || 0)) })).filter(s => s.balance > 0.009);
  const totalPending = rows.reduce((s,v)=>s+asNum(v.balance),0);
  const selected = rows.find(r => r.id === form.sale_id) || rows[0];
  async function savePayment(e) {
    e.preventDefault();
    if (!selected) return alert('No hay crédito seleccionado.');
    if (!cashSession?.session?.id) return alert('Primero abre una caja activa para registrar un abono.');
    const amount = Math.min(asNum(form.amount), asNum(selected.balance));
    if (amount <= 0) return alert('El abono debe ser mayor a cero.');
    const { error } = await supabase.rpc('clomar_register_credit_payment_r3', {
      p_sale_id: selected.id,
      p_amount: amount,
      p_payment_method: form.method,
      p_note: form.note || null,
    });
    if (error) return alert(`No se pudo registrar el abono: ${error.message}. No se realizó un registro parcial.`);
    setForm({ sale_id:'', amount:'0', method:'Efectivo', note:'' });
    setPaymentOpen(false);
    await Promise.all([loadCredits(), cashSession.reload?.()]);
  }

  const byClient = rows.reduce((acc, s) => { acc[s.customer_name || 'Cliente'] = (acc[s.customer_name || 'Cliente'] || 0) + asNum(s.balance); return acc; }, {});
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>💳 Créditos</h1><p>Control de deuda, abonos y saldo real por cliente.</p></div>
      <div className="kpi-grid credit-compact-kpis"><Kpi label="Saldo pendiente" value={money(totalPending)} helper={`${rows.length} comprobantes`} /><Kpi label="Clientes" value={Object.keys(byClient).length} helper="Con deuda" /><Kpi label="Abonado" value={money(payments.reduce((s,p)=>s+asNum(p.amount),0))} helper="Historial" /><Kpi label="Estado" value={loading ? 'Cargando' : 'Activo'} helper="Supabase" /></div>
      <div className="mobile-page-actions">
        <button type="button" className="primary-btn" onClick={() => setPaymentOpen(true)} disabled={!rows.length}>+ Registrar abono</button>
        <span>{rows.length} crédito(s) pendiente(s)</span>
      </div>
      <div className="two-col credit-layout-pro">
        <section className="card compact-card"><h3>Deuda por cliente</h3>{Object.entries(byClient).map(([client, amount])=><div className="list-row" key={client}><span>{client}</span><strong>{money(amount)}</strong></div>)}{!rows.length && <p className="muted">No hay créditos pendientes.</p>}</section>
        <form className={`card form-grid credit-payment-sheet ${paymentOpen ? 'form-open' : ''}`} onSubmit={savePayment}>
          <button className="sheet-close-btn form-sheet-close" type="button" onClick={() => setPaymentOpen(false)}>Cerrar ×</button>
          <h3>Registrar abono</h3>
          <label>Crédito<select value={form.sale_id || selected?.id || ''} onChange={e=>setForm({...form,sale_id:e.target.value})}>{rows.map(s=><option key={s.id} value={s.id}>B{s.receipt_number} · {s.customer_name} · Saldo {money(s.balance)}</option>)}</select></label>
          <label>Monto del abono<input value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})} inputMode="decimal" placeholder="Monto recibido" /></label>
          <label>Método<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option></select></label>
          <label>Nota<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Observación del abono" /></label>
          <button className="primary-btn" disabled={!rows.length}>Guardar abono</button>
        </form>
      </div>
      <section className="card compact-card extra-row"><h3>Comprobantes pendientes</h3>{rows.map(s=><div className="list-row" key={s.id}><span>B{s.receipt_number} · {s.customer_name || 'Cliente'}<small>Total {money(s.total)} · Abonado {money(s.paid)} · {fmtDate(s.created_at)}</small></span><strong>{money(s.balance)}</strong></div>)}{!rows.length && <p className="muted">Sin comprobantes pendientes.</p>}</section>
    </div>
  );
}


function Reports({ products, profile }) {
  const [sales, setSales] = useState([]);
  const [items, setItems] = useState([]);
  const [movements, setMovements] = useState([]);
  const [payments, setPayments] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadErrors, setLoadErrors] = useState([]);
  const [startDate, setStartDate] = useState(todayISO());
  const [endDate, setEndDate] = useState(todayISO());
  const [methodFilter, setMethodFilter] = useState('todos');
  const [sellerFilter, setSellerFilter] = useState('todos');
  const [query, setQuery] = useState('');

  const storeId = profile?.store_id || DEFAULT_STORE_ID;
  const endDatePlusOne = (value) => {
    const d = new Date(`${value || todayISO()}T00:00:00`);
    d.setDate(d.getDate() + 1);
    return d.toISOString().slice(0, 10);
  };

  async function loadReport() {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const from = startDate || todayISO();
    const to = endDatePlusOne(endDate || from);

    const salesQuery = supabase
      .from('sales')
      .select('*')
      .eq('store_id', storeId)
      .gte('created_at', from)
      .lt('created_at', to)
      .order('created_at', { ascending: false })
      .limit(2000);

    const itemsQuery = supabase
      .from('sale_items')
      .select('*')
      .eq('store_id', storeId)
      .gte('created_at', from)
      .lt('created_at', to)
      .limit(5000);

    const movementsQuery = supabase
      .from('cash_movements')
      .select('*')
      .eq('store_id', storeId)
      .gte('created_at', from)
      .lt('created_at', to)
      .order('created_at', { ascending: false })
      .limit(3000);

    const paymentsQuery = supabase
      .from('credit_payments')
      .select('*')
      .eq('store_id', storeId)
      .gte('created_at', from)
      .lt('created_at', to)
      .order('created_at', { ascending: false })
      .limit(2000);

    const profilesQuery = supabase
      .from('profiles')
      .select('id,email,full_name,role,status,store_id')
      .eq('store_id', storeId);

    const [salesRes, itemsRes, movementsRes, paymentsRes, profilesRes] = await Promise.all([
      salesQuery,
      itemsQuery,
      movementsQuery,
      paymentsQuery,
      profilesQuery,
    ]);

    const errors = [
      ['Ventas', salesRes.error],
      ['Detalle de ventas', itemsRes.error],
      ['Caja', movementsRes.error],
      ['Créditos', paymentsRes.error],
      ['Usuarios', profilesRes.error],
    ].filter(([, err]) => err).map(([label, err]) => `${label}: ${err.message || 'error de consulta'}`);
    setLoadErrors(errors);
    setSales(salesRes.data || []);
    setItems(itemsRes.data || []);
    setMovements(movementsRes.data || []);
    setPayments(paymentsRes.data || []);
    setProfiles(profilesRes.data || []);
    setLoading(false);
  }

  useEffect(() => { loadReport(); }, [profile?.store_id, startDate, endDate]);

  const sellerName = (idOrEmail) => {
    if (!idOrEmail) return 'Sin vendedor';
    const found = profiles.find(p => p.id === idOrEmail || p.email === idOrEmail);
    return found?.full_name || found?.email || String(idOrEmail);
  };

  const saleById = useMemo(() => {
    const map = {};
    sales.forEach(s => { map[s.id] = s; });
    return map;
  }, [sales]);

  const productById = useMemo(() => {
    const map = {};
    products.forEach(p => { map[p.id] = p; });
    return map;
  }, [products]);

  const methods = useMemo(() => ['todos', ...Array.from(new Set(sales.map(s => s.payment_method || 'Sin método')))], [sales]);
  const sellers = useMemo(() => {
    const ids = Array.from(new Set(sales.map(s => s.user_id || s.seller_email || 'sin-vendedor')));
    return ['todos', ...ids];
  }, [sales]);

  const filteredSales = useMemo(() => sales.filter(s => {
    const methodOk = methodFilter === 'todos' || (s.payment_method || 'Sin método') === methodFilter;
    const sellerKey = s.user_id || s.seller_email || 'sin-vendedor';
    const sellerOk = sellerFilter === 'todos' || sellerKey === sellerFilter;
    const q = normalizeText(query);
    const textOk = !q || [
      s.receipt_number,
      s.customer_name,
      s.payment_method,
      s.status,
      sellerName(s.user_id || s.seller_email),
    ].some(v => normalizeText(v).includes(q));
    return methodOk && sellerOk && textOk;
  }), [sales, methodFilter, sellerFilter, query, profiles]);

  const filteredSaleIds = useMemo(() => new Set(filteredSales.map(s => s.id)), [filteredSales]);

  const filteredItems = useMemo(() => items.filter(it => filteredSaleIds.has(it.sale_id)), [items, filteredSaleIds]);

  const creditSales = filteredSales.filter(s => s.status === 'Crédito' || s.payment_method === 'Crédito');
  const totalSales = filteredSales.reduce((sum, s) => sum + asNum(s.total), 0);
  const totalCost = filteredItems.reduce((sum, it) => {
    const cost = asNum(it.unit_cost ?? productById[it.product_id]?.cost);
    return sum + cost * asNum(it.qty);
  }, 0);
  const totalProfit = filteredItems.reduce((sum, it) => {
    if (it.profit !== null && it.profit !== undefined) return sum + asNum(it.profit);
    const cost = asNum(it.unit_cost ?? productById[it.product_id]?.cost);
    return sum + ((asNum(it.price) - cost) * asNum(it.qty));
  }, 0);
  const totalQty = filteredItems.reduce((sum, it) => sum + asNum(it.qty), 0);
  const ticketAvg = filteredSales.length ? totalSales / filteredSales.length : 0;
  const marginPct = totalSales > 0 ? (totalProfit / totalSales) * 100 : 0;
  const totalCredit = creditSales.reduce((sum, s) => sum + asNum(s.total), 0);
  const totalCreditPaid = payments.reduce((sum, p) => sum + asNum(p.amount), 0);
  const pendingCredit = Math.max(0, totalCredit - totalCreditPaid);

  const groupRows = (rows, keyFn, valueFn) => {
    const acc = {};
    rows.forEach(row => {
      const key = keyFn(row) || 'Sin dato';
      acc[key] = (acc[key] || 0) + asNum(valueFn(row));
    });
    return Object.entries(acc).sort((a,b) => b[1] - a[1]);
  };

  const byMethod = groupRows(filteredSales, s => s.payment_method || 'Sin método', s => s.total);
  const bySeller = groupRows(filteredSales, s => sellerName(s.user_id || s.seller_email || 'sin-vendedor'), s => s.total);
  const byCategory = groupRows(filteredItems, it => productById[it.product_id]?.category || it.category || 'Sin categoría', it => it.subtotal || asNum(it.qty) * asNum(it.price));
  const byProduct = Object.values(filteredItems.reduce((acc, it) => {
    const product = productById[it.product_id] || {};
    const name = it.product_name || product.name || 'Producto';
    const key = `${it.product_id || name}`;
    const cost = asNum(it.unit_cost ?? product.cost);
    const subtotal = asNum(it.subtotal || asNum(it.qty) * asNum(it.price));
    const profit = it.profit !== null && it.profit !== undefined ? asNum(it.profit) : (asNum(it.price) - cost) * asNum(it.qty);
    if (!acc[key]) acc[key] = { name, code: product.code || '', qty: 0, total: 0, profit: 0 };
    acc[key].qty += asNum(it.qty);
    acc[key].total += subtotal;
    acc[key].profit += profit;
    return acc;
  }, {})).sort((a,b) => b.total - a.total);

  const cashIn = movements.filter(m => ['Ingreso','Apertura'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  const cashOut = movements.filter(m => ['Egreso','Compra','Retiro','Compra crédito'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  const cashNet = cashIn - cashOut;

  const reportRows = filteredSales.map(s => ({
    Comprobante: receiptNumber(s),
    Fecha: fmtDate(s.created_at),
    Cliente: s.customer_name || 'Cliente',
    Vendedor: sellerName(s.user_id || s.seller_email),
    Metodo: s.payment_method || 'Sin método',
    Estado: s.status || '',
    Total: asNum(s.total),
  }));

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(reportRows), 'Ventas');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byProduct.map(p => ({
      Producto: p.name,
      Codigo: p.code,
      Cantidad: p.qty,
      Vendido: p.total,
      Ganancia: p.profit,
    }))), 'Productos');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bySeller.map(([seller, total]) => ({ Vendedor: seller, Total: total }))), 'Vendedores');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byMethod.map(([method, total]) => ({ Metodo: method, Total: total }))), 'Metodos');
    XLSX.writeFile(wb, `reporte-clomar-${startDate}-a-${endDate}.xlsx`);
  }

  function printReport() {
    const rowsHtml = reportRows.map(r => `<tr><td>${escapeHtml(r.Comprobante)}</td><td>${escapeHtml(r.Fecha)}</td><td>${escapeHtml(r.Cliente)}</td><td>${escapeHtml(r.Vendedor)}</td><td>${escapeHtml(r.Metodo)}</td><td>${money(r.Total)}</td></tr>`).join('');
    const topProductsHtml = byProduct.slice(0, 10).map(p => `<tr><td>${escapeHtml(p.name)}</td><td>${p.qty}</td><td>${money(p.total)}</td><td>${money(p.profit)}</td></tr>`).join('');
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>Reporte Clomar Store</title><style>
      body{font-family:Arial,Helvetica,sans-serif;color:#111827;padding:24px}
      h1{margin:0 0 4px;font-size:24px} .muted{color:#64748b}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0}.kpi{border:1px solid #e5e7eb;border-radius:12px;padding:12px}.kpi span{font-size:10px;text-transform:uppercase;color:#64748b;font-weight:800}.kpi strong{display:block;font-size:20px;margin-top:5px}
      table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border-bottom:1px solid #e5e7eb;padding:7px;text-align:left;font-size:12px}th{text-transform:uppercase;color:#64748b;font-size:10px}
      @media print{@page{size:A4;margin:12mm}}
    
@media print{.price strong{white-space:nowrap!important}.composition-two-column .price strong{white-space:nowrap!important}.composition-two-column.price-max .price strong{font-size:13.6px!important}.label.orientation-vertical.vertical-compact{gap:.38mm!important;align-content:start!important}.label.orientation-vertical.vertical-compact .codes{align-items:center!important}.label.orientation-vertical.price-max .price strong{font-size:26px!important}.density-vertical-medium.orientation-vertical{grid-template-rows:3.95mm 11.9mm 14.3mm 25.4mm!important;gap:.38mm!important;padding:1.95mm 2.05mm 1.6mm!important}.density-vertical-medium.orientation-vertical.price-max .price strong{font-size:25.8px!important}.density-vertical-medium.orientation-vertical .price strong{font-size:22.8px!important}.density-vertical-medium.orientation-vertical.mode-barcode .barcode-svg{height:17.4mm!important}.density-vertical-compact.orientation-vertical{grid-template-rows:3.65mm 10.4mm 13.4mm 24.2mm!important;gap:.34mm!important;padding:1.7mm 1.7mm 1.45mm!important}.density-vertical-compact.orientation-vertical .price strong{font-size:20.2px!important}.density-vertical-compact.orientation-vertical.price-max .price strong{font-size:22.8px!important}.density-vertical-compact.orientation-vertical.mode-barcode .barcode-svg{height:16.8mm!important}.density-vertical-large.orientation-vertical{grid-template-rows:5.1mm 18.6mm 18.7mm 32.4mm!important;gap:.58mm!important;padding:2.75mm 2.95mm 2.35mm!important}.density-vertical-large.orientation-vertical.price-max .price strong{font-size:33px!important}.density-vertical-large.orientation-vertical .price strong{font-size:29.2px!important}.label.orientation-vertical.composition-two-column .price{display:flex!important;flex-direction:row!important;align-items:baseline!important}.label.orientation-vertical.composition-two-column .price strong{white-space:nowrap!important;font-size:22px!important}.qr-large .qr-wrap{grid-template-columns:13.8mm!important;grid-template-rows:13.8mm auto!important}.qr-large .qr{width:13.8mm!important;height:13.8mm!important}.qr-max .qr-wrap{grid-template-columns:17mm!important;grid-template-rows:17mm auto!important}.qr-max .qr{width:17mm!important;height:17mm!important}}
</style></head><body>
      <h1>Clomar Store — Reporte profesional</h1>
      <p class="muted">Periodo: ${escapeHtml(startDate)} a ${escapeHtml(endDate)} · Generado: ${escapeHtml(fmtDate(new Date()))}</p>
      <section class="kpis"><div class="kpi"><span>Ventas</span><strong>${money(totalSales)}</strong></div><div class="kpi"><span>Ganancia</span><strong>${money(totalProfit)}</strong></div><div class="kpi"><span>Tickets</span><strong>${filteredSales.length}</strong></div><div class="kpi"><span>Caja neta</span><strong>${money(cashNet)}</strong></div></section>
      <h2>Productos más vendidos</h2><table><thead><tr><th>Producto</th><th>Cant.</th><th>Vendido</th><th>Ganancia</th></tr></thead><tbody>${topProductsHtml}</tbody></table>
      <h2>Ventas</h2><table><thead><tr><th>N°</th><th>Fecha</th><th>Cliente</th><th>Vendedor</th><th>Método</th><th>Total</th></tr></thead><tbody>${rowsHtml}</tbody></table>
      <script>window.onload=()=>setTimeout(()=>window.print(),250)</script>
    </body></html>`;
    const win = window.open('', '_blank', 'width=1000,height=800');
    if (!win) return alert('Permite ventanas emergentes para imprimir o guardar el reporte.');
    win.document.open();
    win.document.write(html);
    win.document.close();
  }

  const maxMethod = Math.max(1, ...byMethod.map(([,v]) => v));
  const maxSeller = Math.max(1, ...bySeller.map(([,v]) => v));
  const maxCategory = Math.max(1, ...byCategory.map(([,v]) => v));

  return (
    <div className="page reports-pro-page">
      <div className="hero compact-hero">
        <h1>📈 Reportes profesionales</h1>
        <p>Ventas, ganancias, vendedores, productos, categorías, caja y créditos con filtros comerciales.</p>
      </div>
      {loadErrors.length > 0 && <div className="data-error"><strong>Hay datos que no se pudieron leer:</strong> {loadErrors.join(' · ')}. Ejecute el SQL R2 y pulse Actualizar.</div>}

      <section className="card compact-card report-filters">
        <div className="report-filter-head">
          <div>
            <h3>Filtros del reporte</h3>
            <p className="muted">Usa fechas, vendedor, método de pago o búsqueda para analizar resultados reales.</p>
          </div>
          <div className="report-actions">
            <button className="secondary-btn" type="button" onClick={loadReport}>{loading ? 'Cargando...' : 'Actualizar'}</button>
            <button className="secondary-btn" type="button" onClick={exportExcel}>Exportar Excel</button>
            <button className="primary-btn" type="button" onClick={printReport}>PDF / Imprimir</button>
          </div>
        </div>
        <div className="report-filter-grid">
          <label>Desde<input type="date" value={startDate} onChange={e=>setStartDate(e.target.value)} /></label>
          <label>Hasta<input type="date" value={endDate} onChange={e=>setEndDate(e.target.value)} /></label>
          <label>Método<select value={methodFilter} onChange={e=>setMethodFilter(e.target.value)}>{methods.map(m => <option key={m} value={m}>{m === 'todos' ? 'Todos' : m}</option>)}</select></label>
          <label>Vendedor<select value={sellerFilter} onChange={e=>setSellerFilter(e.target.value)}>{sellers.map(s => <option key={s} value={s}>{s === 'todos' ? 'Todos' : sellerName(s)}</option>)}</select></label>
          <label className="report-query">Buscar<input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Cliente, comprobante, vendedor..." /></label>
        </div>
      </section>

      <div className="report-kpi-grid">
        <Kpi label="Total vendido" value={money(totalSales)} helper={`${filteredSales.length} comprobantes`} />
        <Kpi label="Ganancia bruta" value={money(totalProfit)} helper={`Margen ${marginPct.toFixed(1)}%`} />
        <Kpi label="Ticket promedio" value={money(ticketAvg)} helper={`${totalQty} unidades vendidas`} />
        <Kpi label="Caja neta" value={money(cashNet)} helper={`Ingresos ${money(cashIn)} · Egresos ${money(cashOut)}`} />
        <Kpi label="Crédito pendiente" value={money(pendingCredit)} helper={`${creditSales.length} ventas a crédito`} />
        <Kpi label="Costo vendido" value={money(totalCost)} helper="Base para ganancia" />
      </div>

      <div className="report-dashboard-grid">
        <section className="card compact-card">
          <h3>Ventas por método</h3>
          <div className="bar-list">
            {byMethod.map(([method, methodTotal]) => <div className="bar-row" key={method}><div><strong>{method}</strong><small>{((methodTotal / maxMethod) * 100).toFixed(0)}% del mayor</small></div><div className="bar-track"><span style={{ width: `${Math.max(5, (methodTotal / maxMethod) * 100)}%` }} /></div><b>{money(methodTotal)}</b></div>)}
            {!byMethod.length && <p className="muted">Sin ventas en el periodo.</p>}
          </div>
        </section>

        <section className="card compact-card">
          <h3>Ventas por vendedor</h3>
          <div className="bar-list">
            {bySeller.map(([seller, sellerTotal]) => <div className="bar-row" key={seller}><div><strong>{seller}</strong><small>Rendimiento comercial</small></div><div className="bar-track"><span style={{ width: `${Math.max(5, (sellerTotal / maxSeller) * 100)}%` }} /></div><b>{money(sellerTotal)}</b></div>)}
            {!bySeller.length && <p className="muted">Sin vendedores con ventas.</p>}
          </div>
        </section>

        <section className="card compact-card">
          <h3>Ventas por categoría</h3>
          <div className="bar-list">
            {byCategory.map(([cat, catTotal]) => <div className="bar-row" key={cat}><div><strong>{cat}</strong><small>Según productos vendidos</small></div><div className="bar-track"><span style={{ width: `${Math.max(5, (catTotal / maxCategory) * 100)}%` }} /></div><b>{money(catTotal)}</b></div>)}
            {!byCategory.length && <p className="muted">Sin categorías vendidas.</p>}
          </div>
        </section>

        <section className="card compact-card">
          <h3>Productos más vendidos</h3>
          <div className="mini-table">
            <div className="mini-table-head"><span>Producto</span><span>Cant.</span><span>Vendido</span><span>Ganancia</span></div>
            {byProduct.slice(0, 10).map(p => <div className="mini-table-row" key={p.name}><span><strong>{p.name}</strong><small>{p.code || 'Sin código'}</small></span><b>{p.qty}</b><b>{money(p.total)}</b><b className={p.profit < 0 ? 'danger-text' : ''}>{money(p.profit)}</b></div>)}
            {!byProduct.length && <p className="muted">Sin productos vendidos en el periodo.</p>}
          </div>
        </section>
      </div>

      <section className="card compact-card report-sales-table">
        <div className="report-filter-head">
          <div><h3>Detalle de ventas</h3><p className="muted">Historial filtrado por fecha, vendedor y método.</p></div>
          <span className="result-pill">{filteredSales.length} ventas</span>
        </div>
        <div className="mini-table sales-detail-table">
          <div className="mini-table-head"><span>N°</span><span>Fecha</span><span>Cliente</span><span>Vendedor</span><span>Método</span><span>Total</span></div>
          {filteredSales.map(s => <div className="mini-table-row" key={s.id}><span>{receiptNumber(s)}</span><span>{fmtDate(s.created_at)}</span><span>{s.customer_name || 'Cliente'}</span><span>{sellerName(s.user_id || s.seller_email)}</span><span>{s.payment_method || 'Sin método'}</span><b>{money(s.total)}</b></div>)}
          {!filteredSales.length && <p className="muted">No hay ventas con los filtros actuales.</p>}
        </div>
      </section>
    </div>
  );
}

function UsersAdmin({ profile }) {
  const [profiles, setProfiles] = useState([]);
  const [saving, setSaving] = useState('');
  const [drafts, setDrafts] = useState({});
  const roles = ['dueno', 'admin', 'cajero', 'almacen', 'lectura'];

  async function loadProfiles() {
    const { data, error } = await supabase
      .from('profiles')
      .select('id,email,full_name,role,status,created_at,store_id')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: true });
    if (!error) setProfiles(data || []);
  }

  useEffect(() => { loadProfiles(); }, [profile?.store_id]);

  function draft(row) {
    return drafts[row.id] || { full_name: row.full_name || '', role: row.role || 'cajero', status: row.status || 'Activo' };
  }

  function setDraftValue(row, key, value) {
    setDrafts(prev => ({ ...prev, [row.id]: { ...draft(row), [key]: value } }));
  }

  async function saveProfile(row) {
    const next = draft(row);
    setSaving(row.id);
    const { error } = await supabase
      .from('profiles')
      .update({ full_name: next.full_name, role: next.role, status: next.status, updated_at: new Date().toISOString() })
      .eq('id', row.id);
    setSaving('');
    if (error) return alert(error.message || 'No se pudo actualizar el usuario. Revisa políticas SQL V01.6.');
    await loadProfiles();
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🧑‍💼 Usuarios y roles</h1><p>Administra permisos por usuario. Los usuarios nuevos se crean primero en Supabase Auth.</p></div>
      <section className="card compact-card">
        <h3>Cómo crear un usuario nuevo</h3>
        <p className="muted">Ruta: Supabase → Authentication → Users → Add user → Create new user. Activa Auto Confirm User. Luego regresa aquí para asignar rol.</p>
      </section>
      <section className="card compact-card extra-row">
        <h3>Usuarios registrados</h3>
        {profiles.map(row => {
          const d = draft(row);
          return (
            <div className="list-row" key={row.id}>
              <span>
                <strong>{row.email || 'Sin correo'}</strong>
                <small>Creado: {fmtDate(row.created_at)} · Rol actual: {ROLE_LABELS[row.role] || row.role}</small>
              </span>
              <input value={d.full_name} onChange={e=>setDraftValue(row, 'full_name', e.target.value)} placeholder="Nombre" />
              <select value={d.role} onChange={e=>setDraftValue(row, 'role', e.target.value)}>
                {roles.map(r => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
              </select>
              <select value={d.status} onChange={e=>setDraftValue(row, 'status', e.target.value)}>
                <option>Activo</option><option>Inactivo</option>
              </select>
              <button className="secondary-btn" onClick={() => saveProfile(row)} disabled={saving === row.id}>{saving === row.id ? 'Guardando...' : 'Guardar'}</button>
            </div>
          );
        })}
        {!profiles.length && <p className="muted">No hay perfiles registrados todavía.</p>}
      </section>
    </div>
  );

}

function FriendlyNotice({ notice, onClose, primaryText = 'Entendido', secondaryText, onSecondary }) {
  if (!notice) return null;
  return (
    <div className="notice-backdrop" role="dialog" aria-modal="true">
      <div className={`notice-card ${notice.type ? `notice-${notice.type}` : ''}`}>
        <div className="notice-icon">{notice.icon || '⚠️'}</div>
        <div className="notice-content">
          <h3>{notice.title || 'Aviso'}</h3>
          {notice.message && <p>{notice.message}</p>}
          {notice.details && <div className="notice-details">{notice.details}</div>}
          <div className="notice-actions">
            {secondaryText && <button type="button" className="secondary-btn" onClick={onSecondary}>{secondaryText}</button>}
            <button type="button" className="primary-btn" onClick={onClose}>{primaryText}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PricesAdmin({ products = [], reload, profile }) {
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('pendientes');
  const [drafts, setDrafts] = useState({});
  const [savingId, setSavingId] = useState('');
  const [notice, setNotice] = useState(null);
  const [expandedId, setExpandedId] = useState(null);

  const rows = useMemo(() => products.filter(p => {
    const q = normalizeText(query);
    const match = !q || [p.name, p.code, p.barcode, p.brand, p.category, p.subcategory, p.color].some(v => normalizeText(v).includes(q));
    const status = productPriceStatus(p);
    const margin = productMarginPercent(p);
    const byFilter = filter === 'todos'
      || (filter === 'pendientes' && status !== 'Validado')
      || (filter === 'sin_costo' && asNum(p.cost) <= 0)
      || (filter === 'sin_precio' && asNum(p.price) <= 0)
      || (filter === 'margen_bajo' && status === 'Validado' && asNum(p.price) > 0 && margin < 25);
    return match && byFilter;
  }), [products, query, filter]);

  const summary = useMemo(() => {
    const pending = products.filter(p => productPriceStatus(p) !== 'Validado').length;
    const validated = products.filter(p => productPriceStatus(p) === 'Validado').length;
    const review = products.filter(p => productPriceStatus(p) === 'Revisar').length;
    const noCost = products.filter(p => asNum(p.cost) <= 0).length;
    const noPrice = products.filter(p => asNum(p.price) <= 0).length;
    const lowMargin = products.filter(p => productPriceStatus(p) === 'Validado' && asNum(p.price) > 0 && productMarginPercent(p) < 25).length;
    return { pending, validated, review, noCost, noPrice, lowMargin };
  }, [products]);

  function draft(product) {
    return drafts[product.id] || {
      cost: String(product.cost ?? 0),
      price: String(product.price ?? 0),
      margin_target: String(product.margin_target ?? 50),
      min_price: String(product.min_price ?? 0),
      price_status: productPriceStatus(product),
      price_notes: product.price_notes || '',
    };
  }
  function setDraft(product, key, value) {
    setDrafts(prev => ({ ...prev, [product.id]: { ...draft(product), [key]: value } }));
  }
  function useSuggested(product) {
    const d = draft(product);
    setDraft(product, 'price', String(suggestedPrice(d.cost, d.margin_target)));
    setDraft(product, 'price_status', 'Revisar');
  }
  function markValidated(product) {
    const d = draft(product);
    setDrafts(prev => ({ ...prev, [product.id]: { ...d, price_status: 'Validado', price_notes: d.price_notes || 'Precio validado desde panel de control.' } }));
  }
  function resetDraft(product) {
    setDrafts(prev => { const next = { ...prev }; delete next[product.id]; return next; });
  }
  function priceHealth(d) {
    const cost = asNum(d.cost);
    const price = asNum(d.price);
    const min = asNum(d.min_price);
    if (price <= 0) return { kind: 'bad', text: 'Sin precio de venta' };
    if (cost <= 0) return { kind: 'warn', text: 'Costo pendiente' };
    if (min > 0 && price < min) return { kind: 'bad', text: 'Debajo del mínimo' };
    const margin = ((price - cost) / price) * 100;
    if (margin < 0) return { kind: 'bad', text: 'Vende con pérdida' };
    if (margin < 25) return { kind: 'warn', text: 'Margen bajo' };
    return { kind: 'ok', text: 'Precio saludable' };
  }

  async function savePrice(product) {
    const d = draft(product);
    const payload = {
      cost: asNum(d.cost),
      price: asNum(d.price),
      margin_target: asNum(d.margin_target || 50),
      min_price: asNum(d.min_price || 0),
      price_status: d.price_status || 'Pendiente',
      price_notes: d.price_notes || '',
      price_updated_at: new Date().toISOString(),
      price_updated_by: profile?.id || null,
    };
    if (payload.price_status === 'Validado' && payload.price <= 0) return setNotice({ type: 'warning', icon: '💰', title: 'No se puede validar con precio 0', message: 'Coloca un precio de venta mayor a 0 antes de marcar el producto como Validado.' });
    if (payload.price_status === 'Validado' && payload.min_price > 0 && payload.price < payload.min_price) return setNotice({ type: 'warning', icon: '📉', title: 'Precio menor al mínimo permitido', message: 'El precio validado no puede ser menor al precio mínimo definido para este producto.' });
    setSavingId(product.id);
    try {
      const { error } = await supabase.from('products').update(payload).eq('id', product.id);
      if (error) throw error;
      await supabase.from('product_price_history').insert({
        store_id: profile?.store_id || DEFAULT_STORE_ID,
        product_id: product.id,
        user_id: profile?.id || null,
        old_cost: asNum(product.cost),
        new_cost: payload.cost,
        old_price: asNum(product.price),
        new_price: payload.price,
        old_status: productPriceStatus(product),
        new_status: payload.price_status,
        note: payload.price_notes,
      }).then(()=>{});
      setDrafts(prev => { const next = { ...prev }; delete next[product.id]; return next; });
      setNotice({ type: 'success', icon: '✅', title: 'Precio actualizado', message: `${product.name} quedó como ${payload.price_status}.` });
      await reload?.();
    } catch (error) {
      setNotice({ type: 'warning', icon: '⚠️', title: 'No se pudo actualizar el precio', message: error.message || 'Verifica que ejecutaste el SQL V01.12.' });
    } finally {
      setSavingId('');
    }
  }

  return <div className="page prices-page prices-page-pro">
    <FriendlyNotice notice={notice} onClose={()=>setNotice(null)} />
    <div className="hero compact-hero"><h1>💰 Control de precios</h1><p>Valida costos, precios y márgenes antes de vender, imprimir etiquetas o generar comprobantes.</p></div>
    <div className="price-command-center">
      <div className="price-command-copy">
        <span className="eyebrow">No asumir precios</span>
        <h3>Los productos pendientes no se venden hasta validar costo y precio.</h3>
        <p>Usa esta pantalla para confirmar costos reales, definir precio mínimo, revisar margen y dejar trazabilidad de cada cambio.</p>
      </div>
      <div className="price-kpi-row">
        <Kpi label="Pendientes" value={summary.pending} helper="requieren validación" />
        <Kpi label="Validados" value={summary.validated} helper="listos para vender" />
        <Kpi label="Sin costo" value={summary.noCost} helper="afecta ganancia" />
        <Kpi label="Margen bajo" value={summary.lowMargin} helper="menor a 25%" />
      </div>
    </div>

    <section className="card compact-card price-filter-card">
      <div className="price-filter-head">
        <div><h3>Filtros de control</h3><p className="muted">Encuentra productos pendientes, sin costo, sin precio o con margen bajo.</p></div>
        <span className="result-pill">{rows.length} productos</span>
      </div>
      <div className="form-split price-filter-grid">
        <label>Buscar producto<input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Nombre, código, marca, barcode..." /></label>
        <label>Estado<select value={filter} onChange={e=>setFilter(e.target.value)}><option value="pendientes">Pendientes/Revisar</option><option value="sin_costo">Sin costo</option><option value="sin_precio">Sin precio</option><option value="margen_bajo">Margen bajo</option><option value="todos">Todos</option></select></label>
      </div>
    </section>

    <section className="price-card-list">
      {rows.map(product => {
        const d = draft(product);
        const cost = asNum(d.cost);
        const price = asNum(d.price);
        const profit = price - cost;
        const margin = price > 0 ? (profit / price) * 100 : 0;
        const markup = cost > 0 ? (profit / cost) * 100 : 0;
        const suggested = suggestedPrice(d.cost, d.margin_target);
        const health = priceHealth(d);
        const open = expandedId === product.id;
        return <article className="price-product-card" key={product.id}>
          <div className="price-card-main">
            <img src={productImageSrc(product)} alt={product.name} />
            <div className="price-card-title">
              <div className="price-title-line"><strong>{product.name}</strong><span className={priceBadgeClass(d.price_status)}>{d.price_status}</span></div>
              <small>{product.code} · {product.barcode || 'Sin barcode'} · {product.category || 'General'}{product.subcategory ? ` / ${product.subcategory}` : ''}</small>
              <div className={`health-chip health-${health.kind}`}>{health.text}</div>
            </div>
            <div className="price-card-total">
              <span>Precio actual</span>
              <strong>{money(price)}</strong>
              <small>Costo {money(cost)}</small>
            </div>
            <button type="button" className="secondary-btn price-expand-btn" onClick={()=>setExpandedId(open ? null : product.id)}>{open ? 'Cerrar edición' : 'Editar precio'}</button>
          </div>

          <div className="price-metric-strip">
            <div><span>Sugerido</span><strong>{money(suggested)}</strong></div>
            <div><span>Ganancia</span><strong className={profit < 0 ? 'danger-text' : ''}>{money(profit)}</strong></div>
            <div><span>Margen precio</span><strong className={margin < 25 ? 'warn-text' : ''}>{margin.toFixed(1)}%</strong></div>
            <div><span>Sobre costo</span><strong>{markup.toFixed(1)}%</strong></div>
          </div>

          {open && <div className="price-edit-panel">
            <div className="price-edit-grid">
              <label>Costo real<input value={d.cost} inputMode="decimal" onChange={e=>setDraft(product, 'cost', e.target.value)} /></label>
              <label>Precio venta<input value={d.price} inputMode="decimal" onChange={e=>setDraft(product, 'price', e.target.value)} /></label>
              <label>Margen objetivo %<input value={d.margin_target} inputMode="decimal" onChange={e=>setDraft(product, 'margin_target', e.target.value)} /></label>
              <label>Precio mínimo<input value={d.min_price} inputMode="decimal" onChange={e=>setDraft(product, 'min_price', e.target.value)} /></label>
              <label>Estado<select value={d.price_status} onChange={e=>setDraft(product, 'price_status', e.target.value)}>{PRICE_STATUS_OPTIONS.map(o=><option key={o}>{o}</option>)}</select></label>
            </div>
            <label className="price-note-full">Nota de control<input value={d.price_notes} onChange={e=>setDraft(product, 'price_notes', e.target.value)} placeholder="Ej.: confirmado con proveedor, revisar por promoción, precio temporal..." /></label>
            <div className="price-edit-actions">
              <button type="button" className="secondary-btn" onClick={()=>useSuggested(product)}>Usar sugerido</button>
              <button type="button" className="secondary-btn" onClick={()=>markValidated(product)}>Marcar validado</button>
              <button type="button" className="secondary-btn" onClick={()=>resetDraft(product)}>Restaurar</button>
              <button type="button" className="primary-btn" disabled={savingId===product.id} onClick={()=>savePrice(product)}>{savingId===product.id ? 'Guardando...' : 'Guardar cambios'}</button>
            </div>
          </div>}
        </article>;
      })}
      {!rows.length && <section className="card compact-card"><p className="muted">No hay productos para este filtro.</p></section>}
    </section>
  </div>;
}

function LabelsAdmin({ products = [], categories = [], subcategories = [], store }) {
  const [search, setSearch] = useState('');
  const [categoryId, setCategoryId] = useState('all');
  const [subcategoryId, setSubcategoryId] = useState('all');
  const [mode, setMode] = useState('auto');
  const [labelSize, setLabelSize] = useState('medium');
  const [labelTemplate, setLabelTemplate] = useState('commercial');
  const [labelUse, setLabelUse] = useState('auto');
  const [orientation, setOrientation] = useState('auto');
  const [composition, setComposition] = useState('auto');
  const [visualPreset, setVisualPreset] = useState('retail_auto');
  const [priceEmphasis, setPriceEmphasis] = useState('featured');
  const [qrEmphasis, setQrEmphasis] = useState('balanced');
  const [verticalDensity, setVerticalDensity] = useState('compact');
  const [sheetLayout, setSheetLayout] = useState('a4_3x7');
  const [showPrice, setShowPrice] = useState(true);
  const [showLogo, setShowLogo] = useState(true);
  const [showCodeText, setShowCodeText] = useState(true);
  const [defaultQty, setDefaultQty] = useState(1);
  const [selected, setSelected] = useState(new Set());
  const [quantities, setQuantities] = useState({});
  const [catalogUrlInput, setCatalogUrlInput] = useState(() => getSavedCatalogBaseUrl() || CONFIGURED_CATALOG_URL || '');
  const [catalogNotice, setCatalogNotice] = useState('');

  const activeProducts = useMemo(() => products.filter(p => p.active !== false && p.status !== 'Inactivo'), [products]);
  const filtered = useMemo(() => activeProducts.filter(p => {
    const q = normalizeText(search);
    const matchText = !q || [p.name, p.code, p.barcode, p.brand, p.color, p.category, p.subcategory].some(v => normalizeText(v).includes(q));
    const matchCat = categoryId === 'all' || p.category_id === categoryId || normalizeText(p.category) === normalizeText(categories.find(c=>c.id===categoryId)?.name);
    const matchSub = subcategoryId === 'all' || p.subcategory_id === subcategoryId || normalizeText(p.subcategory) === normalizeText(subcategories.find(c=>c.id===subcategoryId)?.name);
    return matchText && matchCat && matchSub;
  }), [activeProducts, search, categoryId, subcategoryId, categories, subcategories]);
  const chosen = useMemo(() => activeProducts.filter(p => selected.has(p.id)), [activeProducts, selected]);
  const basePrintable = chosen.length ? chosen : filtered;
  const visibleSubcategories = useMemo(() => categoryId === 'all' ? subcategories : subcategories.filter(s => s.parent_id === categoryId), [categoryId, subcategories]);
  const printableItems = useMemo(() => {
    const items = [];
    for (const product of basePrintable) {
      const qty = Math.max(1, Math.min(500, Number(quantities[product.id] || defaultQty || 1)));
      for (let i = 0; i < qty; i++) items.push({ product, copy: i + 1, key: `${product.id}-${i}` });
    }
    return items;
  }, [basePrintable, quantities, defaultQty]);
  const layoutProfile = labelUse === 'auto' ? inferredLabelProfile(basePrintable[0] || filtered[0] || {}) : labelUse;
  const activeOrientation = resolveLabelOrientation(orientation, layoutProfile);
  const effectiveSheetLayout = resolveAdaptiveLayout(sheetLayout, activeOrientation, layoutProfile);
  const effectiveLayout = LABEL_LAYOUTS[effectiveSheetLayout] || LABEL_LAYOUTS.a4_3x7;
  const sheetInfo = effectiveLayout.label || 'A4 · 3 columnas × 7 filas';
  const stableCatalogUrl = normalizeCatalogBaseUrl(catalogUrlInput);
  const requiresQr = basePrintable.some(product => { const profile = resolveLabelProfile(product, labelUse); const resolvedComposition = resolveLabelComposition(composition, profile, activeOrientation, effectiveLayout); const baseMode = resolveLabelCodeMode(mode, profile, effectiveLayout); const effectiveMode = resolveCompositionCodeMode(baseMode, resolvedComposition, profile, effectiveLayout); return effectiveMode === 'qr' || effectiveMode === 'both'; });

  function applyVisualPreset(key) {
    const preset = LABEL_VISUAL_PRESETS[key];
    if (!preset) return;
    setVisualPreset(key);
    setLabelUse(preset.labelUse);
    setOrientation(preset.orientation);
    setComposition(preset.composition);
    setMode(preset.mode);
    setLabelTemplate(preset.labelTemplate);
    setSheetLayout(preset.sheetLayout);
    setPriceEmphasis(preset.priceEmphasis);
    setQrEmphasis(preset.qrEmphasis);
    setVerticalDensity(preset.verticalDensity);
    if (typeof preset.showPrice === 'boolean') setShowPrice(preset.showPrice);
    if (typeof preset.showLogo === 'boolean') setShowLogo(preset.showLogo);
    if (typeof preset.showCodeText === 'boolean') setShowCodeText(preset.showCodeText);
  }
  function toggleProduct(id) { setSelected(prev => { const next = new Set(prev); next.has(id) ? next.delete(id) : next.add(id); return next; }); }
  function selectFiltered() { setSelected(new Set(filtered.map(p => p.id))); }
  function clearSelection() { setSelected(new Set()); }
  function setQty(id, value) { const qty = Math.max(1, Math.min(500, Number(value || 1))); setQuantities(prev => ({ ...prev, [id]: qty })); }
  function applyQtyToFiltered() { const qty = Math.max(1, Math.min(500, Number(defaultQty || 1))); const next = { ...quantities }; filtered.forEach(p => { next[p.id] = qty; }); setQuantities(next); }
  function applyStockQtyToFiltered() { const next = { ...quantities }; filtered.forEach(p => { next[p.id] = Math.max(1, Math.min(500, Number(p.stock || 1))); }); setQuantities(next); }
  function saveCatalogUrl() {
    if (!stableCatalogUrl) return setCatalogNotice('Ingrese la URL pública real del catálogo antes de imprimir QR.');
    try { localStorage.setItem(CATALOG_URL_STORAGE_KEY, stableCatalogUrl); } catch (_) {}
    setCatalogUrlInput(stableCatalogUrl);
    setCatalogNotice('Enlace público guardado. Los próximos QR usarán este dominio.');
  }
  function testCatalogUrl() {
    if (!stableCatalogUrl) return setCatalogNotice('Ingrese primero el enlace público del catálogo.');
    const sample = basePrintable[0] || activeProducts[0];
    window.open(sample ? catalogProductUrl(sample, stableCatalogUrl) : stableCatalogUrl, '_blank', 'noopener,noreferrer');
  }
  function printLabels() {
    if (!printableItems.length) return alert('No hay productos para imprimir.');
    if (requiresQr && !stableCatalogUrl) {
      setCatalogNotice('Para QR de catálogo configure primero una URL pública estable. No use la URL temporal de preview de Vercel.');
      return;
    }
    openLabelsPrintWindow({ items: printableItems, store, mode, showPrice, showLogo, showCodeText, sheetLayout: effectiveSheetLayout, labelStyle: labelSize, labelTemplate, labelUse, orientation: activeOrientation, composition, catalogUrl: stableCatalogUrl, priceEmphasis, qrEmphasis, verticalDensity, visualPreset });
  }
  const previewItems = printableItems.slice(0,4).length ? printableItems.slice(0,4) : filtered.slice(0,4).map(product=>({product,key:product.id}));

  return <div className="page labels-page labels-pro-page labels-v342-page">
    <div className="hero labels-pro-hero labels-v342-hero"><div><span className="eyebrow">V03.4.4 · Editor visual y plantillas por rubro</span><h1>Editor visual de etiquetas por rubro</h1><p>Use plantillas por rubro o ajuste precio, QR, orientación y composición antes de imprimir.</p></div><div className="labels-v342-hero-actions"><span className={stableCatalogUrl ? 'labels-url-state ready' : 'labels-url-state'}>{stableCatalogUrl ? 'QR público configurado' : 'QR público pendiente'}</span><button className="secondary-btn" onClick={testCatalogUrl}>Probar catálogo público</button></div></div>
    <div className="labels-pro-flow"><span><b>1</b> Filtre productos</span><i>→</i><span><b>2</b> Elija rubro y finalidad</span><i>→</i><span><b>3</b> Revise QR y PDF</span><small>Ropa y calzado: barras · Góndola: QR · Hogar: ficha con detalle</small></div>
    <section className="card labels-v342-catalog-url"><div><span className="eyebrow">Enlace público obligatorio para QR</span><h2>Dominio de catálogo para clientes</h2><p>El QR no debe usar una URL temporal ni una vista privada de Vercel. Pegue aquí el enlace público que abre el catálogo sin iniciar sesión.</p></div><div className="labels-v342-url-controls"><input value={catalogUrlInput} onChange={e=>{ setCatalogUrlInput(e.target.value); setCatalogNotice(''); }} placeholder="Ej.: https://mitienda.com/#/catalogo" /><button type="button" className="secondary-btn" onClick={saveCatalogUrl}>Guardar enlace</button><button type="button" className="primary-btn" onClick={testCatalogUrl}>Abrir prueba</button></div>{catalogNotice && <small className={stableCatalogUrl ? 'labels-v342-ok' : 'labels-v342-warn'}>{catalogNotice}</small>}<small className="labels-v342-help">Use siempre la URL de producción. Si al abrir la prueba aparece inicio de sesión de Vercel, debe quitar la protección del dominio público antes de imprimir etiquetas con QR.</small></section>
    <div className="labels-pro-layout">
      <section className="card labels-pro-filter"><h2>Productos</h2><label>Buscar producto<input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Nombre, código, barcode, marca o color" /></label><div className="form-split"><label>Categoría<select value={categoryId} onChange={e=>{ setCategoryId(e.target.value); setSubcategoryId('all'); }}><option value="all">Todas</option>{categories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label><label>Subcategoría<select value={subcategoryId} onChange={e=>setSubcategoryId(e.target.value)}><option value="all">Todas</option>{visibleSubcategories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label></div><div className="labels-pro-stats"><Kpi label="Filtrados" value={filtered.length} helper="productos" /><Kpi label="Seleccionados" value={chosen.length || filtered.length} helper={chosen.length ? 'manual' : 'por filtro'} /><Kpi label="Etiquetas" value={printableItems.length} helper="a imprimir" /></div><div className="button-row"><button className="secondary-btn" onClick={selectFiltered}>Seleccionar filtrados</button><button className="secondary-btn" onClick={clearSelection}>Limpiar selección</button></div></section>
      <section className="card labels-pro-design labels-v342-design"><h2>Rubro, finalidad y diseño</h2><div className="labels-v344-preset-grid">{Object.entries(LABEL_VISUAL_PRESETS).map(([key, preset])=><button type="button" key={key} onClick={()=>applyVisualPreset(key)} className={`labels-v344-preset ${visualPreset===key?'active':''}`}><strong>{preset.label}</strong><small>{preset.help}</small></button>)}</div><div className="labels-v342-use-grid">{Object.entries(LABEL_USE_INFO).map(([key, info])=><button type="button" key={key} onClick={()=>setLabelUse(key)} className={`labels-v342-use-option ${labelUse===key?'active':''}`}><strong>{info.label}</strong><small>{info.help}</small></button>)}</div><div className="label-template-grid">{Object.entries(LABEL_TEMPLATE_INFO).map(([key, info])=><button type="button" key={key} onClick={()=>setLabelTemplate(key)} className={`label-template-option ${labelTemplate===key?'active':''}`}><strong>{info.label}</strong><small>{info.help}</small></button>)}</div><div className="labels-v343-adaptive-grid"><div><span className="eyebrow">Orientación</span><div className="labels-v343-option-row">{[['auto','Automática'],['horizontal','Horizontal'],['vertical','Vertical']].map(([key,label])=><button type="button" key={key} onClick={()=>setOrientation(key)} className={orientation===key?'active':''}><strong>{label}</strong><small>{key==='auto'?'Ropa vertical; calzado y hogar horizontal':key==='vertical'?'Colgante para prendas y accesorios':'Caja, góndola y mostrador'}</small></button>)}</div></div><div><span className="eyebrow">Composición</span><div className="labels-v343-option-row composition">{[['auto','Automática'],['classic','Clásica'],['two-column','Dos columnas'],['price','Precio protagonista'],['qr','QR protagonista'],['barcode','Barras POS']].map(([key,label])=><button type="button" key={key} onClick={()=>setComposition(key)} className={composition===key?'active':''}><strong>{label}</strong></button>)}</div></div></div><div className="form-split"><label>Tipo de código<select value={mode} onChange={e=>setMode(e.target.value)}><option value="auto">Automático según perfil</option><option value="both">QR catálogo + barras POS</option><option value="qr">Solo QR catálogo</option><option value="barcode">Solo código de barras</option></select></label><label>Formato<select value={sheetLayout} onChange={e=>setSheetLayout(e.target.value)}><option value="a4_2x6">A4: 2 columnas × 6 filas (90 × 43 mm)</option><option value="a4_3x7">A4: 3 columnas × 7 filas (58 × 37 mm)</option><option value="a4_4x8">A4: 4 columnas × 8 filas (43 × 32 mm)</option><option value="a4_vertical_2x3">A4 vertical premium: 2 columnas × 3 filas (90 × 84 mm)</option><option value="a4_vertical_3x4">A4 vertical retail: 3 columnas × 4 filas (58 × 63 mm)</option><option value="a4_vertical_3x3">A4 vertical largo: 3 columnas × 3 filas (52 × 80 mm)</option><option value="a4_vertical_4x4">A4 vertical compacto: 4 columnas × 4 filas (43 × 63 mm)</option><option value="a4_vertical_5x4">A4 vertical fino: 5 columnas × 4 filas (33 × 63 mm)</option><option value="roll_1col">Rollo térmico: 1 columna (60 × 40 mm)</option></select></label></div><div className="form-split"><label>Escala de contenido<select value={labelSize} onChange={e=>setLabelSize(e.target.value)}><option value="small">Compacta</option><option value="medium">Equilibrada</option><option value="large">Destacada</option></select></label><label>Cantidad rápida<input type="number" min="1" max="500" value={defaultQty} onChange={e=>setDefaultQty(e.target.value)} /></label></div><div className="labels-v344-editor"><div><span className="eyebrow">Editor visual</span><h3>Ajuste fino antes de imprimir</h3></div><label>Precio<select value={priceEmphasis} onChange={e=>setPriceEmphasis(e.target.value)}><option value="normal">Normal</option><option value="featured">Destacado</option><option value="max">Máximo</option></select></label><label>QR<select value={qrEmphasis} onChange={e=>setQrEmphasis(e.target.value)}><option value="small">Pequeño</option><option value="balanced">Equilibrado</option><option value="large">Grande</option><option value="max">Máximo</option></select></label><label>Espacio vertical<select value={verticalDensity} onChange={e=>setVerticalDensity(e.target.value)}><option value="compact">Compacto</option><option value="relaxed">Normal</option></select></label></div><div className="labels-pro-checks"><label className="check-row"><input type="checkbox" checked={showPrice} onChange={e=>setShowPrice(e.target.checked)} /> Precio destacado</label><label className="check-row"><input type="checkbox" checked={showLogo} onChange={e=>setShowLogo(e.target.checked)} /> Marca Clomar Store</label><label className="check-row"><input type="checkbox" checked={showCodeText} onChange={e=>setShowCodeText(e.target.checked)} /> Código escrito</label></div><div className="button-row"><button className="secondary-btn" onClick={applyQtyToFiltered}>Aplicar cantidad</button><button className="secondary-btn" onClick={applyStockQtyToFiltered}>Usar stock como cantidad</button></div><button className="primary-btn labels-pro-print" onClick={printLabels}>Imprimir / Guardar PDF · {printableItems.length} etiquetas</button><p className="muted">Formato adaptativo: <strong>{sheetInfo}</strong>. Orientación: <strong>{orientation === 'auto' ? `automática: ${activeOrientation}` : activeOrientation}</strong>. En impresión: A4, escala 100 %, márgenes Ninguno y encabezados/pies desactivados.</p></section>
    </div>
    <section className="card labels-pro-products"><div className="section-row"><div><span className="eyebrow">Selección de impresión</span><h2>Productos para etiquetas</h2></div><span className="muted">El perfil automático adapta talla, color, medida o marca según el rubro.</span></div><div className="labels-pro-product-grid">{filtered.map(p=>{const code=productScanCode(p); const selectedNow=selected.has(p.id); const profile=inferredLabelProfile(p); return <article key={p.id} className={`labels-pro-product ${selectedNow?'selected':''}`}><label className="labels-pro-product-check"><input type="checkbox" checked={selectedNow} onChange={()=>toggleProduct(p.id)} /> Seleccionar</label><img src={productImageSrc(p)} alt={p.name}/><div className="labels-pro-product-body"><strong>{p.name}</strong><small>{labelMetaFor(p, profile)}</small><div><span className={`labels-v342-profile-tag ${profile}`}>{LABEL_USE_INFO[profile]?.label || 'General'}</span><em>{productPriceStatus(p)==='Validado' ? money(p.price) : 'Precio pendiente'}</em></div></div><label className="labels-pro-qty"><small>Copias</small><input type="number" min="1" max="500" value={quantities[p.id] || defaultQty} onChange={e=>setQty(p.id,e.target.value)} /></label></article>})}{!filtered.length&&<p className="muted">No hay productos con esos filtros.</p>}</div></section>
    <section className="card labels-pro-preview labels-v342-preview"><div className="section-row"><div><span className="eyebrow">Vista previa contextual</span><h2>{LABEL_USE_INFO[labelUse]?.label || 'Etiqueta inteligente'}</h2></div><span className={stableCatalogUrl ? 'preview-link' : 'preview-link warning'}>{stableCatalogUrl ? 'QR enlazado a catálogo público' : 'Configure URL pública para QR'}</span></div><div className={`labels-pro-preview-grid template-${resolveLabelTemplate(labelTemplate,labelUse)}`}>{previewItems.map(({product,key})=>{const code=productScanCode(product);const profile=resolveLabelProfile(product,labelUse);const previewOrientation=activeOrientation;const previewComposition=resolveLabelComposition(composition,profile,previewOrientation,effectiveLayout);const template=previewComposition==='qr'?'showcase':resolveLabelTemplate(labelTemplate,labelUse);const baseMode=resolveLabelCodeMode(mode,profile,effectiveLayout);const effectiveMode=resolveCompositionCodeMode(baseMode,previewComposition,profile,effectiveLayout);const ready=productPriceStatus(product)==='Validado'&&Number(product.price||0)>0;return <article className={`labels-pro-card template-${template} profile-${profile} orientation-${previewOrientation} composition-${previewComposition} price-${priceEmphasis} qr-${qrEmphasis} vertical-${verticalDensity}`} key={key}>{showLogo&&<div className="label-pro-brand"><img src={APP_ICON} alt=""/><span>{store?.name||'Clomar Store'}</span></div>}<div className="label-pro-name">{product.name}</div><div className="label-pro-meta">{labelMetaFor(product, profile)}</div>{showPrice&&<div className={`label-pro-price ${ready?'':'pending'}`}><small>{profile==='inventory'?'PRECIO REF.':ready?'PRECIO':'REVISAR'}</small><b>{ready?money(product.price):'Precio pendiente'}</b></div>}<div className={`label-pro-codes mode-${effectiveMode}`}>{(effectiveMode==='qr'||effectiveMode==='both')&&<div><img src={qrUrl(catalogQrValue(product, stableCatalogUrl))} alt="QR catálogo"/><small>{profile==='gondola'?'Escanea':'Catálogo'}</small></div>}{(effectiveMode==='barcode'||effectiveMode==='both')&&<div className="label-pro-barcode"><BarcodeSVG value={code}/>{showCodeText&&<small>{code}</small>}</div>}</div></article>})}</div></section>
  </div>;
}

function ToolsAdmin({ profile, products = [], categories = [], subcategories = [], reloadProducts, reloadCustomers }) {
  const [confirmText, setConfirmText] = useState('');
  const [resetting, setResetting] = useState(false);
  const [deleteImages, setDeleteImages] = useState(false);
  const [fileName, setFileName] = useState('');
  const [previewRows, setPreviewRows] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [importMode, setImportMode] = useState('Completo');
  const [exporting, setExporting] = useState(false);

  const categoryByName = useMemo(() => { const map = new Map(); categories.forEach(c => map.set(normalizeText(c.name), c)); return map; }, [categories]);
  const subcategoryByParentAndName = useMemo(() => { const map = new Map(); subcategories.forEach(sc => map.set(`${sc.parent_id}|${normalizeText(sc.name)}`, sc)); return map; }, [subcategories]);
  const productByCode = useMemo(() => { const map = new Map(); products.forEach(p => { const key = normalizeText(p.code || ''); if (key) map.set(key, p); }); return map; }, [products]);
  const importOverview = useMemo(() => {
    const errors = rawRows.filter(r => r.errors.length).length;
    const valid = rawRows.length - errors;
    const existing = rawRows.filter(r => productByCode.has(normalizeText(r.code))).length;
    return { errors, valid, existing, newRows: Math.max(0, valid - existing) };
  }, [rawRows, productByCode]);

  function normalizeImportedRow(row, idx) {
    const normalized = {};
    Object.entries(row || {}).forEach(([key, value]) => { const field = canonicalProductField(normalizeHeader(key)); normalized[field] = typeof value === 'string' ? value.trim() : value; });
    const categoryName = String(normalized.category || '').trim();
    const subcategoryName = String(normalized.subcategory || '').trim();
    const category = categoryByName.get(normalizeText(categoryName));
    const subcategory = category ? subcategoryByParentAndName.get(`${category.id}|${normalizeText(subcategoryName)}`) : null;
    const generatedCode = `${categoryPrefix(categoryName)}-${String(products.length + idx + 1).padStart(6, '0')}`;
    const code = String(normalized.code || '').trim() || generatedCode;
    const barcode = String(normalized.barcode || '').trim() || code;
    const priceStatusImported = normalizePriceStatus(normalized.price_status);
    const marginTarget = parseMoneyLike(normalized.margin_target || 50);
    const minPrice = parseMoneyLike(normalized.min_price || 0);
    const errors = [];
    if (!String(normalized.name || '').trim()) errors.push('Falta nombre');
    if (!categoryName) errors.push('Falta categoría');
    if (categoryName && !category) errors.push(`Categoría no existe: ${categoryName}`);
    if (subcategoryName && category && !subcategory) errors.push(`Subcategoría no existe: ${subcategoryName}`);
    if (priceStatusImported === 'Validado' && parseMoneyLike(normalized.price) <= 0) errors.push('Precio validado debe ser mayor a 0');
    if (priceStatusImported === 'Validado' && minPrice > 0 && parseMoneyLike(normalized.price) < minPrice) errors.push('Precio validado menor al mínimo permitido');
    return { rowNumber: idx + 2, code, barcode, name: String(normalized.name || '').trim(), category: category?.name || categoryName || 'General', subcategory: subcategory?.name || subcategoryName || '', category_id: category?.id || null, subcategory_id: subcategory?.id || null, brand: String(normalized.brand || '').trim(), size: String(normalized.size || '').trim(), color: String(normalized.color || '').trim(), description: String(normalized.description || '').trim(), cost: parseMoneyLike(normalized.cost), price: parseMoneyLike(normalized.price), price_status: priceStatusImported, margin_target: marginTarget || 50, min_price: minPrice, price_notes: String(normalized.price_notes || '').trim(), stock: parseMoneyLike(normalized.stock), stock_min: parseMoneyLike(normalized.stock_min || 1), image_url: String(normalized.image_url || '').trim(), active: boolActive(normalized.active ?? true), errors };
  }

  async function parseFile(e) {
    const file = e.target.files?.[0]; setImportResult(null); setPreviewRows([]); setRawRows([]); if (!file) return; setFileName(file.name);
    try { const buffer = await file.arrayBuffer(); const workbook = XLSX.read(buffer, { type: 'array' }); const sheet = workbook.Sheets[workbook.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }); const normalized = rows.map((r, idx) => normalizeImportedRow(r, idx)); const codeCounts = normalized.reduce((acc, row) => { const key = normalizeText(row.code || ''); if (key) acc.set(key, (acc.get(key) || 0) + 1); return acc; }, new Map()); normalized.forEach(row => { if (codeCounts.get(normalizeText(row.code || '')) > 1) row.errors.push('Código interno duplicado en el Excel'); }); setRawRows(normalized); setPreviewRows(normalized.slice(0, 30)); } catch (error) { alert(error.message || 'No se pudo leer el archivo Excel.'); }
  }

  async function importProducts() {
    if (!rawRows.length) return alert('Primero carga un Excel.');
    const invalid = rawRows.filter(r => r.errors.length);
    if (invalid.length) return alert(`Hay ${invalid.length} filas con errores. Corrige el Excel antes de importar.`);
    const storeId = profile?.store_id || DEFAULT_STORE_ID;
    if (importMode !== 'Completo') {
      const missing = rawRows.filter(r => !productByCode.has(normalizeText(r.code)));
      if (missing.length) return alert(`En modo “${importMode}” todos los códigos deben existir. Hay ${missing.length} código(s) que no se encontraron.`);
    }
    setImporting(true);
    try {
      let count = 0;
      let created = 0;
      let updated = 0;
      if (importMode === 'Solo stock') {
        for (const row of rawRows) {
          const { error } = await supabase.from('products').update({ stock: row.stock, stock_min: row.stock_min, updated_at: new Date().toISOString() }).eq('code', row.code).eq('store_id', storeId);
          if (error) throw error;
          count += 1; updated += 1;
        }
      } else if (importMode === 'Solo precios') {
        for (const row of rawRows) {
          const { error } = await supabase.from('products').update({ cost: row.cost, price: row.price, price_status: row.price_status, margin_target: row.margin_target, min_price: row.min_price, price_notes: row.price_notes, price_updated_at: new Date().toISOString(), price_updated_by: profile?.id || null, updated_at: new Date().toISOString() }).eq('code', row.code).eq('store_id', storeId);
          if (error) throw error;
          count += 1; updated += 1;
        }
      } else {
        const payloads = rawRows.map(r => ({ code: r.code, barcode: r.barcode, name: r.name, category: r.category, subcategory: r.subcategory, category_id: r.category_id, subcategory_id: r.subcategory_id, brand: r.brand, size: r.size, color: r.color, description: r.description, cost: r.cost, price: r.price, price_status: r.price_status, margin_target: r.margin_target, min_price: r.min_price, price_notes: r.price_notes, price_updated_at: new Date().toISOString(), price_updated_by: profile?.id || null, stock: r.stock, stock_min: r.stock_min, image_url: r.image_url, image_path: '', active: r.active, status: r.active ? 'Activo' : 'Inactivo', store_id: storeId, created_by: profile?.id || null, updated_at: new Date().toISOString() }));
        const chunks = []; for (let i = 0; i < payloads.length; i += 200) chunks.push(payloads.slice(i, i + 200));
        for (const chunk of chunks) { const { error } = await supabase.from('products').upsert(chunk, { onConflict: 'code' }); if (error) throw error; count += chunk.length; }
        updated = rawRows.filter(r => productByCode.has(normalizeText(r.code))).length;
        created = Math.max(0, count - updated);
      }
      await supabase.from('product_import_batches').insert({ store_id: storeId, user_id: profile?.id || null, file_name: fileName, total_rows: rawRows.length, imported_rows: count, status: 'Importado', notes: `V02.2K · ${importMode} · nuevos ${created} · actualizados ${updated}` }).then(()=>{});
      setImportResult({ count, created, updated, mode: importMode }); setPreviewRows([]); setRawRows([]); setFileName(''); await reloadProducts?.(); alert(`Importación completada: ${count} producto(s).`);
    } catch (error) { alert(error.message || 'No se pudo importar productos.'); } finally { setImporting(false); }
  }

  async function exportGeneralBackup() {
    if (!hasSupabaseConfig) return alert('Configura Supabase para exportar datos reales.');
    setExporting(true);
    try {
      const storeId = profile?.store_id || DEFAULT_STORE_ID;
      const tables = ['sales','sale_items','cash_movements','credit_payments','stock_movements','customers','product_price_history'];
      const data = { exported_at: new Date().toISOString(), store_id: storeId, products, categories, subcategories };
      for (const table of tables) { const { data: rows } = await supabase.from(table).select('*').eq('store_id', storeId).limit(10000); data[table] = rows || []; }
      const json = JSON.stringify(data, null, 2);
      downloadText(`backup-clomar-${todayISO()}.json`, json);
      const wb = XLSX.utils.book_new();
      Object.entries(data).forEach(([key, value]) => { if (Array.isArray(value)) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(value), key.slice(0,31)); });
      XLSX.writeFile(wb, `backup-clomar-${todayISO()}.xlsx`);
    } catch (error) { alert(error.message || 'No se pudo exportar backup.'); } finally { setExporting(false); }
  }

  async function deleteProductImagesFromStorage() { try { const { data } = await supabase.storage.from('product-images').list('', { limit: 1000 }); const files = (data || []).map(f => f.name).filter(Boolean); if (files.length) await supabase.storage.from('product-images').remove(files); } catch (_) {} }
  async function controlledReset() {
    if (profile?.role !== 'dueno') return alert('Solo el dueño puede reiniciar datos.');
    if (confirmText !== 'REINICIAR CLOMAR') return alert('Debes escribir exactamente: REINICIAR CLOMAR');
    if (!confirm('Esto borrará productos, clientes, ventas, caja, créditos y movimientos de prueba. Se conservan usuarios, tienda y categorías. ¿Continuar?')) return;
    setResetting(true);
    try { const { error } = await supabase.rpc('clomar_reset_operational_data', { confirm_text: confirmText, store_uuid: profile?.store_id || DEFAULT_STORE_ID }); if (error) throw error; if (deleteImages) await deleteProductImagesFromStorage(); setConfirmText(''); await reloadProducts?.(); await reloadCustomers?.(); alert('Reinicio controlado completado.'); } catch (error) { alert(error.message || 'No se pudo reiniciar. Verifica que ejecutaste el SQL V01.10.'); } finally { setResetting(false); }
  }

  return (
    <div className="page tools-pro-page">
      <div className="hero compact-hero"><h1>🛠️ Herramientas y backup</h1><p>Importación, exportación general, respaldo y reinicio controlado.</p></div>
      <div className="tool-grid">
        <section className="card compact-card backup-card"><h3>Backup / exportación general</h3><p className="muted">Descarga un respaldo JSON y Excel con productos, ventas, caja, créditos, clientes, movimientos y precios.</p><div className="backup-points"><div><strong>{products.length}</strong><small>productos</small></div><div><strong>{categories.length}</strong><small>categorías</small></div><div><strong>{subcategories.length}</strong><small>subcategorías</small></div></div><button className="primary-btn" disabled={exporting} onClick={exportGeneralBackup}>{exporting ? 'Exportando...' : 'Exportar backup JSON + Excel'}</button><p className="muted">Recomendación: hacer backup antes de cada actualización grande.</p></section>
        <section className="card compact-card danger-zone-card"><h3>Reinicio controlado</h3><p className="muted">Borra datos operativos de prueba y conserva usuarios, roles, tienda, logo, categorías y subcategorías.</p><div className="reset-keep-delete"><div><strong>Conserva</strong><small>Usuarios · Roles · Tienda · Logo · Categorías</small></div><div><strong>Borra</strong><small>Productos · Ventas · Caja · Créditos · Clientes · Movimientos</small></div></div><label className="check-row"><input type="checkbox" checked={deleteImages} onChange={e=>setDeleteImages(e.target.checked)} /> Borrar también imágenes del bucket product-images</label><label>Confirmación obligatoria<input value={confirmText} onChange={e=>setConfirmText(e.target.value)} placeholder="Escribe REINICIAR CLOMAR" /></label><button className="danger-btn" disabled={resetting} onClick={controlledReset}>{resetting ? 'Reiniciando...' : 'Reiniciar datos de prueba'}</button></section>
        <section className="card compact-card import-pro-card"><div className="import-pro-head"><div><span className="eyebrow">Asistente de carga</span><h3>Importar productos desde Excel</h3><p className="muted">Valida códigos, detecta duplicados y elige si deseas importar todo, actualizar solo stock o solo precios.</p></div><a className="secondary-btn" href="/plantilla_productos_clomar_v0112.xlsx" download>Descargar plantilla</a></div><div className="import-steps"><span>1. Plantilla</span><span>2. Revisar</span><span>3. Importar</span></div><div className="import-mode-picker"><button type="button" className={importMode === 'Completo' ? 'active' : ''} onClick={()=>setImportMode('Completo')}>Completo</button><button type="button" className={importMode === 'Solo stock' ? 'active' : ''} onClick={()=>setImportMode('Solo stock')}>Solo stock</button><button type="button" className={importMode === 'Solo precios' ? 'active' : ''} onClick={()=>setImportMode('Solo precios')}>Solo precios</button></div><p className="import-mode-note">{importMode === 'Completo' ? 'Crea productos nuevos y actualiza los que coincidan por código interno.' : importMode === 'Solo stock' ? 'Actualiza stock y stock mínimo. Todos los códigos del Excel deben existir.' : 'Actualiza costo, precio y validación. Todos los códigos del Excel deben existir.'}</p><label className="import-file-field">Seleccionar archivo Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} /></label>{fileName && <div className="info-box">Archivo cargado: <strong>{fileName}</strong> · Filas leídas: {rawRows.length}</div>}{rawRows.length > 0 && <div className="import-summary import-pro-summary"><Kpi label="Filas" value={rawRows.length} helper="leídas" /><Kpi label="Nuevos" value={importOverview.newRows} helper="por crear" /><Kpi label="Actualizar" value={importOverview.existing} helper="ya existen" /><Kpi label="Errores" value={importOverview.errors} helper="corregir antes" /></div>}{previewRows.length > 0 && <div className="preview-table-wrap"><table className="preview-table"><thead><tr><th>Fila</th><th>Código</th><th>Producto</th><th>Categoría</th><th>Subcategoría</th><th>Precio</th><th>Estado precio</th><th>Stock</th><th>Acción</th><th>Errores</th></tr></thead><tbody>{previewRows.map(r => <tr key={r.rowNumber} className={r.errors.length ? 'row-error' : ''}><td>{r.rowNumber}</td><td>{r.code}</td><td>{r.name}</td><td>{r.category}</td><td>{r.subcategory}</td><td>{money(r.price)}</td><td>{r.price_status}</td><td>{r.stock}</td><td>{productByCode.has(normalizeText(r.code)) ? 'Actualizar' : importMode === 'Completo' ? 'Crear' : 'No existe'}</td><td>{r.errors.length ? r.errors.join('; ') : 'OK'}</td></tr>)}</tbody></table></div>}<button className="primary-btn" disabled={!rawRows.length || importing || rawRows.some(r=>r.errors.length)} onClick={importProducts}>{importing ? 'Importando...' : importMode === 'Completo' ? 'Importar productos' : importMode === 'Solo stock' ? 'Actualizar stock' : 'Actualizar precios'}</button>{importResult && <div className="success-box">Importación completada: {importResult.count} producto(s). Nuevos: {importResult.created} · Actualizados: {importResult.updated} · Modo: {importResult.mode}.</div>}</section>
      </div>
    </div>
  );
}

function StoreSettings({ store, reloadProfile }) {
  const [form, setForm] = useState(store || {});
  const [saving, setSaving] = useState(false);

  useEffect(() => { setForm(store || {}); }, [store?.id]);

  async function save(e) {
    e.preventDefault();
    setSaving(true);
    const payload = {
      name: form.name || 'Clomar Store Pro',
      ruc: form.ruc || '',
      address: form.address || '',
      phone: form.phone || '',
      whatsapp_number: normalizeWhatsappNumber(form.whatsapp_number || form.phone || CATALOG_WHATSAPP_FALLBACK),
      email: form.email || '',
      logo_url: form.logo_url || '',
      updated_at: new Date().toISOString(),
    };
    const { error } = await supabase.from('stores').update(payload).eq('id', store?.id || DEFAULT_STORE_ID);
    setSaving(false);
    if (error) return alert(error.message || 'No se pudo actualizar la tienda.');
    await reloadProfile();
    alert('Datos de tienda actualizados.');
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🏪 Configuración de tienda</h1><p>Datos base para usuarios, tickets, comprobantes y catálogo.</p></div>
      <form className="card form-grid" onSubmit={save}>
        <label>Nombre comercial<input value={form.name || ''} onChange={e=>setForm({...form, name:e.target.value})} placeholder="Clomar Store Pro" /></label>
        <label>RUC<input value={form.ruc || ''} onChange={e=>setForm({...form, ruc:e.target.value})} placeholder="RUC de la tienda" /></label>
        <label>Dirección<input value={form.address || ''} onChange={e=>setForm({...form, address:e.target.value})} placeholder="Dirección" /></label>
        <label>Teléfono de tienda<input value={form.phone || ''} onChange={e=>setForm({...form, phone:e.target.value})} placeholder="Celular" /></label>
        <label>WhatsApp oficial del catálogo<input value={form.whatsapp_number || ''} onChange={e=>setForm({...form, whatsapp_number:e.target.value})} placeholder="51931709871" /></label>
        <label>Correo<input value={form.email || ''} onChange={e=>setForm({...form, email:e.target.value})} placeholder="correo@tienda.com" /></label>
        <label>Logo URL<input value={form.logo_url || ''} onChange={e=>setForm({...form, logo_url:e.target.value})} placeholder="Opcional: pega una URL pública del logo" /></label>
        <div className="brand-preview">
          <span>Vista de marca</span>
          <img src={form.logo_url || APP_LOGO_FULL} alt="Logo Clomar Store" />
        </div>
        <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar tienda'}</button>
      </form>
    </div>
  );
}



/* =========================================================
   V03.1 — Catálogo público, pedidos web y WhatsApp comercial
   El catálogo solo utiliza RPC públicas limitadas. No expone
   costo, margen, stock exacto ni datos internos del ERP.
   ========================================================= */
function CatalogAdmin({ products = [], profile, store, reload }) {
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('todos');
  const [drafts, setDrafts] = useState({});
  const [savingId, setSavingId] = useState('');
  const [copied, setCopied] = useState(false);
  const [editingId, setEditingId] = useState('');
  const [selectedIds, setSelectedIds] = useState([]);
  const [batchWorking, setBatchWorking] = useState('');
  const [catalogPage, setCatalogPage] = useState(1);

  const rows = useMemo(() => products.filter((p) => {
    const search = normalizeText(query);
    const matchText = !search || [p.name, p.code, p.barcode, p.category, p.brand, p.color, p.size].some(v => normalizeText(v).includes(search));
    const matchFilter = filter === 'todos'
      || (filter === 'publicados' && p.public_visible && p.catalog_status === 'Publicado')
      || (filter === 'borradores' && (!p.public_visible || p.catalog_status === 'Borrador'))
      || (filter === 'agotados' && asNum(p.stock) <= 0)
      || (filter === 'pendientes_precio' && (productPriceStatus(p) !== 'Validado' || asNum(p.price) <= 0));
    return matchText && matchFilter;
  }), [products, query, filter]);

  const CATALOG_PAGE_SIZE = 18;
  const catalogPageCount = Math.max(1, Math.ceil(rows.length / CATALOG_PAGE_SIZE));
  const safeCatalogPage = Math.min(catalogPage, catalogPageCount);
  const pagedRows = rows.slice((safeCatalogPage - 1) * CATALOG_PAGE_SIZE, safeCatalogPage * CATALOG_PAGE_SIZE);
  useEffect(() => { setCatalogPage(1); }, [query, filter]);

  const draftFor = (p) => drafts[p.id] || {
    public_visible: Boolean(p.public_visible),
    catalog_status: p.catalog_status || (p.public_visible ? 'Publicado' : 'Borrador'),
    catalog_featured: Boolean(p.catalog_featured),
    catalog_description: p.catalog_description || '',
    catalog_position: String(p.catalog_position ?? 999),
  };
  const updateDraft = (p, patch) => setDrafts(prev => ({ ...prev, [p.id]: { ...draftFor(p), ...patch } }));
  const publishedCount = products.filter(p => p.public_visible && p.catalog_status === 'Publicado').length;
  const shareUrl = catalogBaseUrl();
  const readyForPublic = (p) => p.active !== false && p.status === 'Activo' && productPriceStatus(p) === 'Validado' && asNum(p.price) > 0;

  async function saveProductCatalog(p) {
    const draft = draftFor(p);
    if (draft.public_visible && draft.catalog_status === 'Publicado' && !readyForPublic(p)) {
      return alert('Para publicar, el producto debe estar activo y tener un precio validado mayor a cero. Corrija el precio antes de publicar.');
    }
    setSavingId(p.id);
    const payload = {
      public_visible: Boolean(draft.public_visible),
      catalog_status: draft.public_visible ? (draft.catalog_status || 'Borrador') : 'Borrador',
      catalog_featured: Boolean(draft.catalog_featured),
      catalog_description: String(draft.catalog_description || '').trim(),
      catalog_position: Math.max(0, Math.trunc(asNum(draft.catalog_position || 999))),
      catalog_updated_at: new Date().toISOString(),
    };
    const { error } = await supabase.from('products').update(payload).eq('id', p.id).eq('store_id', profile?.store_id || DEFAULT_STORE_ID);
    setSavingId('');
    if (error) return alert(`No se pudo actualizar el catálogo: ${error.message}`);
    setEditingId('');
    await reload?.();
  }

  async function copyCatalogLink() {
    try { await navigator.clipboard.writeText(catalogBaseUrl()); setCopied(true); setTimeout(() => setCopied(false), 1600); }
    catch (_) { window.prompt('Copia este enlace del catálogo:', catalogBaseUrl()); }
  }

  function toggleProductSelection(productId) {
    setSelectedIds(prev => prev.includes(productId) ? prev.filter(id => id !== productId) : [...prev, productId]);
  }

  async function applyBatch(action) {
    if (!selectedIds.length) return;
    const selectedProducts = products.filter(p => selectedIds.includes(p.id));
    if (action === 'publish') {
      const invalid = selectedProducts.filter(p => !readyForPublic(p));
      if (invalid.length) return alert(`${invalid.length} producto(s) no tiene(n) precio validado o están inactivos. Corrija esos productos antes de publicar.`);
    }
    setBatchWorking(action);
    const now = new Date().toISOString();
    const payload = action === 'publish'
      ? { public_visible: true, catalog_status: 'Publicado', catalog_updated_at: now }
      : action === 'hide'
        ? { public_visible: false, catalog_status: 'Oculto', catalog_updated_at: now }
        : { catalog_featured: true, catalog_updated_at: now };
    const { error } = await supabase.from('products').update(payload).in('id', selectedIds).eq('store_id', profile?.store_id || DEFAULT_STORE_ID);
    setBatchWorking('');
    if (error) return alert(`No se pudo aplicar la acción: ${error.message}`);
    setSelectedIds([]);
    await reload?.();
  }

  return (
    <div className="page catalog-admin-page catalog-v321-page">
      <div className="hero compact-hero catalog-admin-hero catalog-v321-hero">
        <div><span className="eyebrow">Canal comercial</span><h1>Catálogo público</h1><p>Controle qué productos ve el cliente. Publicar no descuenta stock ni modifica precios.</p></div>
        <div className="catalog-link-actions"><div className="catalog-share-address"><span>Enlace para compartir con clientes</span><code>{shareUrl}</code></div><button type="button" className="secondary-btn" onClick={copyCatalogLink}>{copied ? '✓ Enlace copiado' : 'Copiar enlace'}</button><button type="button" className="primary-btn" onClick={() => window.open(shareUrl, '_blank', 'noopener,noreferrer')}>Abrir catálogo</button></div>
      </div>
      <section className="catalog-workflow-strip"><span>1. Valide precio</span><span>2. Active publicación</span><span>3. Revise vista pública</span><span>4. Atienda pedidos por WhatsApp</span></section>
      <section className="card compact-card catalog-summary-card">
        <div className="catalog-summary-grid"><Kpi label="Publicados" value={publishedCount} helper="visibles para clientes" /><Kpi label="Borradores" value={products.filter(p => !p.public_visible || p.catalog_status === 'Borrador').length} helper="aún no visibles" /><Kpi label="Agotados" value={products.filter(p => p.public_visible && asNum(p.stock) <= 0).length} helper="se muestran como agotados" /><Kpi label="WhatsApp" value="Canal activo" helper={normalizeWhatsappNumber(store?.whatsapp_number || store?.phone)} /></div>
      </section>
      <section className="card compact-card catalog-filter-card catalog-v321-filter">
        <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar producto, código, marca o color..." /></div>
        <div className="catalog-filter-tabs">{[['todos','Todos'],['publicados','Publicados'],['borradores','Borradores'],['agotados','Agotados'],['pendientes_precio','Precio pendiente']].map(([key,label]) => <button key={key} type="button" className={filter===key?'active':''} onClick={()=>setFilter(key)}>{label}</button>)}</div>
      </section>
      <section className="catalog-batch-bar"><div><strong>{selectedIds.length ? `${selectedIds.length} seleccionado(s)` : `${rows.length} producto(s) encontrados`}</strong><small>{selectedIds.length ? 'Aplique una acción masiva sin abrir cada ficha.' : 'Seleccione productos para publicar, ocultar o destacar varios a la vez.'}</small></div><div><button type="button" className="secondary-btn" disabled={!selectedIds.length || batchWorking} onClick={() => applyBatch('publish')}>{batchWorking === 'publish' ? 'Publicando...' : 'Publicar seleccionados'}</button><button type="button" className="secondary-btn" disabled={!selectedIds.length || batchWorking} onClick={() => applyBatch('hide')}>{batchWorking === 'hide' ? 'Ocultando...' : 'Ocultar'}</button><button type="button" className="secondary-btn" disabled={!selectedIds.length || batchWorking} onClick={() => applyBatch('feature')}>{batchWorking === 'feature' ? 'Guardando...' : 'Destacar'}</button>{selectedIds.length > 0 && <button type="button" className="catalog-clear-selection" onClick={() => setSelectedIds([])}>Limpiar</button>}</div></section>
      <div className="catalog-results-line"><strong>{rows.length} producto(s)</strong><span>Edite solo el producto necesario o use acciones masivas para ahorrar tiempo.</span></div>
      <section className="catalog-admin-list catalog-v321-list">
        {pagedRows.map(p => {
          const draft = draftFor(p);
          const availability = asNum(p.stock) <= 0 ? 'Agotado' : asNum(p.stock) <= asNum(p.stock_min) ? 'Últimas unidades' : 'Disponible';
          const publicNow = Boolean(p.public_visible) && p.catalog_status === 'Publicado';
          const isEditing = editingId === p.id;
          const isSelected = selectedIds.includes(p.id);
          return <article className={`card catalog-admin-product catalog-v321-product ${isEditing ? 'is-editing' : ''} ${isSelected ? 'is-selected' : ''}`} key={p.id}>
            <label className="catalog-select-product"><input type="checkbox" checked={isSelected} onChange={() => toggleProductSelection(p.id)} /><span>Seleccionar</span></label>
            <div className="catalog-admin-product-main"><img src={productImageSrc(p)} alt={p.name}/><div className="catalog-product-content"><div className="catalog-product-title"><h3>{p.name}</h3><span className={publicNow ? 'catalog-public-pill published' : 'catalog-public-pill draft'}>{publicNow ? 'Publicado' : 'Borrador'}</span></div><p>{p.code || 'Sin código'} · {p.category || 'General'}{p.brand ? ` · ${p.brand}` : ''}</p><div className="catalog-meta-row"><strong>{money(p.price)}</strong><span className={availability === 'Agotado' ? 'availability-pill soldout' : availability === 'Últimas unidades' ? 'availability-pill low' : 'availability-pill available'}>{availability}</span><small>{productPriceStatus(p) === 'Validado' ? 'Precio validado' : 'Precio pendiente'}</small></div>{!readyForPublic(p) && <div className="catalog-warning">No se puede publicar hasta validar el precio.</div>}</div></div>
            <div className="catalog-card-actions"><button type="button" className={isEditing ? 'secondary-btn' : 'primary-btn'} onClick={()=>setEditingId(isEditing ? '' : p.id)}>{isEditing ? 'Cerrar edición' : 'Editar publicación'}</button>{publicNow && <button type="button" className="secondary-btn" onClick={()=>window.open(catalogProductUrl(p), '_blank', 'noopener,noreferrer')}>Ver ficha pública</button>}</div>
            {isEditing && <div className="catalog-admin-form catalog-v321-form">
              <div className="catalog-edit-heading"><div><span className="eyebrow">Configuración de publicación</span><strong>{p.name}</strong></div><span className="result-pill">{publicNow ? 'Visible' : 'No visible'}</span></div>
              <label className="check-row"><input type="checkbox" checked={Boolean(draft.public_visible)} onChange={e=>updateDraft(p,{public_visible:e.target.checked, catalog_status:e.target.checked && draft.catalog_status==='Borrador' ? 'Publicado' : draft.catalog_status})} /> Mostrar en el catálogo público</label>
              <label>Estado<select value={draft.catalog_status} onChange={e=>updateDraft(p,{catalog_status:e.target.value})} disabled={!draft.public_visible}><option>Borrador</option><option>Publicado</option><option>Oculto</option></select></label>
              <label className="check-row"><input type="checkbox" checked={Boolean(draft.catalog_featured)} onChange={e=>updateDraft(p,{catalog_featured:e.target.checked})} /> Mostrar como destacado</label>
              <label>Orden de aparición<input value={draft.catalog_position} inputMode="numeric" onChange={e=>updateDraft(p,{catalog_position:e.target.value})} /></label>
              <label className="catalog-description-field">Descripción para el cliente<textarea value={draft.catalog_description} onChange={e=>updateDraft(p,{catalog_description:e.target.value})} placeholder={p.description || 'Material, uso o detalle principal'} rows="3" /></label>
              <div className="catalog-form-actions"><button type="button" className="secondary-btn" onClick={()=>setEditingId('')}>Cancelar</button><button type="button" className="primary-btn" disabled={savingId===p.id} onClick={()=>saveProductCatalog(p)}>{savingId===p.id ? 'Guardando...' : 'Guardar cambios'}</button></div>
            </div>}
          </article>;
        })}
        {!rows.length && <section className="card compact-card"><p className="muted">No hay productos que coincidan con el filtro.</p></section>}
      </section>
      {rows.length > CATALOG_PAGE_SIZE && <div className="compact-pagination catalog-pagination"><span>Mostrando {(safeCatalogPage - 1) * CATALOG_PAGE_SIZE + 1}–{Math.min(safeCatalogPage * CATALOG_PAGE_SIZE, rows.length)} de {rows.length}</span><div><button type="button" className="secondary-btn" disabled={safeCatalogPage === 1} onClick={() => setCatalogPage(p => Math.max(1, p - 1))}>← Anterior</button><span className="pagination-page">Página {safeCatalogPage} de {catalogPageCount}</span><button type="button" className="secondary-btn" disabled={safeCatalogPage === catalogPageCount} onClick={() => setCatalogPage(p => Math.min(catalogPageCount, p + 1))}>Siguiente →</button></div></div>}
    </div>
  );
}

function WhatsAppCRM({ profile, store, products = [], onNavigate }) {
  const [channels, setChannels] = useState([]);
  const [conversations, setConversations] = useState([]);
  const [contacts, setContacts] = useState([]);
  const [messages, setMessages] = useState([]);
  const [orders, setOrders] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState('Conversaciones');
  const [filter, setFilter] = useState('Todas');
  const [selectedId, setSelectedId] = useState('');
  const [draft, setDraft] = useState('');
  const [sending, setSending] = useState(false);
  const [drafting, setDrafting] = useState(false);
  const [notice, setNotice] = useState('');
  const [channelForm, setChannelForm] = useState({ phone_number_id: '', display_phone_number: '', business_account_id: '', auto_reply_enabled: false });
  const [savingChannel, setSavingChannel] = useState(false);
  const storeId = profile?.store_id || DEFAULT_STORE_ID;
  const statuses = ['Nuevo','Consultando','Pendiente de pago','Pago por validar','En preparación','Entregado','Cancelado'];

  const selectedConversation = conversations.find(row => row.id === selectedId) || conversations[0] || null;
  const contactById = useMemo(() => contacts.reduce((acc, row) => { acc[row.id] = row; return acc; }, {}), [contacts]);
  const messagesByConversation = useMemo(() => messages.reduce((acc, row) => { (acc[row.conversation_id] ||= []).push(row); return acc; }, {}), [messages]);
  const selectedMessages = useMemo(() => (messagesByConversation[selectedConversation?.id] || []).sort((a, b) => new Date(a.created_at) - new Date(b.created_at)), [messagesByConversation, selectedConversation?.id]);
  const activeChannel = channels.find(row => row.active !== false) || channels[0] || null;
  const filteredConversations = useMemo(() => {
    if (filter === 'Todas') return conversations;
    return conversations.filter(row => row.status === filter);
  }, [conversations, filter]);
  const pendingCount = conversations.filter(row => !['Entregado','Cancelado'].includes(row.status || 'Nuevo')).length;
  const unreadCount = conversations.filter(row => (row.unread_count || 0) > 0).length;
  const pendingOrders = orders.filter(order => !['Entregado','Cancelado'].includes(order.status || 'Nuevo')).length;
  const webhookUrl = `${import.meta.env.VITE_SUPABASE_URL || 'https://TU-PROYECTO.supabase.co'}/functions/v1/clomar-whatsapp`;

  function humanTime(value) {
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    const diff = Math.max(0, Date.now() - date.getTime());
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return 'Ahora';
    if (mins < 60) return `${mins} min`;
    const hours = Math.floor(mins / 60);
    if (hours < 24) return `${hours} h`;
    return date.toLocaleDateString('es-PE', { day: '2-digit', month: 'short' });
  }

  async function loadCRM() {
    if (!hasSupabaseConfig) return;
    setLoading(true); setError('');
    const [channelsRes, contactsRes, conversationsRes, messagesRes, ordersRes] = await Promise.all([
      supabase.from('whatsapp_channels').select('*').eq('store_id', storeId).order('created_at', { ascending: false }),
      supabase.from('whatsapp_contacts').select('*').eq('store_id', storeId).order('last_seen_at', { ascending: false }).limit(600),
      supabase.from('whatsapp_conversations').select('*').eq('store_id', storeId).order('last_message_at', { ascending: false }).limit(600),
      supabase.from('whatsapp_messages').select('*').eq('store_id', storeId).order('created_at', { ascending: false }).limit(3000),
      supabase.from('catalog_orders').select('*').eq('store_id', storeId).order('created_at', { ascending: false }).limit(300),
    ]);
    const results = [channelsRes, contactsRes, conversationsRes, messagesRes, ordersRes];
    const firstError = results.find(result => result.error)?.error;
    if (firstError) {
      setError(firstError.message || 'No se pudo cargar el CRM. Ejecute primero el SQL de V03.4.');
      setLoading(false);
      return;
    }
    setChannels(channelsRes.data || []);
    setContacts(contactsRes.data || []);
    setConversations(conversationsRes.data || []);
    setMessages(messagesRes.data || []);
    setOrders(ordersRes.data || []);
    setSelectedId(current => current || conversationsRes.data?.[0]?.id || '');
    const firstChannel = channelsRes.data?.[0];
    if (firstChannel) setChannelForm({
      phone_number_id: firstChannel.phone_number_id || '',
      display_phone_number: firstChannel.display_phone_number || '',
      business_account_id: firstChannel.business_account_id || '',
      auto_reply_enabled: Boolean(firstChannel.auto_reply_enabled),
    });
    setLoading(false);
  }

  useEffect(() => { loadCRM(); }, [storeId]);

  async function updateConversationStatus(conversation, status) {
    if (!conversation) return;
    const { error: updateError } = await supabase.from('whatsapp_conversations').update({ status, updated_at: new Date().toISOString() }).eq('id', conversation.id).eq('store_id', storeId);
    if (updateError) { setNotice(`No se pudo actualizar: ${updateError.message}`); return; }
    setConversations(prev => prev.map(row => row.id === conversation.id ? { ...row, status } : row));
    setNotice('Estado comercial actualizado.');
  }

  async function saveChannel(e) {
    e?.preventDefault?.();
    const phoneNumberId = String(channelForm.phone_number_id || '').trim();
    if (!phoneNumberId) { setNotice('Ingrese el Phone Number ID de WhatsApp Cloud API.'); return; }
    setSavingChannel(true); setNotice('');
    const payload = {
      store_id: storeId,
      phone_number_id: phoneNumberId,
      display_phone_number: String(channelForm.display_phone_number || '').trim(),
      business_account_id: String(channelForm.business_account_id || '').trim(),
      auto_reply_enabled: Boolean(channelForm.auto_reply_enabled),
      active: true,
      updated_at: new Date().toISOString(),
    };
    const { error: saveError } = activeChannel
      ? await supabase.from('whatsapp_channels').update(payload).eq('id', activeChannel.id).eq('store_id', storeId)
      : await supabase.from('whatsapp_channels').insert(payload);
    setSavingChannel(false);
    if (saveError) { setNotice(`No se pudo guardar el canal: ${saveError.message}`); return; }
    setNotice('Canal guardado. Configure el webhook en Meta y haga una prueba con su número.');
    await loadCRM();
  }

  async function invokeHub(body) {
    const { data, error: functionError } = await supabase.functions.invoke('clomar-whatsapp', { body });
    if (functionError) throw functionError;
    if (data?.error) throw new Error(data.error);
    return data;
  }

  async function createAIDraft() {
    if (!selectedConversation) { setNotice('Seleccione una conversación.'); return; }
    setDrafting(true); setNotice('');
    try {
      const data = await invokeHub({ action: 'ai_draft', store_id: storeId, conversation_id: selectedConversation.id, tone: 'Cercano' });
      setDraft(data?.draft || '');
      setNotice(data?.mode === 'gemini' ? 'Borrador generado por Gemini con datos verificados.' : 'Borrador preparado con información verificada del ERP.');
    } catch (err) {
      setNotice(err?.message || 'No se pudo generar el borrador.');
    } finally { setDrafting(false); }
  }

  async function sendMessage() {
    if (!selectedConversation || !String(draft || '').trim()) { setNotice('Escriba o genere un mensaje antes de enviarlo.'); return; }
    setSending(true); setNotice('');
    try {
      const data = await invokeHub({ action: 'send', store_id: storeId, conversation_id: selectedConversation.id, text: draft.trim() });
      const message = data?.message;
      if (message) setMessages(prev => [...prev, message]);
      setConversations(prev => prev.map(row => row.id === selectedConversation.id ? { ...row, last_message_at: new Date().toISOString(), last_message_preview: draft.trim(), unread_count: 0 } : row));
      setDraft('');
      setNotice('Mensaje enviado por WhatsApp y registrado en el CRM.');
    } catch (err) {
      setNotice(err?.message || 'No se pudo enviar el mensaje. Revise que el canal y la ventana de atención estén activos.');
    } finally { setSending(false); }
  }

  const selectedContact = selectedConversation ? contactById[selectedConversation.contact_id] : null;
  const canSend = Boolean(activeChannel?.phone_number_id && selectedConversation && selectedContact);

  return (
    <div className="page wa-crm-page">
      <section className="wa-crm-hero">
        <div><span className="eyebrow">WhatsApp Cloud API + CRM operativo</span><h1>Centro comercial de conversaciones</h1><p>Reciba consultas, prepare respuestas verificadas, cree pedidos y derive los casos sensibles a una persona.</p></div>
        <div className={`wa-channel-state ${activeChannel?.phone_number_id ? 'ready' : ''}`}><span className="status-dot"/>{activeChannel?.phone_number_id ? 'Canal configurado' : 'Canal pendiente'}</div>
      </section>
      <section className="wa-crm-safe"><strong>Control humano:</strong> la IA puede sugerir y preparar respuestas. Los descuentos, pagos, créditos, cambios y devoluciones deben ser confirmados por una persona.</section>
      {error && <section className="data-error"><strong>CRM no disponible todavía.</strong> Ejecute <code>supabase/INSTALAR_V03_4_WHATSAPP_CRM.sql</code>. Detalle: {error}</section>}
      <section className="wa-crm-kpis">
        <Kpi label="Conversaciones abiertas" value={pendingCount} helper="requieren atención" />
        <Kpi label="Mensajes nuevos" value={unreadCount} helper="sin revisar" />
        <Kpi label="Pedidos por seguir" value={pendingOrders} helper="solicitudes web y WhatsApp" />
        <Kpi label="Canal" value={activeChannel?.display_phone_number || 'Pendiente'} helper={activeChannel?.auto_reply_enabled ? 'IA con respuesta controlada' : 'Respuesta automática desactivada'} />
      </section>
      <section className="wa-crm-tabs">
        {['Conversaciones','Pedidos','Configuración'].map(tab => <button type="button" key={tab} className={activeTab === tab ? 'active' : ''} onClick={() => setActiveTab(tab)}>{tab}</button>)}
      </section>
      {activeTab === 'Configuración' && <section className="wa-setup-grid">
        <form className="card wa-channel-form" onSubmit={saveChannel}>
          <span className="eyebrow">Canal oficial</span><h2>Conectar WhatsApp Cloud API</h2><p>Los tokens privados se guardan solo como secretos de Supabase. Aquí registre datos no sensibles del número oficial.</p>
          <label>Phone Number ID<input value={channelForm.phone_number_id} onChange={e => setChannelForm({ ...channelForm, phone_number_id: e.target.value })} placeholder="Ej.: 123456789012345" required /></label>
          <label>Número visible<input value={channelForm.display_phone_number} onChange={e => setChannelForm({ ...channelForm, display_phone_number: e.target.value })} placeholder="+51 931 709 871" /></label>
          <label>WhatsApp Business Account ID <small>Opcional, útil para control interno</small><input value={channelForm.business_account_id} onChange={e => setChannelForm({ ...channelForm, business_account_id: e.target.value })} placeholder="WABA ID" /></label>
          <label className="wa-check"><input type="checkbox" checked={Boolean(channelForm.auto_reply_enabled)} onChange={e => setChannelForm({ ...channelForm, auto_reply_enabled: e.target.checked })} /><span><strong>Activar respuestas automáticas controladas</strong><small>Solo responde mensajes entrantes con datos verificables. Manténgalo desactivado durante las pruebas.</small></span></label>
          <button className="primary-btn" disabled={savingChannel}>{savingChannel ? 'Guardando...' : 'Guardar canal'}</button>
        </form>
        <section className="card wa-webhook-guide">
          <span className="eyebrow">Paso Meta</span><h2>Webhook de Clomar Store</h2>
          <p>En Meta Developers agregue esta URL como Callback URL:</p>
          <code>{webhookUrl}</code>
          <p>Use el mismo valor que guardará como secreto <strong>WHATSAPP_VERIFY_TOKEN</strong> para verificar el webhook.</p>
          <ol><li>Suscriba el campo <strong>messages</strong>.</li><li>Pruebe desde el número permitido por Meta.</li><li>Envíe “Hola” al WhatsApp oficial.</li><li>Revise que aparezca una conversación aquí.</li></ol>
          <button type="button" className="secondary-btn" onClick={() => navigator.clipboard?.writeText(webhookUrl).then(() => setNotice('URL del webhook copiada.'))}>Copiar URL del webhook</button>
        </section>
      </section>}
      {activeTab === 'Pedidos' && <section className="wa-orders-board">
        <div className="wa-board-head"><div><span className="eyebrow">Pedidos comerciales</span><h2>Solicitudes pendientes</h2><p>Confirme disponibilidad y pago antes de registrar la venta final en el POS.</p></div><button type="button" className="secondary-btn" onClick={loadCRM}>{loading ? 'Actualizando...' : 'Actualizar'}</button></div>
        <div className="wa-order-grid">{orders.map(order => <article className="wa-order-card" key={order.id}><div><span>{order.order_code || 'Pedido'}</span><h3>{order.customer_name || 'Cliente'}</h3><small>{order.customer_phone || 'Sin número'} · {fmtDate(order.created_at)}</small></div><div><b>{money(order.total_amount)}</b><select value={order.status || 'Nuevo'} onChange={async e => { const { error: updateError } = await supabase.from('catalog_orders').update({ status: e.target.value, updated_at: new Date().toISOString() }).eq('id', order.id).eq('store_id', storeId); if (updateError) setNotice(updateError.message); else loadCRM(); }}>{statuses.map(status => <option key={status}>{status}</option>)}</select></div></article>)}{!orders.length && <div className="wa-empty">No hay pedidos registrados todavía.</div>}</div>
      </section>}
      {activeTab === 'Conversaciones' && <section className="wa-crm-workspace">
        <aside className="wa-conversation-list"><div className="wa-list-head"><div><span className="eyebrow">Bandeja</span><h2>Conversaciones</h2></div><button type="button" onClick={loadCRM}>{loading ? '…' : '↻'}</button></div><div className="wa-filter-row"><button type="button" className={filter==='Todas'?'active':''} onClick={() => setFilter('Todas')}>Todas</button><button type="button" className={filter==='Nuevo'?'active':''} onClick={() => setFilter('Nuevo')}>Nuevas</button><button type="button" className={filter==='Pendiente de pago'?'active':''} onClick={() => setFilter('Pendiente de pago')}>Pagos</button></div><div className="wa-conversation-scroll">{filteredConversations.map(conversation => { const contact = contactById[conversation.contact_id] || {}; const active = selectedConversation?.id === conversation.id; return <button key={conversation.id} type="button" className={`wa-conversation-row ${active ? 'active' : ''}`} onClick={() => { setSelectedId(conversation.id); setDraft(''); }}><span className="wa-avatar">{String(contact.display_name || contact.wa_id || '?').slice(0,1).toUpperCase()}</span><span><strong>{contact.display_name || contact.wa_id || 'Cliente WhatsApp'}</strong><small>{conversation.last_message_preview || 'Sin mensajes todavía'}</small></span><em>{humanTime(conversation.last_message_at)}{conversation.unread_count ? <b>{conversation.unread_count}</b> : null}</em></button>; })}{!filteredConversations.length && <div className="wa-empty">Aún no hay conversaciones. Configure Meta y escriba al número oficial para probar.</div>}</div></aside>
        <main className="wa-chat-panel">{selectedConversation ? <><header className="wa-chat-head"><div><span className="wa-avatar">{String(selectedContact?.display_name || selectedContact?.wa_id || '?').slice(0,1).toUpperCase()}</span><div><h2>{selectedContact?.display_name || selectedContact?.wa_id || 'Cliente WhatsApp'}</h2><small>{selectedContact?.phone_number || selectedContact?.wa_id || 'WhatsApp'} · {selectedConversation.status || 'Nuevo'}</small></div></div><select value={selectedConversation.status || 'Nuevo'} onChange={e => updateConversationStatus(selectedConversation, e.target.value)}>{statuses.map(status => <option key={status}>{status}</option>)}</select></header><div className="wa-message-thread">{selectedMessages.map(message => <article key={message.id} className={`wa-message ${message.direction === 'Saliente' ? 'out' : 'in'}`}><p>{message.body || '[Mensaje no textual]'}</p><small>{fmtDate(message.created_at)} · {message.status || (message.direction === 'Saliente' ? 'Enviado' : 'Recibido')}</small></article>)}{!selectedMessages.length && <div className="wa-empty">No hay mensajes guardados todavía.</div>}</div><div className="wa-compose"><div className="wa-compose-actions"><button type="button" className="secondary-btn" onClick={createAIDraft} disabled={drafting || !canSend}>{drafting ? 'Generando...' : 'Sugerir respuesta IA'}</button><button type="button" className="ghost-btn" onClick={() => setDraft('Hola, gracias por comunicarte con Clomar Store. ¿En qué producto o talla te puedo ayudar?')}>Saludo</button></div><textarea value={draft} onChange={e => setDraft(e.target.value)} placeholder="Escriba una respuesta o pida un borrador a la IA..." rows="4" /><div className="wa-send-row"><small>{activeChannel?.auto_reply_enabled ? 'Bot controlado activo: revise siempre antes de confirmar ventas o pagos.' : 'Respuesta automática desactivada: usted decide cada envío.'}</small><button className="primary-btn" disabled={!canSend || sending || !draft.trim()} onClick={sendMessage}>{sending ? 'Enviando...' : 'Enviar por WhatsApp'}</button></div></div></> : <div className="wa-chat-empty"><h2>Seleccione una conversación</h2><p>Las consultas de clientes aparecerán aquí después de conectar el webhook de Meta.</p></div>}</main>
        <aside className="wa-customer-panel">{selectedConversation ? <><span className="eyebrow">Ficha comercial</span><h2>{selectedContact?.display_name || 'Cliente WhatsApp'}</h2><div className="wa-customer-data"><span>WhatsApp</span><strong>{selectedContact?.phone_number || selectedContact?.wa_id || '—'}</strong><span>Estado</span><strong>{selectedConversation.status || 'Nuevo'}</strong><span>Último contacto</span><strong>{fmtDate(selectedConversation.last_message_at) || '—'}</strong></div><div className="wa-ai-rules"><strong>La IA puede</strong><p>Buscar productos, explicar precio, disponibilidad y preparar borradores.</p><strong>Requiere persona</strong><p>Descuentos, crédito, pago, entrega, reclamos, cambios y devoluciones.</p></div><button type="button" className="secondary-btn" onClick={() => onNavigate?.('clientes')}>Abrir clientes</button><button type="button" className="secondary-btn" onClick={() => onNavigate?.('ventas')}>Registrar venta final</button></> : <><span className="eyebrow">CRM operativo</span><h2>Atención guiada</h2><p>Cuando un cliente escriba, podrá convertir la consulta en pedido y luego registrar la venta desde el POS.</p></>}</aside>
      </section>}
      {notice && <div className="wa-toast" role="status">{notice}<button type="button" onClick={() => setNotice('')}>×</button></div>}
    </div>
  );
}

function isPublicCatalogLocation() {
  return String(window.location.hash || '').toLowerCase().startsWith('#/catalogo');
}
function catalogProductIdFromLocation() {
  const query = String(window.location.hash || '').split('?')[1] || '';
  return new URLSearchParams(query).get('product') || '';
}
function catalogProductMessage(product, store) {
  const bits = [
    'Hola, deseo consultar este producto:',
    `Producto: ${product?.name || ''}`,
    product?.size ? `Talla: ${product.size}` : '',
    product?.color ? `Color: ${product.color}` : '',
    `Precio: ${money(product?.price || 0)}`,
    product?.code ? `Código: ${product.code}` : '',
  ].filter(Boolean);
  return bits.join('\n');
}

function PublicCatalogApp() {
  const [store, setStore] = useState({ id: DEFAULT_STORE_ID, name: 'Clomar Store', whatsapp_number: CATALOG_WHATSAPP_FALLBACK });
  const [products, setProducts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [category, setCategory] = useState('Todas');
  const [selectedId, setSelectedId] = useState(() => catalogProductIdFromLocation());
  const [cart, setCart] = useState([]);
  const [cartOpen, setCartOpen] = useState(false);
  const [form, setForm] = useState({ name: '', phone: '', note: '' });
  const [submitting, setSubmitting] = useState(false);
  const [notice, setNotice] = useState('');

  async function loadCatalog() {
    if (!hasSupabaseConfig) { setError('El catálogo requiere las variables de Supabase configuradas.'); setLoading(false); return; }
    setLoading(true); setError('');
    const [storeRes, productsRes] = await Promise.all([
      supabase.rpc('clomar_public_catalog_store_v31', { p_store_id: DEFAULT_STORE_ID }),
      supabase.rpc('clomar_public_catalog_products_v31', { p_store_id: DEFAULT_STORE_ID, p_query: null, p_product_id: null }),
    ]);
    if (storeRes.error || productsRes.error) setError([storeRes.error?.message, productsRes.error?.message].filter(Boolean).join(' · ') || 'No se pudo cargar el catálogo.');
    else {
      const publicStore = Array.isArray(storeRes.data) ? storeRes.data[0] : storeRes.data;
      setStore(publicStore || { id: DEFAULT_STORE_ID, name: 'Clomar Store', whatsapp_number: CATALOG_WHATSAPP_FALLBACK });
      setProducts(productsRes.data || []);
    }
    setLoading(false);
  }
  useEffect(() => { loadCatalog(); }, []);
  useEffect(() => { const onHash = () => setSelectedId(catalogProductIdFromLocation()); window.addEventListener('hashchange', onHash); return () => window.removeEventListener('hashchange', onHash); }, []);

  const categories = useMemo(() => ['Todas', ...Array.from(new Set(products.map(p => p.category || 'General')))], [products]);
  const categoryCounts = useMemo(() => products.reduce((acc,p) => { const key = p.category || 'General'; acc[key] = (acc[key] || 0) + 1; return acc; }, {}), [products]);
  const visibleProducts = useMemo(() => products.filter(p => {
    const q = normalizeText(query);
    const textMatch = !q || [p.name,p.category,p.subcategory,p.brand,p.color,p.size,p.code].some(v => normalizeText(v).includes(q));
    const catMatch = category === 'Todas' || (p.category || 'General') === category;
    return textMatch && catMatch;
  }), [products, query, category]);
  const selected = products.find(p => p.id === selectedId) || null;
  const cartTotal = cart.reduce((sum,line)=>sum+asNum(line.price)*asNum(line.qty),0);
  const cartQty = cart.reduce((sum,line)=>sum+asNum(line.qty),0);
  const inCart = (id) => cart.find(line => line.id === id);
  const canOrder = (p) => catalogAvailabilityLabel(p.availability) !== 'Agotado';
  const featuredProducts = useMemo(() => {
    const featured = products.filter(p => p.catalog_featured);
    return (featured.length ? featured : products).slice(0, 3);
  }, [products]);

  function selectProduct(p) { setSelectedId(p.id); window.history.replaceState(null, '', `#/catalogo?product=${encodeURIComponent(p.id)}`); }
  function closeProduct() { setSelectedId(''); window.history.replaceState(null, '', '#/catalogo'); }
  function addToCart(p) {
    if (!canOrder(p)) return alert('Este producto está agotado. Puede consultar opciones similares por WhatsApp.');
    setCart(prev => { const exists = prev.find(line => line.id === p.id); if (exists) return prev.map(line => line.id === p.id ? { ...line, qty: Math.min(10, asNum(line.qty) + 1) } : line); return [...prev, { id:p.id, name:p.name, code:p.code, color:p.color, size:p.size, price:asNum(p.price), image_url:p.image_url, qty:1 }]; });
    setNotice(`${p.name} se agregó al pedido.`); setTimeout(() => setNotice(''), 1800);
  }
  function changeCartQty(id, delta) { setCart(prev => prev.map(line => line.id === id ? { ...line, qty: Math.max(1, Math.min(10, asNum(line.qty)+delta)) } : line)); }
  function removeCartLine(id) { setCart(prev => prev.filter(line => line.id !== id)); }
  function consultProduct(p) { window.open(publicWhatsAppLink(store?.whatsapp_number || store?.phone, catalogProductMessage(p, store)), '_blank', 'noopener,noreferrer'); }
  function openWhatsAppCatalog() { window.open(publicWhatsAppLink(store?.whatsapp_number || store?.phone, 'Hola, deseo consultar el catálogo de productos.'), '_blank', 'noopener,noreferrer'); }

  async function submitOrder(e) {
    e.preventDefault();
    if (!cart.length) return;
    if (String(form.name || '').trim().length < 2) return alert('Indique su nombre para registrar el pedido.');
    if (String(form.phone || '').replace(/\D/g,'').length < 7) return alert('Indique un teléfono o WhatsApp válido.');
    const popup = window.open('', '_blank'); setSubmitting(true);
    const { data, error: orderError } = await supabase.rpc('clomar_create_catalog_order_v31', { p_store_id: store?.id || DEFAULT_STORE_ID, p_customer_name: form.name.trim(), p_customer_phone: form.phone.trim(), p_customer_note: form.note.trim() || null, p_items: cart.map(line => ({ product_id: line.id, qty: asNum(line.qty) })) });
    setSubmitting(false);
    if (orderError) { if (popup) popup.close(); return alert(`No se pudo registrar el pedido: ${orderError.message}. Verifique disponibilidad y vuelva a intentarlo.`); }
    const order = Array.isArray(data) ? data[0] : data;
    const lines = cart.map(line => `${line.qty} × ${line.name}${line.color ? ` · ${line.color}` : ''}${line.size ? ` · ${line.size}` : ''} — ${money(asNum(line.price)*asNum(line.qty))}`);
    const message = ['Hola, deseo confirmar este pedido de catálogo:', `Pedido: ${order?.order_code || 'Solicitud web'}`, '', ...lines, '', `Total referencial: ${money(order?.total_amount || cartTotal)}`, 'Nombre: ' + form.name.trim(), 'Teléfono: ' + form.phone.trim(), form.note.trim() ? `Nota: ${form.note.trim()}` : ''].filter(Boolean).join('\n');
    const url = publicWhatsAppLink(store?.whatsapp_number || store?.phone, message);
    if (popup) popup.location.href = url; else window.location.href = url;
    setCart([]); setCartOpen(false); setForm({ name:'', phone:'', note:'' }); setNotice(`Pedido ${order?.order_code || ''} registrado. Continúe en WhatsApp para confirmarlo.`);
  }

  if (loading) return <main className="catalog-public-shell"><div className="catalog-public-loader">Cargando catálogo...</div></main>;
  if (error) return <main className="catalog-public-shell"><section className="catalog-public-error"><h1>Catálogo temporalmente no disponible</h1><p>{error}</p><button className="primary-btn" onClick={loadCatalog}>Reintentar</button></section></main>;

  return (
    <main className="catalog-public-shell catalog-v326-shell">
      <header className="catalog-public-header"><a className="catalog-public-brand" href="#/catalogo" onClick={closeProduct}><img src={publicAssetUrl(store?.logo_url || APP_ICON)} alt={store?.name || 'Clomar Store'}/><span><strong>{store?.name || 'Clomar Store'}</strong><small>Catálogo oficial</small></span></a><div className="catalog-public-header-actions"><button type="button" className="catalog-whatsapp-btn" onClick={openWhatsAppCatalog}>WhatsApp</button><button type="button" className="catalog-cart-btn" onClick={()=>setCartOpen(true)}>Pedido <span>{cartQty}</span></button></div></header>
      <section className="catalog-public-hero catalog-v326-hero">
        <div className="catalog-hero-copy"><span className="catalog-kicker">Compra simple · atención rápida</span><h1>Encuentra productos para tu día a día.</h1><p>Revisa precios, elige tus opciones y confirma tu pedido directamente por WhatsApp.</p><div className="catalog-hero-actions"><button type="button" className="primary-btn" onClick={()=>document.getElementById('catalog-products')?.scrollIntoView({behavior:'smooth'})}>Ver productos</button><button type="button" className="secondary-btn" onClick={openWhatsAppCatalog}>Consultar por WhatsApp</button></div><div className="catalog-hero-points"><span>✓ Precios visibles</span><span>✓ Pedido sin complicaciones</span><span>✓ Atención personalizada</span></div></div>
        <div className="catalog-hero-showcase"><div className="catalog-hero-showcase-top"><span>CATÁLOGO ACTUALIZADO</span><b>{products.length} producto(s) disponibles</b></div><div className="catalog-hero-mini-products">{featuredProducts.map(p=><button type="button" key={p.id} onClick={()=>selectProduct(p)}><img src={p.image_url || APP_ICON} alt={p.name}/><span><strong>{p.name}</strong><small>{money(p.price)}</small></span></button>)}{!featuredProducts.length && <div className="catalog-hero-empty">Pronto habrá productos disponibles.</div>}</div></div>
      </section>
      <section className="catalog-category-rail-wrap"><div className="catalog-category-rail">{categories.map(c=><button key={c} type="button" className={category===c?'active':''} onClick={()=>setCategory(c)}><span>{c}</span><b>{c === 'Todas' ? products.length : categoryCounts[c] || 0}</b></button>)}</div></section>
      <section className="catalog-public-trust catalog-v326-trust"><div><span>01</span><b>Compra con confianza</b><small>Verá precio y disponibilidad antes de consultar.</small></div><div><span>02</span><b>Pedido por WhatsApp</b><small>Envíe su selección y reciba confirmación personal.</small></div><div><span>03</span><b>Atención directa</b><small>Un asesor confirma talla, color y entrega.</small></div></section>
      {notice && <div className="catalog-toast">{notice}</div>}
      <section className="catalog-products-section" id="catalog-products"><div className="catalog-products-head"><div><span className="catalog-kicker">Explorar</span><h2>{category === 'Todas' ? 'Productos disponibles' : category}</h2><p>{visibleProducts.length} resultado(s) · use la búsqueda para filtrar por marca, color, talla o nombre.</p></div><div className="catalog-public-search"><Search size={20}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar producto, talla, color o marca..." /></div></div><section className="catalog-public-grid catalog-v326-grid">{visibleProducts.map(p => <article className="catalog-public-card catalog-v326-card" key={p.id}><button type="button" className="catalog-image-btn" onClick={()=>selectProduct(p)}><img src={p.image_url || APP_ICON} alt={p.name}/>{p.catalog_featured && <span className="catalog-featured-badge">Recomendado</span>}<span className={p.availability==='Agotado'?'catalog-availability soldout':p.availability==='Últimas unidades'?'catalog-availability low':'catalog-availability'}>{catalogAvailabilityLabel(p.availability)}</span></button><div className="catalog-card-body"><p>{p.category || 'General'}</p><h2>{p.name}</h2><small>{[p.brand,p.color,p.size].filter(Boolean).join(' · ') || 'Producto Clomar Store'}</small><strong>{money(p.price)}</strong><div className="catalog-card-actions"><button type="button" className="secondary-btn" onClick={()=>consultProduct(p)}>Consultar</button><button type="button" className="primary-btn" disabled={!canOrder(p)} onClick={()=>addToCart(p)}>{canOrder(p) ? (inCart(p.id) ? 'Agregar otra' : 'Agregar al pedido') : 'Agotado'}</button></div></div></article>)}{visibleProducts.length > 0 && visibleProducts.length < 4 && <article className="catalog-assistance-card catalog-v326-assistance"><span>¿Necesita otra opción?</span><h2>Le ayudamos por WhatsApp</h2><p>Consulte por modelos, tallas, colores o productos similares disponibles.</p><button type="button" onClick={openWhatsAppCatalog}>Hablar con un asesor</button></article>}{!visibleProducts.length && <section className="catalog-empty"><h2>No encontramos productos</h2><p>Pruebe otra búsqueda o consulte por WhatsApp.</p><button type="button" onClick={openWhatsAppCatalog}>Consultar por WhatsApp</button></section>}</section></section>
      {selected && <div className="catalog-detail-backdrop" onMouseDown={closeProduct}><article className="catalog-detail-modal" onMouseDown={e=>e.stopPropagation()}><button type="button" className="catalog-detail-close" onClick={closeProduct}>×</button><img src={selected.image_url || APP_ICON} alt={selected.name}/><div><span className={selected.availability==='Agotado'?'catalog-availability soldout':selected.availability==='Últimas unidades'?'catalog-availability low':'catalog-availability'}>{catalogAvailabilityLabel(selected.availability)}</span><p>{selected.category || 'General'}</p><h2>{selected.name}</h2><strong>{money(selected.price)}</strong><div className="catalog-detail-specs">{selected.brand && <span>Marca: {selected.brand}</span>}{selected.color && <span>Color: {selected.color}</span>}{selected.size && <span>Talla: {selected.size}</span>}{selected.code && <span>Código: {selected.code}</span>}</div><p className="catalog-detail-description">{selected.catalog_description || selected.description || 'Consulta disponibilidad y detalles por WhatsApp.'}</p><div className="catalog-detail-actions"><button className="secondary-btn" type="button" onClick={()=>consultProduct(selected)}>Consultar por WhatsApp</button><button className="primary-btn" type="button" disabled={!canOrder(selected)} onClick={()=>{addToCart(selected); setCartOpen(true);}}>{canOrder(selected)?'Agregar al pedido':'Agotado'}</button></div></div></article></div>}
      {cartOpen && <div className="catalog-cart-backdrop" onMouseDown={()=>setCartOpen(false)}><aside className="catalog-cart-drawer" onMouseDown={e=>e.stopPropagation()}><div className="catalog-cart-head"><div><span>Pedido por WhatsApp</span><h2>Tu selección</h2></div><button type="button" onClick={()=>setCartOpen(false)}>×</button></div>{cart.length ? <><div className="catalog-cart-lines">{cart.map(line=><article key={line.id}><img src={line.image_url || APP_ICON} alt={line.name}/><div><strong>{line.name}</strong><small>{[line.color,line.size,line.code].filter(Boolean).join(' · ')}</small><b>{money(line.price)}</b></div><div className="cart-qty-control"><button type="button" onClick={()=>changeCartQty(line.id,-1)}>−</button><span>{line.qty}</span><button type="button" onClick={()=>changeCartQty(line.id,1)}>+</button><button type="button" className="cart-remove" onClick={()=>removeCartLine(line.id)}>Quitar</button></div></article>)}</div><div className="catalog-cart-total"><span>Total referencial</span><strong>{money(cartTotal)}</strong></div><form className="catalog-order-form" onSubmit={submitOrder}><label>Nombre<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Nombre y apellido" required /></label><label>Tu WhatsApp<input value={form.phone} onChange={e=>setForm({...form, phone:e.target.value})} inputMode="tel" placeholder="999 999 999" required /></label><label>Nota opcional<textarea value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Color alternativo, entrega, consulta..." rows="3" /></label><p>Al enviar, registramos la solicitud y abrimos WhatsApp para confirmar disponibilidad y pago.</p><button type="submit" className="primary-btn" disabled={submitting}>{submitting ? 'Registrando...' : 'Enviar pedido por WhatsApp'}</button></form></> : <div className="catalog-cart-empty"><h3>Tu pedido está vacío</h3><p>Agregue productos del catálogo para enviarlos por WhatsApp.</p></div>}</aside></div>}
      <footer className="catalog-public-footer"><strong>{store?.name || 'Clomar Store'}</strong><span>Precios sujetos a confirmación de disponibilidad al momento de atender el pedido.</span></footer>
    </main>
  );
}

function MobileQuickAssistant({ open, onClose, products = [], setCurrent }) {
  const [query, setQuery] = useState('');
  const [mode, setMode] = useState('buscar');
  useEffect(() => { if (!open) { setQuery(''); setMode('buscar'); } }, [open]);
  if (!open) return null;
  const rows = products.filter(p => {
    if (p.active === false) return false;
    const match = !query.trim() || `${p.name || ''} ${p.code || ''} ${p.category || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(query.trim().toLowerCase());
    const modeMatch = mode !== 'low' || asNum(p.stock) <= asNum(p.stock_min);
    return match && modeMatch;
  }).slice(0, 6);
  const chooseProduct = (p) => {
    try { sessionStorage.setItem('clomar_pos_assistant_search', p.code || p.name || ''); } catch (_) {}
    setCurrent('ventas'); onClose();
  };
  return (
    <div className="mobile-ai-backdrop" role="dialog" aria-modal="true" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <section className="mobile-ai-sheet" onMouseDown={e => e.stopPropagation()}>
        <div className="mobile-ai-handle" />
        <header><div><span>ASISTENTE EXPRESS</span><h3>¿Qué necesita revisar?</h3><p>Consulta productos y stock con información real del ERP.</p></div><button type="button" aria-label="Cerrar asistente" onClick={onClose}>×</button></header>
        <div className="mobile-ai-quick"><button type="button" className={mode === 'buscar' ? 'active' : ''} onClick={() => setMode('buscar')}>Buscar producto</button><button type="button" className={mode === 'low' ? 'active' : ''} onClick={() => setMode('low')}>Stock bajo</button><button type="button" onClick={() => { setCurrent('pedidos'); onClose(); }}>Pedidos web</button></div>
        <div className="mobile-ai-search"><Search size={18}/><input autoFocus value={query} onChange={e => setQuery(e.target.value)} placeholder={mode === 'low' ? 'Buscar entre productos por reponer...' : 'Nombre, código, talla o color'} /></div>
        <div className="mobile-ai-results">{rows.length ? rows.map(p => { const low = asNum(p.stock) <= asNum(p.stock_min); return <button type="button" className="mobile-ai-result" key={p.id} onClick={() => chooseProduct(p)}><img src={productImageSrc(p)} alt={p.name}/><span><strong>{p.name}</strong><small>{p.code || 'Sin código'} · {p.color || p.category || 'General'}</small></span><b className={low ? 'low' : ''}>{low ? `Reponer ${asNum(p.stock)}` : `Stock ${asNum(p.stock)}`}</b></button>; }) : <div className="mobile-ai-empty">No encontramos productos con ese criterio.</div>}</div>
        <footer><button type="button" className="secondary-btn" onClick={() => { setCurrent('ia'); onClose(); }}>Abrir asistente completo</button></footer>
      </section>
    </div>
  );
}

function MobileBottomNav({ current, setCurrent, role, menuOpen = false, onAssistant }) {
  const labelMap = {
    panel: ['📊', 'Inicio'], ventas: ['🧾', 'Vender'], reportes: ['📈', 'Reportes'], caja: ['💰', 'Caja'], comprobantes: ['🧾', 'Tickets'], productos: ['📦', 'Productos'], catalogo: ['🛍️', 'Catálogo'], pedidos: ['📬', 'Pedidos'], inventario: ['📘', 'Stock'], ingreso: ['📥', 'Compras'], creditos: ['💳', 'Créditos']
  };
  const preferred = role === 'almacen'
    ? ['productos', 'inventario', 'ingreso', 'catalogo']
    : role === 'cajero'
      ? ['ventas', 'comprobantes', 'caja', 'creditos']
      : role === 'lectura'
        ? ['panel', 'reportes', 'inventario']
        : ['panel', 'ventas', 'caja', 'pedidos'];
  const items = preferred.filter(key => canAccess(role, key) && labelMap[key]).slice(0, 4);
  if (!items.length || menuOpen) return null;
  return (
    <nav className="mobile-bottom-nav mobile-bottom-nav-v326" aria-label="Navegación móvil principal">
      {items.map(key => { const [icon, label] = labelMap[key]; return <button key={key} type="button" className={current === key ? 'active' : ''} onClick={() => setCurrent(key)}><span>{icon}</span><small>{label}</small></button>; })}
      <button type="button" className="mobile-assistant-nav" onClick={onAssistant}><span>✦</span><small>Asistente</small></button>
    </nav>
  );
}

function AppShell({ session }) {
  const [current, setCurrent] = useState(() => {
    try { return localStorage.getItem('clomar_last_module') || 'ventas'; } catch (err) { return 'ventas'; }
  });
  const [open, setOpen] = useState(false);
  const [mobileAssistantOpen, setMobileAssistantOpen] = useState(false);
  const { profile, store, loading: profileLoading, reload: reloadProfile } = useUserProfile(session);
  const { products, loading, reload } = useProducts(profile);
  const { customers, reload: reloadCustomers } = useCustomers(profile);
  const { categories, subcategories, reload: reloadCategories } = useCategories(profile);
  const cashSession = useCashSession(profile);

  useEffect(() => { if (profile && !canAccess(profile.role, current)) setCurrent(firstAllowedModule(profile.role)); }, [profile?.role, current]);
  useEffect(() => { try { localStorage.setItem('clomar_last_module', current); } catch (err) {} }, [current]);

  if (profileLoading) return <div className="loader full">Cargando perfil y permisos...</div>;
  if (profile?.status === 'Inactivo') return <InactiveUser profile={profile} />;

  const contentMap = {
    panel: <Panel products={products} profile={profile} setCurrent={setCurrent}/>,
    ia: <AssistantAI profile={profile} products={products} store={store} onNavigate={setCurrent}/>,
    ventas: <POS products={products} reloadProducts={reload} customers={customers} profile={profile} store={store} cashSession={cashSession} menuOpen={open} onGoReceipts={() => setCurrent('comprobantes')}/>,
    comprobantes: <ReceiptsPage profile={profile} store={store}/>,
    productos: <Products products={products} reload={reload} profile={profile} categories={categories} subcategories={subcategories} reloadCategories={reloadCategories}/>,
    catalogo: <CatalogAdmin products={products} reload={reload} profile={profile} store={store}/>,
    pedidos: <WhatsAppCRM profile={profile} store={store} products={products} onNavigate={setCurrent}/>,
    precios: <PricesAdmin products={products} reload={reload} profile={profile}/>,
    categorias: <CategoriesAdmin profile={profile} categories={categories} subcategories={subcategories} products={products} reloadCategories={reloadCategories}/>,
    etiquetas: <LabelsAdmin products={products} categories={categories} subcategories={subcategories} store={store}/>,
    inventario: <Inventory products={products}/>,
    reportes: <Reports products={products} profile={profile}/>,
    creditos: <Credits profile={profile} cashSession={cashSession}/>,
    caja: <CashPage profile={profile} cashSession={cashSession} />,
    ingreso: <StockEntry products={products} reloadProducts={reload} profile={profile} cashSession={cashSession}/>,
    clientes: <Customers customers={customers} reload={reloadCustomers} profile={profile}/>,
    usuarios: <UsersAdmin profile={profile}/>,
    tienda: <StoreSettings store={store} reloadProfile={reloadProfile}/>,
    herramientas: <ToolsAdmin profile={profile} products={products} categories={categories} subcategories={subcategories} reloadProducts={reload} reloadCustomers={reloadCustomers}/>,
  };
  const content = canAccess(profile?.role, current) ? contentMap[current] : <AccessDenied profile={profile} setCurrent={setCurrent} />;
  return (
    <div className={`app app-mobile-pro ux-premium-shell role-${profile?.role || 'cajero'}`}>
      <button className={`sidebar-scrim ${open ? 'show' : ''}`} type="button" aria-label="Cerrar menú" onClick={() => setOpen(false)} />
      <Sidebar current={current} setCurrent={setCurrent} open={open} setOpen={setOpen} session={session} profile={profile} store={store}/>
      <main className="main main-pro"><Header setOpen={setOpen} current={current} profile={profile} store={store} setCurrent={setCurrent}/>{loading ? <div className="loader">Cargando...</div> : content}</main>
      <MobileBottomNav current={current} setCurrent={setCurrent} role={profile?.role || 'cajero'} menuOpen={open} onAssistant={() => setMobileAssistantOpen(true)} />
      <MobileQuickAssistant open={mobileAssistantOpen} onClose={() => setMobileAssistantOpen(false)} products={products} setCurrent={setCurrent} />
    </div>
  );
}
function Root() {
  const { session, loading } = useAuth();
  const [publicCatalog, setPublicCatalog] = useState(() => isPublicCatalogLocation());
  useEffect(() => {
    const onHash = () => setPublicCatalog(isPublicCatalogLocation());
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);
  useEffect(() => {
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(()=>{});
  }, []);
  if (publicCatalog) return <PublicCatalogApp />;
  if (loading) return <div className="loader full">Iniciando Clomar Store...</div>;
  if (!hasSupabaseConfig) return <Login />;
  return session ? <AppShell session={session}/> : <Login/>;
}

createRoot(document.getElementById('root')).render(<Root />);
