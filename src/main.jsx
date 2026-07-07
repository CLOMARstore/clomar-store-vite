/* Clomar Store V03.0-R3 — Automatización operativa transaccional */
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
const APP_VERSION = 'V03.0-R3 · Operación automática';
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
  a4_2x6: { key: 'a4_2x6', label: 'A4 · 2 columnas × 6 filas', paper: 'a4', columns: 2, rows: 6, width: 90, height: 43, gapX: 5, gapY: 3, density: 'large' },
  a4_3x8: { key: 'a4_3x8', label: 'A4 · 3 columnas × 8 filas', paper: 'a4', columns: 3, rows: 8, width: 58, height: 32, gapX: 4, gapY: 2, density: 'medium' },
  a4_4x10: { key: 'a4_4x10', label: 'A4 · 4 columnas × 10 filas', paper: 'a4', columns: 4, rows: 10, width: 43, height: 25, gapX: 2, gapY: 2, density: 'compact' },
  roll_1col: { key: 'roll_1col', label: 'Rollo térmico · 1 columna', paper: 'roll', columns: 1, rows: 1, width: 60, height: 40, gapX: 0, gapY: 2, density: 'large' },
};

const labelPrintEsc = (value = '') => escapeHtml(String(value ?? ''));
const splitEvery = (items, chunkSize) => {
  const groups = [];
  for (let i = 0; i < items.length; i += chunkSize) groups.push(items.slice(i, i + chunkSize));
  return groups;
};

const buildLabelsPrintHTML = ({
  items = [],
  store = {},
  mode = 'both',
  showPrice = true,
  showLogo = true,
  showCodeText = true,
  sheetLayout = 'a4_3x8',
  labelStyle = 'medium',
}) => {
  const layout = LABEL_LAYOUTS[sheetLayout] || LABEL_LAYOUTS.a4_3x8;
  const perPage = layout.paper === 'a4' ? layout.columns * layout.rows : 1;
  const pages = splitEvery(items, perPage);
  const logo = publicAssetUrl(store?.logo_url || APP_ICON);
  const storeName = store?.name || 'Clomar Store';
  const styleClass = labelStyle === 'small' ? 'style-compact' : labelStyle === 'large' ? 'style-detailed' : 'style-standard';
  const labelMarkup = ({ product }) => {
    const code = productScanCode(product);
    const priceIsReady = productPriceStatus(product) === 'Validado' && Number(product?.price || 0) > 0;
    const hasQr = mode === 'qr' || mode === 'both';
    const hasBarcode = mode === 'barcode' || mode === 'both';
    const logoBlock = showLogo ? `<div class="brand"><img src="${labelPrintEsc(logo)}" alt="" onerror="this.style.display='none'"/><span>${labelPrintEsc(storeName)}</span></div>` : '';
    const priceBlock = !showPrice ? '' : priceIsReady
      ? `<div class="price">${labelPrintEsc(money(product.price))}</div>`
      : `<div class="price pending">Precio pendiente</div>`;
    const qrBlock = hasQr ? `<img class="qr" src="${labelPrintEsc(qrUrl(code))}" alt="QR ${labelPrintEsc(code)}"/>` : '';
    const barcodeBlock = hasBarcode ? barcodeSvgMarkup(code, 46) : '';
    const codeBlock = showCodeText ? `<div class="code-text">${labelPrintEsc(code)}</div>` : '';
    return `<article class="label ${styleClass} mode-${labelPrintEsc(mode)}">${logoBlock}<div class="name">${labelPrintEsc(product?.name || 'Producto')}</div>${priceBlock}<div class="codes ${hasQr && hasBarcode ? 'codes-both' : ''}">${qrBlock}${barcodeBlock}</div>${codeBlock}</article>`;
  };

  const pagesMarkup = pages.map((pageItems, pageIndex) => {
    const filler = layout.paper === 'a4' ? Array.from({ length: Math.max(0, perPage - pageItems.length) }, (_, i) => `<div class="label empty" aria-hidden="true" data-empty="${i}"></div>`).join('') : '';
    return `<section class="label-page ${layout.paper === 'roll' ? 'roll-page' : 'a4-page'}" data-page="${pageIndex + 1}"><div class="label-grid">${pageItems.map(labelMarkup).join('')}${filler}</div></section>`;
  }).join('');

  const pageCss = layout.paper === 'a4'
    ? `@page { size: A4 portrait; margin: 0; } .label-page{width:210mm;height:297mm;padding:10mm;page-break-after:always;break-after:page;} .label-page:last-child{page-break-after:auto;break-after:auto;}`
    : `@page { size: 62mm auto; margin: 0; } .label-page{width:62mm;min-height:40mm;padding:1mm;page-break-after:always;break-after:page;} .label-page:last-child{page-break-after:auto;break-after:auto;}`;

  return `<!doctype html>
<html lang="es"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Etiquetas · ${labelPrintEsc(storeName)}</title>
<style>
  :root{--label-w:${layout.width}mm;--label-h:${layout.height}mm;--cols:${layout.columns};--gap-x:${layout.gapX}mm;--gap-y:${layout.gapY}mm;}
  *{box-sizing:border-box} html,body{margin:0;padding:0;background:#fff;color:#000;font-family:Arial,Helvetica,sans-serif} body{print-color-adjust:exact;-webkit-print-color-adjust:exact}
  ${pageCss}
  .label-grid{display:grid;grid-template-columns:repeat(var(--cols),var(--label-w));grid-auto-rows:var(--label-h);column-gap:var(--gap-x);row-gap:var(--gap-y);justify-content:center;align-content:start}
  .roll-page .label-grid{justify-content:start;grid-template-columns:var(--label-w)}
  .label{width:var(--label-w);height:var(--label-h);overflow:hidden;border:.22mm solid #b9c0ca;border-radius:1.8mm;background:#fff;padding:1.5mm 1.7mm;display:grid;grid-template-rows:auto minmax(0,1fr) auto auto auto;align-content:start;text-align:center;break-inside:avoid;page-break-inside:avoid}
  .label.empty{visibility:hidden}.brand{min-height:3.4mm;display:flex;align-items:center;justify-content:center;gap:1mm;color:#111827;font-size:7.4px;font-weight:800;line-height:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.brand img{width:3.2mm;height:3.2mm;object-fit:contain;flex:0 0 auto}.name{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:2;overflow:hidden;align-self:center;font-size:9px;line-height:1.08;font-weight:900;color:#05070b;min-height:6.3mm}.price{min-height:4.5mm;margin-top:.4mm;font-size:11.5px;line-height:1;font-weight:950;color:#080d18}.price.pending{font-size:7px;color:#bf360c;text-transform:uppercase;letter-spacing:.02em}.codes{height:11.5mm;display:flex;align-items:center;justify-content:center;gap:1.4mm;min-width:0;margin-top:.5mm}.codes-both{justify-content:space-evenly}.qr{width:11.5mm;height:11.5mm;object-fit:contain;image-rendering:auto}.barcode-svg{width:calc(var(--label-w) - 7mm);height:9.5mm;display:block;fill:#000;background:#fff}.codes-both .barcode-svg{width:calc(var(--label-w) - 21mm);height:9.5mm}.code-text{min-height:2.7mm;margin-top:.3mm;font-size:6.4px;line-height:1;font-weight:750;letter-spacing:.035em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#111827}
  .style-compact .name{font-size:8px}.style-compact .brand{font-size:6.5px;min-height:3mm}.style-compact .brand img{width:2.8mm;height:2.8mm}.style-compact .price{font-size:9.2px;min-height:3.8mm}.style-compact .price.pending{font-size:6px}.style-compact .codes{height:9mm}.style-compact .qr{width:9mm;height:9mm}.style-compact .barcode-svg{height:7.6mm}.style-compact .codes-both .barcode-svg{width:calc(var(--label-w) - 17mm);height:7.6mm}.style-compact .code-text{display:none}
  .style-detailed .name{font-size:10.2px}.style-detailed .price{font-size:13px}.style-detailed .codes{height:13.5mm}.style-detailed .qr{width:13.5mm;height:13.5mm}.style-detailed .barcode-svg{height:11.5mm}.style-detailed .codes-both .barcode-svg{width:calc(var(--label-w) - 25mm);height:11.5mm}.style-detailed .code-text{font-size:7px}
  .mode-qr .qr{width:13mm;height:13mm}.mode-qr .codes{height:13mm}.mode-barcode .barcode-svg{width:calc(var(--label-w) - 6mm);height:11mm}.mode-barcode .codes{height:11mm}
  @media screen{body{background:#f3f4f6;padding:12mm}.label-page{margin:0 auto 12mm;box-shadow:0 2mm 8mm rgba(15,23,42,.18)}}
  @media print{html,body{width:100%;height:auto;background:#fff}.label-page{box-shadow:none;margin:0}.label{border-color:#aeb5bf}}
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
  ventas: ['dueno', 'admin', 'cajero'],
  comprobantes: ['dueno', 'admin', 'cajero'],
  creditos: ['dueno', 'admin', 'cajero'],
  caja: ['dueno', 'admin', 'cajero'],
  reportes: ['dueno', 'admin', 'lectura'],
  productos: ['dueno', 'admin', 'almacen'],
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
  const sections = [
    { title: 'Gestionar negocio', items: [
      ['panel', '📊', 'Panel dueño'],
      ['ventas', '🧾', 'Ventas'],
      ['comprobantes', '🧾', 'Comprobantes'],
      ['creditos', '💳', 'Créditos'],
      ['caja', '💰', 'Caja'],
      ['reportes', '📈', 'Reportes'],
    ]},
    { title: 'Productos e inventario', items: [
      ['productos', '📦', 'Productos'],
      ['precios', '💰', 'Precios'],
      ['categorias', '🏷️', 'Categorías'],
      ['etiquetas', '🏷️', 'Etiquetas'],
      ['inventario', '📘', 'Inventario'],
      ['ingreso', '📥', 'Ingreso mercadería'],
    ]},
    { title: 'Contactos', items: [['clientes', '👥', 'Clientes']]},
    { title: 'Administración', items: [
      ['usuarios', '🧑‍💼', 'Usuarios'],
      ['tienda', '🏪', 'Tienda'],
      ['herramientas', '🛠️', 'Herramientas'],
    ]},
  ].map(section => ({ ...section, items: section.items.filter(([key]) => canAccess(role, key)) })).filter(section => section.items.length);
  return (
    <aside className={`sidebar ${open ? 'open' : ''}`}>
      <div className="sidebar-head">
        <div className="mini-logo"><img src={logoSrc(store)} alt="Logo tienda" /></div>
        <div>
          <strong>{store?.name || 'Clomar Store Pro'}</strong>
          <small>{profile?.full_name || session?.user?.email || 'Usuario'} · {roleMeta(profile)}</small>
        </div>
        <button className="ghost mobile-only" onClick={() => setOpen(false)}><X size={18}/></button>
      </div>
      {sections.map((section) => (
        <div key={section.title} className="menu-section">
          <span>{section.title}</span>
          {section.items.map(([key, icon, label]) => (
            <button key={key} className={current === key ? 'active' : ''} onClick={() => { setCurrent(key); setOpen(false); }}>
              <span>{icon}</span>{label}
            </button>
          ))}
        </div>
      ))}
      <button className="logout" onClick={() => supabase?.auth.signOut()}><LogOut size={16}/> Cerrar sesión</button>
    </aside>
  );
}

function Header({ setOpen, current, profile, store }) {
  const titleMap = {
    panel: 'Panel dueño', ventas: 'Venta rápida', comprobantes: 'Comprobantes', creditos: 'Créditos', caja: 'Caja diaria', reportes: 'Reportes', productos: 'Productos', precios: 'Control de precios', categorias: 'Categorías', etiquetas: 'Etiquetas QR y barras', inventario: 'Inventario', ingreso: 'Compras y proveedores', clientes: 'Clientes', usuarios: 'Usuarios y roles', tienda: 'Configuración de tienda', herramientas: 'Herramientas'
  };
  return (
    <header className="app-header app-header-pro">
      <button className="ghost mobile-only menu-toggle-pro" type="button" onClick={() => setOpen(true)}><Menu/></button>
      <div className="header-brand-mobile"><img src={logoSrc(store)} alt="Logo tienda" /></div>
      <div className="header-title-block"><h2>{titleMap[current]}</h2><p>{store?.name || 'Clomar Store Pro'} · {roleMeta(profile)}</p></div>
      <div className="header-status-chip"><span className="status-dot" /><small>{APP_VERSION}</small></div>
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
      .select('id,code,name,category,subcategory,category_id,subcategory_id,price,cost,stock,stock_min,status,store_id,image_url,image_path,brand,size,color,description,barcode,active,price_status,margin_target,min_price,price_notes,price_updated_at,price_updated_by')
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

function Panel({ products, profile }) {
  const { sales } = useSales(profile);
  const { movements } = useCashMovements(profile);
  const stockCritico = products.filter(p => asNum(p.stock) <= asNum(p.stock_min ?? 2));
  const sinPrecio = products.filter(p => asNum(p.price) <= 0 || productPriceStatus(p) !== 'Validado');
  const today = todayISO();
  const salesToday = sales.filter(s => String(s.created_at || '').slice(0,10) === today);
  const yesterday = (() => { const d = new Date(); d.setDate(d.getDate() - 1); return new Date(d.getTime() - d.getTimezoneOffset()*60000).toISOString().slice(0,10); })();
  const salesYesterday = sales.filter(s => String(s.created_at || '').slice(0,10) === yesterday);
  const ingresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Ingreso','Apertura'].includes(m.type));
  const egresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Egreso','Compra','Retiro','Compra crédito'].includes(m.type));
  const totalVentas = salesToday.reduce((s, v) => s + asNum(v.total), 0);
  const totalYesterday = salesYesterday.reduce((s, v) => s + asNum(v.total), 0);
  const cajaNeta = ingresosToday.reduce((s, v) => s + asNum(v.amount), 0) - egresosToday.reduce((s, v) => s + asNum(v.amount), 0);
  const creditosHoy = salesToday.filter(s => s.payment_method === 'Crédito' || s.status === 'Crédito').reduce((s,v)=>s+asNum(v.total),0);
  const ventaMayor = salesToday.slice().sort((a,b)=>asNum(b.total)-asNum(a.total))[0];
  const tendencia = totalYesterday > 0 ? ((totalVentas - totalYesterday) / totalYesterday) * 100 : 0;
  const byMethod = salesToday.reduce((acc, s) => { acc[s.payment_method || 'Sin método'] = (acc[s.payment_method || 'Sin método'] || 0) + asNum(s.total); return acc; }, {});
  const bestMethod = Object.entries(byMethod).sort((a,b)=>b[1]-a[1])[0];
  const alerts = [
    stockCritico.length ? `${stockCritico.length} producto(s) con stock crítico.` : '',
    sinPrecio.length ? `${sinPrecio.length} producto(s) pendientes de precio.` : '',
    creditosHoy > 0 ? `Créditos del día por ${money(creditosHoy)}.` : '',
    cajaNeta < 0 ? 'Caja neta negativa: revisar egresos.' : '',
  ].filter(Boolean);
  return (
    <div className="page compact-page dashboard-owner-page">
      <div className="hero compact-hero owner-hero">
        <div>
          <span className="eyebrow">Vista de dueño</span>
          <h1>📊 Panel comercial</h1>
          <p>Lectura rápida para decidir ventas, reposición, precios, caja y créditos.</p>
        </div>
        <div className="owner-today-card"><span>Ventas hoy</span><strong>{money(totalVentas)}</strong><small>{salesToday.length} comprobantes · {tendencia >= 0 ? '+' : ''}{tendencia.toFixed(1)}% vs ayer</small></div>
      </div>
      <div className="kpi-grid dashboard-kpis">
        <Kpi label="Ventas hoy" value={money(totalVentas)} helper={`${salesToday.length} comprobantes`} />
        <Kpi label="Caja neta" value={money(cajaNeta)} helper="Ingresos menos egresos" />
        <Kpi label="Crédito hoy" value={money(creditosHoy)} helper="Por cobrar" />
        <Kpi label="Método dominante" value={bestMethod?.[0] || '—'} helper={bestMethod ? money(bestMethod[1]) : 'Sin ventas'} />
        <Kpi label="Productos" value={products.length} helper="Activos" />
        <Kpi label="Stock crítico" value={stockCritico.length} helper="Revisar reposición" />
      </div>
      <div className="dashboard-grid-pro">
        <section className="card compact-card"><h3>Acciones recomendadas</h3>{alerts.length ? alerts.map((a, i)=><div className="insight-row" key={i}><span>⚠️</span><strong>{a}</strong></div>) : <div className="insight-row ok"><span>✅</span><strong>Sin alertas críticas por ahora.</strong></div>}</section>
        <section className="card compact-card"><h3>Mejor comprobante del día</h3>{ventaMayor ? <div className="featured-sale"><span>{receiptNumber(ventaMayor)}</span><strong>{money(ventaMayor.total)}</strong><small>{ventaMayor.customer_name || 'Cliente'} · {ventaMayor.payment_method}</small></div> : <p className="muted">Todavía no hay ventas hoy.</p>}</section>
        <section className="card compact-card"><h3>Últimas ventas</h3>{sales.slice(0, 8).length ? sales.slice(0, 8).map(s => (<div className="list-row" key={s.id}><span>{receiptNumber(s)} · {s.customer_name || 'Cliente'}<small>{fmtDate(s.created_at)} · {s.payment_method}</small></span><strong>{money(s.total)}</strong></div>)) : <p className="muted">Todavía no hay ventas registradas.</p>}</section>
        <section className="card compact-card"><h3>Stock crítico</h3>{stockCritico.length ? stockCritico.slice(0, 10).map(p => (<div className="list-row" key={p.id}><span>{p.name}<small>{p.category || 'General'} · mínimo {asNum(p.stock_min)}</small></span><strong className="danger-text">{asNum(p.stock)}</strong></div>)) : <p className="muted">No hay productos críticos.</p>}</section>
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
  const [mixedPayments, setMixedPayments] = useState({ Efectivo: '', Yape: '', Plin: '', Transferencia: '', Tarjeta: '' });
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const searchInputRef = useRef(null);
  const normalized = query.trim().toLowerCase();
  const activeProducts = useMemo(() => products.filter(p => p.active !== false), [products]);
  const fiscalMeta = documentMeta(documentType);
  const isDefaultCheckoutCustomer = !customer || ['Cliente', 'Cliente general', 'Consumidor final'].includes(customer);
  const matches = useMemo(() => {
    const base = !normalized ? activeProducts.slice(0, 12) : activeProducts.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(normalized)).slice(0, 20);
    return base;
  }, [activeProducts, normalized]);
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
  useEffect(() => { if (menuOpen) setMobileCartOpen(false); }, [menuOpen]);
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

  function addProduct(product) {
    if (asNum(product.stock) <= 0) return setNotice({ type: 'warning', icon: '📦', title: 'Producto sin stock disponible', message: `${product.name} no tiene stock para vender. Ingresa a Inventario o Ingreso de mercadería para actualizar existencias.` });
    if (asNum(product.price) <= 0) return setNotice({ type: 'warning', icon: '💰', title: 'Producto sin precio de venta', message: `${product.name} tiene precio 0. Valida costo y precio antes de vender.` });
    if (productPriceStatus(product) !== 'Validado') return setNotice({ type: 'warning', icon: '🔎', title: 'Precio pendiente de validar', message: `${product.name} todavía está marcado como ${productPriceStatus(product)}. Ingresa a Precios, confirma costo y precio, y marca el producto como Validado.` });
    setLastTicket(null);
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
      <div className="hero compact-hero"><h1>🧾 Checkout fiscal</h1><p>Venta interna, boleta y factura preparadas para una futura integración segura con PSE/OSE.</p></div>
      {lastTicket && <LastReceiptBanner ticket={lastTicket} store={store} profile={profile} onOpen={() => { setDismissedTicketId(null); setSaleModal(lastTicket); }} onGoReceipts={onGoReceipts} onDismiss={() => { setDismissedTicketId(lastTicket?.sale?.id || null); clearLastTicketBackup(); setLastTicket(null); setSaleModal(null); }} />}
      <div className="pos-layout">
        <section className="card compact-card">
          <div className="barcode-tools"><div className="search-box barcode-search"><Search size={18}/><input ref={searchInputRef} value={query} onChange={(e)=>setQuery(e.target.value)} onKeyDown={handleSearchKeyDown} placeholder="Buscar o escanear código de barras..." autoFocus /></div><button className="secondary-btn scan-btn" type="button" onClick={()=>setScanOpen(true)}>📷 Escanear con celular</button></div>
          <div className="scanner-help">Lector físico: enfoca el buscador y escanea. Cámara: abre el escáner y apunta al código.</div>
          {scanStatus && <div className="scan-status">{scanStatus}</div>}
          {scanOpen && <div className="scanner-panel"><div className="scanner-head"><strong>Escáner con cámara</strong><button className="icon-btn" type="button" onClick={stopScanner}>×</button></div><div className="scanner-frame"><video ref={videoRef} muted playsInline /></div><p className="muted">Usa la cámara trasera del celular. Si no detecta, escribe el código manualmente en el buscador.</p></div>}
          <div className="product-list">{matches.map(product => {
            const cartItem = cart.find(item => item.id === product.id);
            return <button key={product.id} type="button" className={`product-row product-row-media ${cartItem ? 'product-in-cart' : ''}`} onClick={() => addProduct(product)} aria-label={`Agregar ${product.name} al carrito`}>
              <img className="product-thumb" src={productImageSrc(product)} alt={product.name} />
              <div className="product-row-info"><strong>{product.name}</strong><small>{product.code} · {product.category}{product.subcategory ? ` / ${product.subcategory}` : ''} · {product.brand || 'Sin marca'} · Stock {asNum(product.stock)}</small><span className={priceBadgeClass(productPriceStatus(product))}>{productPriceStatus(product)}</span>{cartItem && <em className="product-cart-badge">En carrito · {cartItem.qty}</em>}</div>
              <b>{money(product.price)}</b>
            </button>;
          })}{!matches.length && <p className="muted">No se encontraron productos.</p>}</div>
        </section>
        <aside className={`card compact-card cart-card pro-cart-card cart-mobile-sheet ${mobileCartOpen ? 'mobile-sheet-open' : ''}`}>
          <button className="sheet-close-btn cart-sheet-close" type="button" aria-label="Cerrar carrito" title="Cerrar carrito" onClick={() => setMobileCartOpen(false)}>×</button>
          <h3><ShoppingCart size={20}/> Carrito</h3>
          <div className="checkout-sheet-scroll">
            {cart.length === 0 ? <div className="empty-checkout-state"><strong>Aún no hay productos</strong><span>Busca, escanea o toca un producto para armar la venta.</span><button type="button" className="secondary-btn" onClick={() => setMobileCartOpen(false)}>Agregar productos</button></div> : cart.map(item => (
              <article className="cart-item-premium" key={item.id}>
                <div className="cart-item-head"><div><strong>{item.name}</strong><small>{money(item.price)} c/u · Stock {asNum(item.stock)}</small></div><button type="button" className="cart-remove-btn" aria-label={`Quitar ${item.name}`} title="Quitar producto" onClick={()=>removeItem(item.id)}>×</button></div>
                <div className="cart-item-controls"><div className="cart-control-field"><span>Cant.</span><div className="quantity-stepper"><button type="button" aria-label="Restar unidad" onClick={()=>updateQty(item.id, asNum(item.qty)-1)}>−</button><input type="number" value={item.qty} min="1" max={asNum(item.stock)} onChange={(e)=>updateQty(item.id, e.target.value)} /><button type="button" aria-label="Sumar unidad" onClick={()=>updateQty(item.id, asNum(item.qty)+1)}>+</button></div></div>{showItemDiscounts || asNum(item.discount) > 0 ? <label className="cart-control-field">Desc.<input value={item.discount || ''} inputMode="decimal" onChange={(e)=>updateItemDiscount(item.id, e.target.value)} placeholder="0.00" /></label> : <button type="button" className="add-line-discount" onClick={()=>setShowItemDiscounts(true)}>Descuento ítem</button>}<div className="cart-item-total"><span>Importe</span><strong>{money(lineSubtotal(item))}</strong></div></div>
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
              <button className="primary-btn checkout-submit-btn" disabled={saving} onClick={submitCheckout}>{checkoutButtonLabel}</button>
            </div>
          </footer>
        </aside>
      </div>
      {cart.length > 0 && !mobileCartOpen && !menuOpen && !customerQuickOpen && !customerPickerOpen && !confirmOpen && !saleModal && <div className="mobile-checkout-bar" role="status" aria-label={`${cart.length} producto(s) en el carrito por ${money(total)}`}>
        <div className="mobile-cart-summary"><small>Carrito activo</small><strong>{cart.length} producto(s) · {money(total)}</strong></div>
        <button type="button" className="mobile-cart-clear" onClick={clearCart}>Vaciar</button>
        <button type="button" className="primary-btn mobile-cart-open-btn" onClick={() => setMobileCartOpen(true)}>Ver carrito</button>
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
  const emptyForm = { code:'', barcode:'', name:'', category_id:'', subcategory_id:'', category:'', subcategory:'', brand:'', size:'', color:'', description:'', price:'', cost:'', stock:'0', stock_min:'2', image_url:'', price_status:'Pendiente', margin_target:'50', min_price:'0', price_notes:'' };
  const [form, setForm] = useState(emptyForm);
  const [imageFile, setImageFile] = useState(null);
  const [previewUrl, setPreviewUrl] = useState('');
  const [saving, setSaving] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [query, setQuery] = useState('');
  const filteredProducts = products.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(query.toLowerCase()));
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
        <form className={`card form-grid product-form-sheet ${formOpen ? 'form-open' : ''}`} onSubmit={save}>
          <button className="sheet-close-btn form-sheet-close" type="button" onClick={() => setFormOpen(false)}>Cerrar ×</button>
          <h3>Nuevo producto</h3>
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
          <label>Descripción<input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="Material, modelo, detalles para catálogo" /></label>
          <div className="form-split">
            <label>Precio de venta<input value={form.price} inputMode="decimal" onChange={e=>setForm({...form,price:e.target.value})} placeholder="0.00" /></label>
            <label>Costo de compra<input value={form.cost} inputMode="decimal" onChange={e=>setForm({...form,cost:e.target.value})} placeholder="0.00" /></label>
          </div>
          <div className="form-split">
            <label>Estado del precio<select value={form.price_status} onChange={e=>setForm({...form,price_status:e.target.value})}>{PRICE_STATUS_OPTIONS.map(o=><option key={o}>{o}</option>)}</select></label>
            <label>Margen objetivo sobre costo %<input value={form.margin_target} inputMode="decimal" onChange={e=>setForm({...form,margin_target:e.target.value})} placeholder="50" /></label>
          </div>
          <div className="form-split">
            <label>Precio mínimo permitido<input value={form.min_price} inputMode="decimal" onChange={e=>setForm({...form,min_price:e.target.value})} placeholder="0.00" /></label>
            <label>Precio sugerido<input value={money(suggestedPrice(form.cost, form.margin_target))} readOnly /></label>
          </div>
          <label>Nota de precio<input value={form.price_notes} onChange={e=>setForm({...form,price_notes:e.target.value})} placeholder="Ej.: falta confirmar precio real con proveedor" /></label>
          <div className="form-split">
            <label>Stock inicial<input value={form.stock} inputMode="decimal" onChange={e=>setForm({...form,stock:e.target.value})} /></label>
            <label>Stock mínimo<input value={form.stock_min} inputMode="decimal" onChange={e=>setForm({...form,stock_min:e.target.value})} /></label>
          </div>
          <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar producto'}</button>
        </form>
        <section className="card compact-card">
          <h3>Lista de productos</h3>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar por nombre, marca, código o barcode..." /></div>
          <div className="product-card-list">
            {filteredProducts.map(p=>(
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
  const active = products.filter(p => p.active !== false);
  const lowStock = active.filter(p => asNum(p.stock) <= asNum(p.stock_min));
  const noStock = active.filter(p => asNum(p.stock) <= 0);
  const filtered = active.filter(p => `${p.code || ''} ${p.barcode || ''} ${p.name || ''} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''}`.toLowerCase().includes(query.toLowerCase()));
  const byCat = filtered.reduce((acc, p) => { (acc[p.category || 'General'] ||= []).push(p); return acc; }, {});
  return (
    <div className="page inventory-page">
      <div className="hero compact-hero"><h1>📘 Inventario</h1><p>Control compacto por categoría, stock disponible y alertas de reposición.</p></div>
      <section className="card compact-card inventory-control-card">
        <div className="inventory-control-head">
          <div><h3>Resumen de inventario</h3><p className="muted">Busca, revisa stock bajo y valida disponibilidad por categoría.</p></div>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar producto, código o categoría..." /></div>
        </div>
        <div className="inventory-kpi-grid">
          <Kpi label="Productos" value={active.length} helper="activos" />
          <Kpi label="Stock bajo" value={lowStock.length} helper="requieren revisión" />
          <Kpi label="Sin stock" value={noStock.length} helper="reposición urgente" />
          <Kpi label="Categorías" value={Object.keys(byCat).length} helper="con resultados" />
        </div>
      </section>
      {Object.entries(byCat).map(([cat, items]) => (
        <section className="card compact-card inventory-block inventory-pro-block" key={cat}>
          <div className="inventory-category-head"><h3>{cat}</h3><span>{items.length} producto(s)</span></div>
          <div className="inventory-card-grid">
            {items.map(p => {
              const low = asNum(p.stock) <= asNum(p.stock_min);
              return (
                <article className="inventory-card-pro" key={p.id}>
                  <img className="product-thumb small" src={productImageSrc(p)} alt={p.name}/>
                  <div className="inventory-card-info">
                    <strong>{p.name}</strong>
                    <small>{p.code || 'Sin código'} · {p.brand || 'Sin marca'}{p.color ? ` · ${p.color}` : ''}</small>
                  </div>
                  <div className={low ? 'stock-pill stock-low' : 'stock-pill'}>Stock {asNum(p.stock)}</div>
                </article>
              );
            })}
          </div>
        </section>
      ))}
      {!filtered.length && <section className="card compact-card"><p className="muted">No se encontraron productos con ese criterio.</p></section>}
    </div>
  );
}

function StockEntry({ products, reloadProducts, profile, cashSession }) {
  const { movements, reload: reloadMovements } = useStockMovements(profile);
  const [form, setForm] = useState({ product_id:'', provider:'', qty:'1', cost:'0', method:'Efectivo', paid:'0', invoice:'', note:'' });
  const [query, setQuery] = useState('');
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
      <div className="hero compact-hero"><h1>📥 Compras y proveedores</h1><p>Registra ingresos de mercadería, proveedor, documento, costo y salida de caja.</p></div>
      <div className="kpi-grid"><Kpi label="Ingresos hoy" value={entriesToday.length} helper="compras registradas" /><Kpi label="Unidades hoy" value={unitsToday} helper="stock agregado" /><Kpi label="Proveedores recientes" value={lastProviders.length} helper="según historial" /><Kpi label="Producto seleccionado" value={selected ? asNum(selected.stock) : 0} helper="stock actual" /></div>
      <div className="two-col">
        <form className="card form-grid purchase-form" onSubmit={save}>
          <h3>Nueva compra / ingreso</h3>
          <label>Producto<select value={form.product_id} onChange={e=>{const p=products.find(x=>x.id===e.target.value); setForm({...form,product_id:e.target.value,cost:String(p?.cost || 0)})}}>{products.map(p=><option key={p.id} value={p.id}>{p.code} · {p.name} · Stock {asNum(p.stock)}</option>)}</select></label>
          <div className="form-split"><label>Proveedor<input value={form.provider} onChange={e=>setForm({...form,provider:e.target.value})} placeholder="Nombre del proveedor" /></label><label>Documento<input value={form.invoice} onChange={e=>setForm({...form,invoice:e.target.value})} placeholder="Factura, boleta, guía" /></label></div>
          <div className="form-split"><label>Cantidad ingresada<input value={form.qty} onChange={e=>setForm({...form,qty:e.target.value})} inputMode="decimal" /></label><label>Costo unitario<input value={form.cost} onChange={e=>setForm({...form,cost:e.target.value})} inputMode="decimal" /></label></div>
          <div className="form-split"><label>Método de pago<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option><option>Crédito</option></select></label><label>Monto pagado<input value={form.paid} onChange={e=>setForm({...form,paid:e.target.value})} inputMode="decimal" /></label></div>
          <label>Observación<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Detalle de compra, cambio de costo, referencia..." /></label>
          <div className="total-box"><span>Total compra</span><strong>{money(total)}</strong><small>Stock después: {selected ? asNum(selected.stock) + asNum(form.qty) : 0}</small></div>
          <button className="primary-btn">Registrar compra e ingreso</button>
        </form>
        <section className="card compact-card"><h3>Producto seleccionado</h3>{selected ? <><div className="list-row inventory-product-row"><img className="product-thumb small" src={productImageSrc(selected)} alt={selected.name}/><span>{selected.name}<small>{selected.code} · {selected.category}{selected.subcategory ? ` / ${selected.subcategory}` : ''} · {selected.brand || 'Sin marca'}</small></span><strong>{money(selected.price)}</strong></div><div className="list-row"><span>Stock actual</span><strong>{asNum(selected.stock)}</strong></div><div className="list-row"><span>Stock después</span><strong>{asNum(selected.stock) + asNum(form.qty)}</strong></div><div className="list-row"><span>Costo anterior</span><strong>{money(selected.cost)}</strong></div><div className="list-row"><span>Nuevo costo</span><strong>{money(form.cost)}</strong></div></> : <p className="muted">No hay productos.</p>}<h4>Proveedores recientes</h4>{lastProviders.map(p => <button key={p} type="button" className="provider-chip" onClick={()=>setForm({...form, provider:p})}>{p}</button>)}</section>
      </div>
      <section className="card compact-card extra-row"><div className="report-filter-head"><div><h3>Historial de compras</h3><p className="muted">Busca por producto, código, proveedor o documento.</p></div><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar compra..." /></div>{filteredEntries.map(m=><div className="list-row" key={m.id}><span>{m.products?.name || 'Producto'}<small>{fmtDate(m.created_at)} · {m.note || 'Sin nota'}</small></span><strong>+{asNum(m.qty)}</strong></div>)}{!filteredEntries.length && <p className="muted">No hay compras registradas.</p>}</section>
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
      <div className="hero compact-hero">
        <div><h1>💰 Caja por turno</h1><p>Las ventas nuevas alimentan caja, inventario y reportes automáticamente.</p></div>
        <span className={session ? 'cash-session-badge open' : 'cash-session-badge closed'}>{session ? 'Caja abierta' : 'Caja cerrada'}</span>
      </div>
      {movementsError && <div className="data-error"><strong>No se pudieron leer algunos movimientos:</strong> {movementsError}</div>}
      {!session ? (
        <section className="card compact-card cash-open-card">
          <div className="cash-state-copy"><span className="eyebrow">Inicio de turno</span><h3>Abra caja antes de vender</h3><p className="muted">El fondo inicial quedará registrado y toda venta posterior se asociará a este turno.</p></div>
          <form className="form-grid cash-open-form" onSubmit={openCash}>
            <label>Fondo inicial en efectivo<input value={opening.amount} onChange={e => setOpening({ ...opening, amount: e.target.value })} inputMode="decimal" placeholder="0.00" /></label>
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
  const [mode, setMode] = useState('both');
  const [labelSize, setLabelSize] = useState('medium');
  const [sheetLayout, setSheetLayout] = useState('a4_3x8');
  const [showPrice, setShowPrice] = useState(true);
  const [showLogo, setShowLogo] = useState(true);
  const [showCodeText, setShowCodeText] = useState(true);
  const [defaultQty, setDefaultQty] = useState(1);
  const [selected, setSelected] = useState(new Set());
  const [quantities, setQuantities] = useState({});

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

  const sheetInfo = {
    a4_2x6: 'A4: 2 columnas x 6 filas',
    a4_3x8: 'A4: 3 columnas x 8 filas',
    a4_4x10: 'A4: 4 columnas x 10 filas',
    roll_1col: 'Rollo térmico: 1 columna',
  }[sheetLayout];

  function toggleProduct(id) {
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }
  function selectFiltered() { setSelected(new Set(filtered.map(p => p.id))); }
  function clearSelection() { setSelected(new Set()); }
  function setQty(id, value) {
    const qty = Math.max(1, Math.min(500, Number(value || 1)));
    setQuantities(prev => ({ ...prev, [id]: qty }));
  }
  function applyQtyToFiltered() {
    const qty = Math.max(1, Math.min(500, Number(defaultQty || 1)));
    const next = { ...quantities };
    filtered.forEach(p => { next[p.id] = qty; });
    setQuantities(next);
  }
  function applyStockQtyToFiltered() {
    const next = { ...quantities };
    filtered.forEach(p => { next[p.id] = Math.max(1, Math.min(500, Number(p.stock || 1))); });
    setQuantities(next);
  }
  function printLabels() {
    if (!printableItems.length) return alert('No hay productos para imprimir.');
    openLabelsPrintWindow({
      items: printableItems,
      store,
      mode,
      showPrice,
      showLogo,
      showCodeText,
      sheetLayout,
      labelStyle: labelSize,
    });
  }

  return (
    <div className="page labels-page">
      <div className="hero compact-hero"><h1>🏷️ Etiquetas QR y código de barras</h1><p>Genera varias etiquetas por hoja PDF, con cantidad por producto y formato A4 o térmico.</p></div>
      <div className="tool-grid labels-tools">
        <section className="card compact-card">
          <h3>Filtros</h3>
          <label>Buscar producto<input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Nombre, código, barcode, marca o color" /></label>
          <label>Categoría<select value={categoryId} onChange={e=>{ setCategoryId(e.target.value); setSubcategoryId('all'); }}><option value="all">Todas</option>{categories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
          <label>Subcategoría<select value={subcategoryId} onChange={e=>setSubcategoryId(e.target.value)}><option value="all">Todas</option>{visibleSubcategories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
          <div className="import-summary label-summary"><Kpi label="Filtrados" value={filtered.length} helper="productos" /><Kpi label="Seleccionados" value={chosen.length || filtered.length} helper={chosen.length ? 'manual' : 'por filtro'} /><Kpi label="Etiquetas" value={printableItems.length} helper="a imprimir" /></div>
          <div className="button-row"><button className="secondary-btn" onClick={selectFiltered}>Seleccionar filtrados</button><button className="secondary-btn" onClick={clearSelection}>Limpiar selección</button></div>
        </section>

        <section className="card compact-card">
          <h3>Diseño de hoja</h3>
          <label>Tipo de código<select value={mode} onChange={e=>setMode(e.target.value)}><option value="both">QR + barras</option><option value="qr">Solo QR</option><option value="barcode">Solo código de barras</option></select></label>
          <label>Estilo de contenido<select value={labelSize} onChange={e=>setLabelSize(e.target.value)}><option value="small">Compacto — recomendado para A4 4 × 10</option><option value="medium">Equilibrado — recomendado para A4 3 × 8</option><option value="large">Detallado — recomendado para A4 2 × 6</option></select></label>
          <label>Formato de hoja<select value={sheetLayout} onChange={e=>setSheetLayout(e.target.value)}><option value="a4_2x6">A4: 2 columnas × 6 filas (90 × 43 mm)</option><option value="a4_3x8">A4: 3 columnas × 8 filas (58 × 32 mm)</option><option value="a4_4x10">A4: 4 columnas × 10 filas (43 × 25 mm)</option><option value="roll_1col">Rollo térmico: 1 columna (60 × 40 mm)</option></select></label>
          <label className="check-row"><input type="checkbox" checked={showPrice} onChange={e=>setShowPrice(e.target.checked)} /> Mostrar precio</label>
          <label className="check-row"><input type="checkbox" checked={showLogo} onChange={e=>setShowLogo(e.target.checked)} /> Mostrar marca Clomar Store</label>
          <label className="check-row"><input type="checkbox" checked={showCodeText} onChange={e=>setShowCodeText(e.target.checked)} /> Mostrar código escrito</label>
          <button className="primary-btn" onClick={printLabels}>Imprimir / Guardar PDF</button>
          <p className="muted">Formato: <strong>{sheetInfo}</strong>. Se abrirá una hoja limpia, sin menú ni barra móvil. En el diálogo de impresión seleccione <strong>Guardar como PDF</strong>, papel <strong>A4</strong>, escala <strong>100 %</strong>, márgenes <strong>Ninguno</strong> y encabezados/pies desactivados.</p>
        </section>
      </div>

      <section className="card compact-card no-print">
        <h3>Cantidad por producto</h3>
        <div className="quantity-tools">
          <label>Cantidad rápida<input type="number" min="1" max="500" value={defaultQty} onChange={e=>setDefaultQty(e.target.value)} /></label>
          <button className="secondary-btn" onClick={applyQtyToFiltered}>Aplicar cantidad a filtrados</button>
          <button className="secondary-btn" onClick={applyStockQtyToFiltered}>Usar stock como cantidad</button>
        </div>
        <p className="muted">Ejemplo: si un producto tiene cantidad 10, saldrán 10 etiquetas de ese producto en el mismo PDF.</p>
      </section>

      <section className="card compact-card no-print">
        <h3>Productos para etiquetas</h3>
        <div className="product-pick-list product-pick-list-qty">
          {filtered.map(p => <label key={p.id} className="product-pick-row qty-row"><input type="checkbox" checked={selected.has(p.id)} onChange={()=>toggleProduct(p.id)} /><img src={productImageSrc(p)} alt={p.name}/><span><strong>{p.name}</strong><small>{p.code} · {p.barcode || 'Sin barcode'} · {p.category || 'Sin categoría'}{p.subcategory ? ` / ${p.subcategory}` : ''}</small><small>Stock: {p.stock ?? 0} · Precio: {money(p.price)}</small></span><div className="qty-box"><small>Cant.</small><input type="number" min="1" max="500" value={quantities[p.id] || defaultQty} onChange={e=>setQty(p.id, e.target.value)} /></div></label>)}
          {!filtered.length && <p className="muted">No hay productos con esos filtros.</p>}
        </div>
      </section>

      <section className="card compact-card no-print">
        <h3>Vista previa</h3>
        <p className="muted">Se imprimirán <strong>{printableItems.length}</strong> etiquetas. Si el precio está pendiente, la etiqueta mostrará “Precio pendiente” en lugar del monto.</p>
      </section>

      <div className={`print-label-sheet label-size-${labelSize} sheet-${sheetLayout}`}>
        {printableItems.map(({ product, key }) => {
          const code = productScanCode(product);
          return <div className="print-label" key={key}>
            {showLogo && <div className="label-brand"><img src={APP_ICON} alt="Clomar"/><span>{store?.name || 'Clomar Store'}</span></div>}
            <div className="label-name">{product.name}</div>
            {showPrice && Number(product.price || 0) > 0 && productPriceStatus(product) === 'Validado' && <div className="label-price">{money(product.price)}</div>}
            {showPrice && productPriceStatus(product) !== 'Validado' && <div className="label-pending">Precio pendiente</div>}
            <div className={`label-codes mode-${mode}`}>{(mode === 'qr' || mode === 'both') && <img className="label-qr" src={qrUrl(code)} alt={`QR ${code}`} />}{(mode === 'barcode' || mode === 'both') && <BarcodeSVG value={code} />}</div>
            {showCodeText && <div className="label-code-text">{code}</div>}
          </div>;
        })}
      </div>
    </div>
  );
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
        <label>Teléfono / WhatsApp<input value={form.phone || ''} onChange={e=>setForm({...form, phone:e.target.value})} placeholder="Celular" /></label>
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


function MobileBottomNav({ current, setCurrent, role, menuOpen = false }) {
  const labelMap = {
    panel: ['📊', 'Inicio'], ventas: ['🧾', 'Vender'], reportes: ['📈', 'Reportes'], caja: ['💰', 'Caja'], comprobantes: ['🧾', 'Tickets'], productos: ['📦', 'Productos'], inventario: ['📘', 'Stock'], ingreso: ['📥', 'Compras'], herramientas: ['🛠️', 'Más'], creditos: ['💳', 'Créditos']
  };
  const preferred = role === 'almacen'
    ? ['productos', 'inventario', 'ingreso', 'etiquetas']
    : role === 'cajero'
      ? ['ventas', 'comprobantes', 'caja', 'creditos']
      : role === 'lectura'
        ? ['panel', 'reportes', 'inventario']
        : ['panel', 'ventas', 'reportes', 'caja', 'herramientas'];
  const items = preferred.filter(key => canAccess(role, key) && labelMap[key]).slice(0, 5);
  if (!items.length || menuOpen) return null;
  return (
    <nav className="mobile-bottom-nav" aria-label="Navegación móvil principal">
      {items.map(key => {
        const [icon, label] = labelMap[key];
        return <button key={key} type="button" className={current === key ? 'active' : ''} onClick={() => setCurrent(key)}><span>{icon}</span><small>{label}</small></button>;
      })}
    </nav>
  );
}

function AppShell({ session }) {
  const [current, setCurrent] = useState(() => {
    try { return localStorage.getItem('clomar_last_module') || 'ventas'; } catch (err) { return 'ventas'; }
  });
  const [open, setOpen] = useState(false);
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
    panel: <Panel products={products} profile={profile}/>,
    ventas: <POS products={products} reloadProducts={reload} customers={customers} profile={profile} store={store} cashSession={cashSession} menuOpen={open} onGoReceipts={() => setCurrent('comprobantes')}/>,
    comprobantes: <ReceiptsPage profile={profile} store={store}/>,
    productos: <Products products={products} reload={reload} profile={profile} categories={categories} subcategories={subcategories} reloadCategories={reloadCategories}/>,
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
    <div className="app app-mobile-pro">
      <button className={`sidebar-scrim ${open ? 'show' : ''}`} type="button" aria-label="Cerrar menú" onClick={() => setOpen(false)} />
      <Sidebar current={current} setCurrent={setCurrent} open={open} setOpen={setOpen} session={session} profile={profile} store={store}/>
      <main className="main main-pro"><Header setOpen={setOpen} current={current} profile={profile} store={store}/>{loading ? <div className="loader">Cargando...</div> : content}</main>
      <MobileBottomNav current={current} setCurrent={setCurrent} role={profile?.role || 'cajero'} menuOpen={open} />
    </div>
  );
}
function Root() {
  const { session, loading } = useAuth();
  useEffect(() => {
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(()=>{});
  }, []);
  if (loading) return <div className="loader full">Iniciando Clomar Store...</div>;
  if (!hasSupabaseConfig) return <Login />;
  return session ? <AppShell session={session}/> : <Login/>;
}

createRoot(document.getElementById('root')).render(<Root />);
