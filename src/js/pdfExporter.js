import { sortBankTransfers } from './bankSorter.js';
import { formatOtraFecha } from './excelExporter.js';

/**
 * Formats a time string or Excel fractional time to 12-hour format (e.g., 08:35 am, 02:40 pm).
 * @param {string|number} val 
 * @returns {string}
 */
function formatTime12h(val) {
  if (val === null || val === undefined || String(val).trim() === '' || val === '-') return '-';
  
  // Case 1: Excel numeric time fraction (0.0 to 1.0)
  if (typeof val === 'number' || (!isNaN(val) && !String(val).includes(':'))) {
    const num = Number(val);
    if (num >= 0 && num <= 1) {
      const totalSec = Math.round(num * 86400);
      const h24 = Math.floor(totalSec / 3600) % 24;
      const m = Math.floor((totalSec % 3600) / 60);
      const period = h24 >= 12 ? 'pm' : 'am';
      const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
      return `${String(h12).padStart(2, '0')}:${String(m).padStart(2, '0')} ${period}`;
    }
  }

  // Case 2: String time (HH:MM, HH:MM:SS, with or without AM/PM)
  const str = String(val).trim();
  const match = str.match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(am|pm)?$/i);
  if (match) {
    let h = parseInt(match[1], 10);
    const m = match[2];
    const explicitP = match[3];
    if (explicitP) {
      return `${String(h).padStart(2, '0')}:${m} ${explicitP.toLowerCase()}`;
    } else {
      const p = h >= 12 ? 'pm' : 'am';
      const h12 = h % 12 === 0 ? 12 : h % 12;
      return `${String(h12).padStart(2, '0')}:${m} ${p}`;
    }
  }

  return str;
}

/**
 * Generates an official, printable PDF-ready daily reconciliation report for Plastiluz.
 * Formatted strictly into 2 separate A4 sheets (210mm x 297mm) with a minimalist aesthetic,
 * Arial typography, prominent numbers, selective cell highlighting, and full support for ANULADO.
 * 
 * @param {Object} session { date, documents, bankTransfers, cashCount, cashAdjustments, totals }
 */
export function exportToPdf(session) {
  const { date, documents, bankTransfers, cashCount, cashAdjustments = [], totals } = session;

  const dateObj = new Date(date);
  const optionsLong = {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  };
  const formattedDate = dateObj.toLocaleDateString('es-PE', optionsLong);
  const formattedDateUpper = formattedDate.charAt(0).toUpperCase() + formattedDate.slice(1);
  const formattedDateFilename = dateObj.toISOString().slice(0, 10);
  const printTimestamp = new Date().toLocaleString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  // Sort & Separate documents
  const docSorter = (a, b) => {
    if (a.type === 'NC' && b.type !== 'NC') return 1;
    if (a.type !== 'NC' && b.type === 'NC') return -1;
    const numA = parseInt(a.numero.replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(b.numero.replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  };

  const nvs = documents.filter(d => d.type === 'NV').sort(docSorter);
  const boletas = documents.filter(d => d.type === 'BOLETA' || (d.type === 'NC' && d.serie.toUpperCase().startsWith('B'))).sort(docSorter);
  const facturas = documents.filter(d => d.type === 'FACTURA' || (d.type === 'NC' && d.serie.toUpperCase().startsWith('F'))).sort(docSorter);

  // Helper checks
  const isAnuladoDoc = (d) => d.condicion === 'anulado' || d.metodoPago === 'anulado';
  const isOtraFechaDoc = (d) => d.condicion === 'otra_fecha';
  const isCreditoDoc = (d) => (d.condicion === 'credito' || d.metodoPago === 'credito') && !isAnuladoDoc(d) && !isOtraFechaDoc(d);
  const isBancoDoc = (d) => d.metodoPago === 'banco' && !isAnuladoDoc(d) && !isOtraFechaDoc(d) && !isCreditoDoc(d);

  // Sums (Exclude ANULADO and OTRA FECHA vouchers from sales totals!)
  const sumNvs = nvs.reduce((acc, d) => (isAnuladoDoc(d) || isOtraFechaDoc(d)) ? acc : (d.type === 'NC' ? acc - d.importe : acc + d.importe), 0);
  const sumBoletas = boletas.reduce((acc, d) => (isAnuladoDoc(d) || isOtraFechaDoc(d)) ? acc : (d.type === 'NC' ? acc - d.importe : acc + d.importe), 0);
  const sumFacturas = facturas.reduce((acc, d) => (isAnuladoDoc(d) || isOtraFechaDoc(d)) ? acc : (d.type === 'NC' ? acc - d.importe : acc + d.importe), 0);

  const maxLen = Math.max(nvs.length, boletas.length, facturas.length);

  // Credit sales (Exclude anulado and otra_fecha)
  const creditDocs = documents.filter(d => isCreditoDoc(d));
  const creditTotal = creditDocs.reduce((acc, d) => d.type === 'NC' ? acc - d.importe : acc + d.importe, 0);

  // Adjustments (Ingresos / Gastos)
  const adjTotal = cashAdjustments.reduce((acc, a) => acc + (a.amount || 0), 0);

  // Cash counts
  const coins = cashCount.filter(c => c.value < 10);
  const bills = cashCount.filter(c => c.value >= 10);
  const coinsSum = coins.reduce((acc, c) => acc + c.total, 0);
  const billsSum = bills.reduce((acc, c) => acc + c.total, 0);
  const physicalCashTotal = coinsSum + billsSum;

  // Reconciliation figures (Self-calculated to ensure 100% accuracy and zero drift from voided vouchers)
  const salesTotal = sumNvs + sumBoletas + sumFacturas;
  const bankTotal = bankTransfers.reduce(
    (acc, t) => (t.type === 'NC' ? acc - Math.abs(t.importe) : acc + t.importe),
    0
  );
  const expectedCash = salesTotal - bankTotal - creditTotal + adjTotal;
  const diff = physicalCashTotal - expectedCash;

  // Currency formatter
  const fCur = (num) => {
    const isNeg = num < 0;
    const absStr = Math.abs(num).toFixed(2);
    return isNeg ? `- S/. ${absStr}` : `S/. ${absStr}`;
  };

  // Generate Comprobantes Rows HTML (Minimalist, large Arial numbers, clean backgrounds)
  let comprobantesRowsHtml = '';
  for (let i = 0; i < maxLen; i++) {
    const nv = nvs[i] || null;
    const bol = boletas[i] || null;
    const fac = facturas[i] || null;

    const renderCell = (doc) => {
      if (!doc) return '<td class="empty-cell"></td><td class="empty-cell"></td>';
      const isNC = doc.type === 'NC';
      const isAnulado = isAnuladoDoc(doc);
      const isOtraFecha = isOtraFechaDoc(doc);
      const isCredit = isCreditoDoc(doc);
      const isBank = isBancoDoc(doc);
      
      let badge = '';
      if (isAnulado) badge = ''; // Word ANULADO must only appear in the value cell, not in document number cell!
      else if (isOtraFecha) badge = ''; // Only number in doc cell as requested
      else if (isBank) badge = '<span class="badge badge-bcp">(*)</span> ';
      else if (isCredit) badge = '<span class="badge badge-crd">CRÉDITO</span> ';
      else if (isNC) badge = '<span class="badge badge-nc">NC</span> ';

      const docDisplay = (isNC && !isOtraFecha && !isAnulado) ? `${doc.serie}-${doc.numero}` : doc.numero;
      const amtVal = isNC ? -doc.importe : doc.importe;
      
      let amtFormatted = fCur(amtVal);
      if (isAnulado) {
        amtFormatted = '<span class="text-anulado">ANULADO</span>';
      } else if (isOtraFecha) {
        amtFormatted = `<span class="text-otra-fecha">${formatOtraFecha(doc.fechaOtra, date)}</span>`;
      } else if (isCredit) {
        amtFormatted = '<span class="text-muted">crédito</span>';
      }

      let amtClass = '';
      if (isAnulado) amtClass = 'anulado-cell';
      else if (isOtraFecha) amtClass = 'text-otra-fecha';
      else if (isNC) amtClass = 'text-red';
      else if (isBank) amtClass = 'text-blue';

      return `
        <td class="doc-num-col ${isNC ? 'nc-row' : ''} ${isAnulado ? 'anulado-row' : ''} ${isOtraFecha ? 'otra-fecha-row' : ''}">${badge}${docDisplay}</td>
        <td class="doc-amt-col ${amtClass} ${isNC ? 'nc-row' : ''} ${isAnulado ? 'anulado-row' : ''} ${isOtraFecha ? 'otra-fecha-row' : ''}">${amtFormatted}</td>
      `;
    };

    comprobantesRowsHtml += `
      <tr>
        ${renderCell(nv)}
        <td class="spacer-col"></td>
        ${renderCell(bol)}
        <td class="spacer-col"></td>
        ${renderCell(fac)}
      </tr>
    `;
  }

  // Generate Bank Transfers Rows (Sorted chronologically, 12h format, clean minimal cells)
  let bankRowsHtml = '';
  const sortedTransfers = sortBankTransfers(bankTransfers);
  if (sortedTransfers.length === 0) {
    bankRowsHtml = `<tr><td colspan="7" class="empty-msg">No se registraron transferencias bancarias</td></tr>`;
  } else {
    sortedTransfers.forEach((t) => {
      const isNC = t.type === 'NC';
      const amt = isNC ? -Math.abs(t.importe) : t.importe;
      const isNeg = amt < 0;
      const time12 = formatTime12h(t.hora);
      bankRowsHtml += `
        <tr class="${isNeg ? 'nc-row' : ''}">
          <td class="text-center">${t.fecha || '-'}</td>
          <td class="text-center font-medium">${t.type}</td>
          <td class="text-center font-medium">${t.serie ? `${t.serie}-${t.numero}` : t.numero || '-'}</td>
          <td class="text-center font-mono">${t.nroOperacion || '-'}</td>
          <td class="text-center font-mono text-hour">${time12}</td>
          <td class="text-right font-amount ${isNeg ? 'text-red' : 'text-blue'}">${fCur(amt)}</td>
          <td class="text-center"><span class="badge-pill ${t.verificado === 'SI' ? 'pill-green' : 'pill-gray'}">${t.verificado || 'SI'}</span></td>
        </tr>
      `;
    });
  }

  // Generate Cash Inventory Rows (Monedas y Billetes with large Arial digits)
  let cashInventoryRowsHtml = '';
  const maxDenomLen = Math.max(coins.length, bills.length);
  for (let i = 0; i < maxDenomLen; i++) {
    const coin = coins[i] || null;
    const bill = bills[i] || null;

    cashInventoryRowsHtml += `
      <tr>
        <td class="text-center font-medium">${coin ? `S/. ${coin.value.toFixed(2)}` : ''}</td>
        <td class="text-center font-qty">${coin && coin.quantity !== null ? coin.quantity : (coin ? '-' : '')}</td>
        <td class="text-right font-amount">${coin ? fCur(coin.total) : ''}</td>
        <td class="spacer-col"></td>
        <td class="text-center font-medium">${bill ? `S/. ${bill.value.toFixed(2)}` : ''}</td>
        <td class="text-center font-qty">${bill && bill.quantity !== null ? bill.quantity : (bill ? '-' : '')}</td>
        <td class="text-right font-amount">${bill ? fCur(bill.total) : ''}</td>
      </tr>
    `;
  }

  // Adjustments rows
  let adjRowsHtml = '';
  if (cashAdjustments.length > 0) {
    cashAdjustments.forEach((adj) => {
      if (!adj.description && !adj.amount) return;
      adjRowsHtml += `
        <tr>
          <td>${adj.description || 'CONCEPTO CAJA'}</td>
          <td class="text-right font-amount ${adj.amount < 0 ? 'text-red' : 'text-blue'}">${fCur(adj.amount)}</td>
        </tr>
      `;
    });
  }

  // Credit sales rows
  let creditRowsHtml = '';
  if (creditDocs.length > 0) {
    creditDocs.forEach((d) => {
      creditRowsHtml += `
        <tr>
          <td class="text-center font-medium">${d.serie ? `${d.serie}-${d.numero}` : d.numero}</td>
          <td class="text-right font-amount text-amber">${fCur(d.type === 'NC' ? -d.importe : d.importe)}</td>
        </tr>
      `;
    });
  }

  // Difference card info
  let diffBadgeClass = 'diff-perfect';
  let diffTitle = 'CAJA CUADRADA EXACTA';
  let diffDesc = 'El dinero físico y transferencias concilian plenamente con las ventas registradas.';

  if (Math.abs(diff) > 0.015) {
    if (diff > 0) {
      diffBadgeClass = 'diff-sobra';
      diffTitle = `SOBRANTE DE CAJA (+${fCur(diff)})`;
      diffDesc = 'Se ha recaudado más dinero del esperado respecto a las ventas del día.';
    } else {
      diffBadgeClass = 'diff-falta';
      diffTitle = `FALTANTE DE CAJA (${fCur(diff)})`;
      diffDesc = 'Existe una diferencia negativa entre el dinero recaudado y las ventas esperadas.';
    }
  }

  const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Plastiluz - Cuadre de Caja (${formattedDateUpper})</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
  <style>
    /* ============================================================
       MINIMALIST A4 PRINT SETUP (210mm x 297mm) - ARIAL TYPOGRAPHY
       ============================================================ */
    @page {
      size: A4 portrait;
      margin: 8mm 10mm 8mm 10mm;
    }

    *, *::before, *::after {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body, table, th, td, input, select, button {
      font-family: Arial, "Helvetica Neue", Helvetica, sans-serif !important;
    }

    body {
      font-size: 11.5px;
      line-height: 1.35;
      color: #1F2937;
      background-color: #E8E5DF;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }

    /* Screen Control Toolbar */
    .no-print-bar {
      background: #2B2521;
      color: #FFF;
      padding: 10px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 2000;
      box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .bar-info {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .bar-title {
      font-size: 14.5px;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .a4-badge {
      background: #E67E22;
      color: #FFF;
      font-size: 11px;
      font-weight: 700;
      padding: 2.5px 8px;
      border-radius: 4px;
      letter-spacing: 0.5px;
    }
    .print-actions {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .btn-action {
      border: none;
      padding: 7px 15px;
      border-radius: 5px;
      font-weight: 700;
      font-size: 12.5px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s;
    }
    .btn-print {
      background-color: #E67E22;
      color: #FFF;
    }
    .btn-print:hover {
      background-color: #D35400;
      box-shadow: 0 2px 10px rgba(230,126,34,0.4);
    }
    .btn-download {
      background-color: #27AE60;
      color: #FFF;
    }
    .btn-download:hover {
      background-color: #219653;
      box-shadow: 0 2px 10px rgba(39,174,96,0.4);
    }
    .btn-close {
      background-color: transparent;
      color: #CCC;
      border: 1px solid #666;
      padding: 7px 12px;
    }
    .btn-close:hover {
      background-color: #444;
      color: #FFF;
    }

    /* Screen A4 Sheet Representation */
    .a4-page {
      width: 210mm;
      max-width: 210mm;
      min-height: 297mm;
      margin: 16px auto;
      background: #FFF;
      padding: 9mm 11mm 9mm 11mm;
      box-shadow: 0 4px 25px rgba(0,0,0,0.18);
      box-sizing: border-box;
      border-radius: 2px;
      position: relative;
      page-break-after: always;
      break-after: page;
    }
    .a4-page:last-child {
      page-break-after: auto;
      break-after: auto;
    }

    /* Print Styles */
    @media print {
      body {
        background: #FFF !important;
        padding: 0 !important;
        margin: 0 !important;
      }
      .no-print-bar {
        display: none !important;
      }
      .a4-page {
        width: 100% !important;
        max-width: 190mm !important;
        min-height: 281mm !important;
        margin: 0 auto !important;
        padding: 0 !important;
        box-shadow: none !important;
        border: none !important;
        border-radius: 0 !important;
        page-break-after: always !important;
        break-after: page !important;
      }
      .a4-page:last-child {
        page-break-after: auto !important;
        break-after: auto !important;
      }
    }

    /* ============================================================
       MINIMALIST COMPACT HEADER
       ============================================================ */
    .simple-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 2px solid #1F2937;
      padding-bottom: 4px;
      margin-bottom: 8px;
    }
    .header-left {
      display: flex;
      align-items: baseline;
      gap: 10px;
    }
    .header-brand {
      font-size: 20px;
      font-weight: 800;
      color: #E67E22;
      letter-spacing: 0.5px;
    }
    .header-divider {
      color: #D1D5DB;
      font-size: 15px;
    }
    .header-label {
      font-size: 12px;
      font-weight: 700;
      color: #374151;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .header-date {
      font-size: 13.5px;
      font-weight: 700;
      color: #111827;
      text-transform: capitalize;
    }
    .header-right {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .header-sheet-tag {
      background-color: transparent;
      color: #4B5563;
      border: 1px solid #D1D5DB;
      font-size: 10px;
      font-weight: 700;
      padding: 2.5px 7px;
      border-radius: 3px;
      letter-spacing: 0.5px;
    }

    /* Minimalist Section Titles (Clean line, no heavy background) */
    .section-title {
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #1F2937;
      padding: 3px 0 3px 7px;
      border-left: 3px solid #E67E22;
      margin-top: 8px;
      margin-bottom: 4px;
      page-break-after: avoid;
      break-after: avoid;
    }

    /* ============================================================
       TABLES (Clean, minimalist, no background fills on normal rows)
       ============================================================ */
    table {
      width: 100%;
      table-layout: fixed;
      border-collapse: collapse;
      font-size: 12px;
      margin-bottom: 6px;
      page-break-inside: auto;
    }
    th {
      background-color: transparent;
      color: #374151;
      padding: 4.5px 5px;
      font-weight: 600;
      font-size: 11px;
      border-top: 1.5px solid #1F2937;
      border-bottom: 1.5px solid #1F2937;
      border-left: 1px solid #E5E7EB;
      border-right: 1px solid #E5E7EB;
      text-align: left;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    td {
      background-color: transparent;
      padding: 4.5px 6px;
      font-size: 12px;
      border: 1px solid #E5E7EB;
      vertical-align: middle;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    tr {
      page-break-inside: avoid;
      break-inside: avoid;
    }
    thead {
      display: table-header-group;
    }
    tfoot {
      display: table-footer-group;
    }

    /* Larger Numbers Occupying Space */
    .doc-num-col {
      font-weight: 500;
      font-size: 13px;
      letter-spacing: 0.2px;
    }
    .doc-amt-col {
      text-align: right;
      font-weight: 500;
      font-size: 13.5px;
      letter-spacing: 0.2px;
    }
    .font-amount {
      font-weight: 500;
      font-size: 13.5px;
      letter-spacing: 0.2px;
    }
    .font-qty {
      font-weight: 500;
      font-size: 13px;
    }
    .text-hour {
      color: #374151;
      font-weight: 400;
      font-size: 12px;
    }

    /* Totals Row (Selective background only for relevant totals!) */
    .totals-row td {
      background-color: #F8F9FA !important;
      border-top: 1.5px solid #1F2937 !important;
      border-bottom: 1.5px solid #1F2937 !important;
    }
    .totals-row .font-amount, .totals-row .doc-amt-col {
      font-weight: 700 !important;
      font-size: 14px !important;
      color: #111827 !important;
    }
    .totals-row .text-blue {
      color: #1E40AF !important;
    }

    .spacer-col {
      width: 6px !important;
      padding: 0 !important;
      border-top: none !important;
      border-bottom: none !important;
      background-color: #FFF !important;
    }
    .empty-cell {
      background-color: transparent;
    }
    .text-center { text-align: center; }
    .text-right { text-align: right; }
    .font-bold { font-weight: 700; }
    .font-medium { font-weight: 500; }
    .font-mono { font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, Courier, monospace !important; }
    .text-red { color: #B91C1C !important; }
    .text-blue { color: #1E40AF !important; }
    .text-amber { color: #92400E !important; }
    .text-muted { color: #9CA3AF; font-style: italic; font-weight: normal; }

    /* Special row states (Selective, subtle highlights) */
    .nc-row {
      background-color: #FEF2F2 !important;
    }
    .anulado-row {
      background-color: #F9FAFB !important;
      color: #9CA3AF !important;
    }
    .otra-fecha-row {
      background-color: #FAF5FF !important;
      color: #6B21A8 !important;
    }
    .text-anulado {
      color: #6B7280;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.5px;
    }
    .text-otra-fecha {
      color: #6B21A8 !important;
      font-weight: 600;
    }

    /* Badges (Crisp, clean minimal badges) */
    .badge {
      font-size: 9.5px;
      font-weight: 700;
      padding: 1px 4px;
      border-radius: 3px;
      display: inline-block;
      line-height: 1;
    }
    .badge-bcp {
      color: #1E40AF;
      background-color: #EFF6FF;
      border: 1px solid #DBEAFE;
    }
    .badge-crd {
      color: #92400E;
      background-color: #FEF3C7;
      border: 1px solid #FDE68A;
    }
    .badge-nc {
      color: #991B1B;
      background-color: #FEE2E2;
      border: 1px solid #FECACA;
    }
    .badge-anulado {
      color: #4B5563;
      background-color: #F3F4F6;
      border: 1px solid #E5E7EB;
    }
    .badge-otra-fecha {
      color: #6B21A8;
      background-color: #F3E8FF;
      border: 1px solid #E9D5FF;
    }
    .badge-pill {
      font-size: 10px;
      font-weight: 600;
      padding: 1.5px 6px;
      border-radius: 4px;
    }
    .pill-green {
      background-color: #ECFDF5;
      color: #065F46;
      border: 1px solid #A7F3D0;
    }
    .pill-gray {
      background-color: #F3F4F6;
      color: #4B5563;
      border: 1px solid #E5E7EB;
    }

    /* Grid Layout for Bottom of Page 1 (Side-by-side Credits & Adjustments with Reconciliation) */
    .page1-bottom-grid {
      display: grid;
      grid-template-columns: 46% 54%;
      gap: 6mm;
      margin-top: 6px;
      page-break-inside: avoid;
      break-inside: avoid;
    }

    /* Minimalist Reconciliation Box */
    .recon-box {
      border: 1px solid #D1D5DB;
      border-radius: 4px;
      padding: 8px 12px;
      background-color: #FFFFFF;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .recon-header {
      font-size: 12px;
      font-weight: 700;
      color: #111827;
      margin-bottom: 5px;
      border-bottom: 1.5px solid #1F2937;
      padding-bottom: 3px;
      display: flex;
      justify-content: space-between;
    }
    .recon-row {
      display: flex;
      justify-content: space-between;
      padding: 2.5px 0;
      font-size: 12px;
      color: #4B5563;
    }
    .recon-row .font-amount {
      color: #111827;
      font-weight: 500;
      font-size: 13px;
    }
    .recon-row.highlight {
      background-color: #F3F4F6;
      padding: 3px 6px;
      border-radius: 3px;
      margin: 2px 0;
      font-weight: 700;
      color: #111827;
    }
    .recon-row.highlight .font-amount {
      font-weight: 700;
      font-size: 13.5px;
    }
    .recon-row.final {
      background-color: #1F2937;
      color: #FFF;
      padding: 5px 8px;
      border-radius: 3px;
      margin-top: 5px;
      font-weight: 700;
      font-size: 14px;
    }
    .recon-row.final span {
      color: #FFF;
    }

    /* Difference status (Selective background) */
    .diff-status-box {
      padding: 6px 10px;
      border-radius: 4px;
      text-align: center;
      margin-top: 6px;
      font-weight: 700;
      font-size: 12.5px;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .diff-perfect {
      background-color: #ECFDF5;
      color: #065F46;
      border: 1px solid #A7F3D0;
    }
    .diff-sobra {
      background-color: #EFF6FF;
      color: #1E40AF;
      border: 1px solid #BFDBFE;
    }
    .diff-falta {
      background-color: #FEF2F2;
      color: #991B1B;
      border: 1px solid #FECACA;
    }
    .diff-sub {
      font-size: 10px;
      font-weight: normal;
      margin-top: 1px;
      opacity: 0.95;
    }

    /* Summary card on Page 2 (Clean and minimalist) */
    .bank-summary-card {
      background-color: #FAFAFA;
      border: 1px solid #E5E7EB;
      border-radius: 4px;
      padding: 9px 12px;
      margin-top: 10px;
    }
    .bank-summary-title {
      font-size: 11px;
      font-weight: 700;
      color: #6B7280;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 5px;
    }
    .bank-summary-row {
      display: flex;
      justify-content: space-between;
      font-size: 12px;
      color: #4B5563;
      padding: 2px 0;
    }
    .bank-summary-row .font-amount {
      color: #111827;
      font-size: 13.5px;
      font-weight: 600;
    }
    .bank-summary-row.total-row {
      border-top: 1.5px solid #1F2937;
      margin-top: 4px;
      padding-top: 5px;
      font-size: 13.5px;
      font-weight: 700;
      color: #111827;
    }
    .bank-summary-row.total-row .font-amount {
      font-size: 15px;
      font-weight: 800;
      color: #E67E22;
    }
  </style>
</head>
<body>

  <!-- Screen Control Toolbar -->
  <div class="no-print-bar">
    <div class="bar-info">
      <div class="bar-title">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M6 9V2h12v7"></path>
          <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
          <path d="M6 14h12v8H6z"></path>
        </svg>
        Vista de Impresión Plastiluz
      </div>
      <span class="a4-badge">2 Hojas A4 (Resumen de Ventas + Banco)</span>
    </div>
    <div class="print-actions">
      <button class="btn-action btn-download" id="btn-download-pdf" onclick="downloadDirectPdf()">
        💾 Descargar PDF (2 Hojas A4)
      </button>
      <button class="btn-action btn-print" onclick="window.print()">
        🖨️ Imprimir en A4
      </button>
      <button class="btn-action btn-close" onclick="window.close()">
        ✖ Cerrar
      </button>
    </div>
  </div>

  <div id="pdf-wrapper">

    <!-- ============================================================
         HOJA 1 DE 2: TODO EL RESUMEN DE VENTAS EN 1 SOLA HOJA A4
         ============================================================ -->
    <div class="a4-page" id="page-1">
      
      <!-- Simplified Compact Header: Only Date & Brand -->
      <div class="simple-header">
        <div class="header-left">
          <span class="header-brand">PLASTILUZ</span>
          <span class="header-divider">|</span>
          <span class="header-label">VENTA DEL DÍA:</span>
          <span class="header-date">${formattedDateUpper}</span>
        </div>
        <div class="header-right">
          <span class="header-sheet-tag">HOJA 1: RESUMEN DE VENTAS</span>
        </div>
      </div>

      <!-- 1. Comprobantes Emitidos -->
      <div class="section-title">1. Resumen de Comprobantes Emitidos</div>
      <table>
        <colgroup>
          <col style="width: 18%;">
          <col style="width: 14%;">
          <col class="spacer-col">
          <col style="width: 18%;">
          <col style="width: 14%;">
          <col class="spacer-col">
          <col style="width: 19%;">
          <col style="width: 15%;">
        </colgroup>
        <thead>
          <tr>
            <th colspan="2" class="text-center">VENTAS S/DOC (NOTAS)</th>
            <th class="spacer-col"></th>
            <th colspan="2" class="text-center">BOLETAS DE VENTA</th>
            <th class="spacer-col"></th>
            <th colspan="2" class="text-center">FACTURAS</th>
          </tr>
          <tr>
            <th>Nro.</th>
            <th class="text-right">Importe</th>
            <th class="spacer-col"></th>
            <th>Nro.</th>
            <th class="text-right">Importe</th>
            <th class="spacer-col"></th>
            <th>Nro.</th>
            <th class="text-right">Importe</th>
          </tr>
        </thead>
        <tbody>
          ${comprobantesRowsHtml}
          <tr class="totals-row">
            <td class="text-center font-bold">TOTAL NV</td>
            <td class="text-right font-amount">${fCur(sumNvs)}</td>
            <td class="spacer-col"></td>
            <td class="text-center font-bold">TOTAL BOL</td>
            <td class="text-right font-amount">${fCur(sumBoletas)}</td>
            <td class="spacer-col"></td>
            <td class="text-center font-bold">TOTAL FAC</td>
            <td class="text-right font-amount">${fCur(sumFacturas)}</td>
          </tr>
        </tbody>
      </table>

      <div style="font-size: 9.5px; color: #6B7280; margin-top: -3px; margin-bottom: 6px;">
        * Leyenda: <strong>(*)</strong> Pagado con transferencia/Yape BCP | <strong>NC</strong> Nota de crédito descontada | <strong>CRÉDITO</strong> Venta a crédito | <strong>ANULADO</strong> Comprobante anulado (S/. 0.00)
      </div>

      <!-- Bottom Layout of Page 1: Side-by-side to guarantee 1 single sheet -->
      <div class="page1-bottom-grid">
        
        <!-- Left Column: Créditos & Ajustes -->
        <div>
          ${creditDocs.length > 0 ? `
            <div class="section-title" style="margin-top: 0;">2. Ventas a Crédito del Día</div>
            <table>
              <colgroup>
                <col style="width: 58%;">
                <col style="width: 42%;">
              </colgroup>
              <thead>
                <tr>
                  <th class="text-center">Comprobante</th>
                  <th class="text-right">Importe</th>
                </tr>
              </thead>
              <tbody>
                ${creditRowsHtml}
                <tr class="totals-row">
                  <td class="text-center font-bold">TOTAL CRÉDITO:</td>
                  <td class="text-right font-amount text-amber">${fCur(creditTotal)}</td>
                </tr>
              </tbody>
            </table>
          ` : ''}

          ${cashAdjustments.length > 0 ? `
            <div class="section-title" style="margin-top: 4px;">${creditDocs.length > 0 ? '3' : '2'}. Otros Conceptos (Ingresos / Gastos)</div>
            <table>
              <colgroup>
                <col style="width: 60%;">
                <col style="width: 40%;">
              </colgroup>
              <thead>
                <tr>
                  <th>Concepto / Motivo</th>
                  <th class="text-right">Importe</th>
                </tr>
              </thead>
              <tbody>
                ${adjRowsHtml}
                <tr class="totals-row">
                  <td class="font-bold">NETO AJUSTES:</td>
                  <td class="text-right font-amount ${adjTotal < 0 ? 'text-red' : 'text-blue'}">${fCur(adjTotal)}</td>
                </tr>
              </tbody>
            </table>
          ` : ''}

          ${creditDocs.length === 0 && cashAdjustments.length === 0 ? `
            <div class="section-title" style="margin-top: 0;">2. Resumen Comercial</div>
            <div style="font-size: 11px; color: #6B7280; padding: 10px; background-color: #FAFAFA; border: 1px solid #E5E7EB; border-radius: 4px;">
              • No se registraron ventas a crédito en esta fecha.<br>
              • No se registraron conceptos ni gastos adicionales de caja.
            </div>
          ` : ''}
        </div>

        <!-- Right Column: Resumen Final de Conciliación de Caja -->
        <div>
          <div class="section-title" style="margin-top: 0;">Resumen de Cierre de Caja</div>
          <div class="recon-box">
            <div class="recon-header">
              <span>CONCILIACIÓN DIARIA</span>
              <span style="color: #6B7280; font-weight: normal; font-size: 10.5px;">Control de Ingresos</span>
            </div>
            <div class="recon-row">
              <span>TOTAL VENTAS DEL DÍA (A):</span>
              <span class="font-amount font-bold">${fCur(salesTotal)}</span>
            </div>
            ${adjTotal !== 0 ? `
              <div class="recon-row">
                <span>(+/-) Ajustes / Conceptos Caja:</span>
                <span class="font-amount ${adjTotal < 0 ? 'text-red' : 'text-blue'}">${fCur(adjTotal)}</span>
              </div>
            ` : ''}
            <div class="recon-row" style="background-color: #F9FAFB; padding: 2.5px 4px;">
              <span>INGRESO DEL DÍA (Ventas Netas):</span>
              <span class="font-amount font-medium">${fCur(salesTotal + adjTotal)}</span>
            </div>
            <div class="recon-row">
              <span>(-) Total Banco BCP / Yape (B):</span>
              <span class="font-amount text-blue">${fCur(bankTotal)}</span>
            </div>
            <div class="recon-row">
              <span>(-) Total Ventas a Crédito (C):</span>
              <span class="font-amount text-amber">${fCur(creditTotal)}</span>
            </div>
            <div class="recon-row highlight">
              <span>EFECTIVO TEÓRICO ESPERADO:</span>
              <span class="font-amount font-bold">${fCur(expectedCash)}</span>
            </div>
            <div class="recon-row">
              <span>Total Efectivo Físico Entregado (D):</span>
              <span class="font-amount">${fCur(physicalCashTotal)}</span>
            </div>
            <div class="recon-row final">
              <span>TOTAL RECAUDADO (B + D):</span>
              <span class="font-amount">${fCur(bankTotal + physicalCashTotal)}</span>
            </div>
          </div>

          <div class="diff-status-box ${diffBadgeClass}">
            <div>${diffTitle}</div>
            <div class="diff-sub">${diffDesc}</div>
          </div>
        </div>

      </div>

    </div>


    <!-- ============================================================
         HOJA 2 DE 2: TRANSFERENCIAS Y DINERO RECAUDADO
         ============================================================ -->
    <div class="a4-page" id="page-2">
      
      <!-- Simplified Header: Only Date & Brand -->
      <div class="simple-header">
        <div class="header-left">
          <span class="header-brand">PLASTILUZ</span>
          <span class="header-divider">|</span>
          <span class="header-label">VENTA DEL DÍA:</span>
          <span class="header-date">${formattedDateUpper}</span>
        </div>
        <div class="header-right">
          <span class="header-sheet-tag">HOJA 2: TRANSFERENCIAS Y BANCO</span>
        </div>
      </div>

      <!-- 1. Transferencias Bancarias con formato 12 horas -->
      <div class="section-title">1. Detalle de Transferencias Bancarias (BCP / Yape)</div>
      <table>
        <colgroup>
          <col style="width: 14%;">
          <col style="width: 11%;">
          <col style="width: 18%;">
          <col style="width: 20%;">
          <col style="width: 12%;">
          <col style="width: 15%;">
          <col style="width: 10%;">
        </colgroup>
        <thead>
          <tr>
            <th class="text-center">Fecha</th>
            <th class="text-center">Tipo</th>
            <th class="text-center">Nro. Doc</th>
            <th class="text-center"># Operación</th>
            <th class="text-center">Hora (12h)</th>
            <th class="text-right">Monto</th>
            <th class="text-center">Verif.</th>
          </tr>
        </thead>
        <tbody>
          ${bankRowsHtml}
          <tr class="totals-row">
            <td colspan="5" class="text-right font-bold">TOTAL YAPE Y TRANSFERENCIA:</td>
            <td class="text-right font-amount text-blue">${fCur(bankTotal)}</td>
            <td></td>
          </tr>
        </tbody>
      </table>

      <!-- 2. Arqueo de Dinero Físico (Monedas y Billetes) -->
      <div class="section-title">2. Arqueo de Dinero Físico Recaudado</div>
      <table>
        <colgroup>
          <col style="width: 30%;">
          <col style="width: 26%;">
          <col style="width: 44%;">
          <col class="spacer-col">
          <col style="width: 30%;">
          <col style="width: 26%;">
          <col style="width: 44%;">
        </colgroup>
        <thead>
          <tr>
            <th colspan="3" class="text-center">MONEDAS</th>
            <th class="spacer-col"></th>
            <th colspan="3" class="text-center">BILLETES</th>
          </tr>
          <tr>
            <th class="text-center">Denominación</th>
            <th class="text-center">Cantidad</th>
            <th class="text-right">Total</th>
            <th class="spacer-col"></th>
            <th class="text-center">Denominación</th>
            <th class="text-center">Cantidad</th>
            <th class="text-right">Total</th>
          </tr>
        </thead>
        <tbody>
          ${cashInventoryRowsHtml}
          <tr class="totals-row">
            <td colspan="2" class="text-center font-bold">Total Monedas</td>
            <td class="text-right font-amount">${fCur(coinsSum)}</td>
            <td class="spacer-col"></td>
            <td colspan="2" class="text-center font-bold">Total Billetes</td>
            <td class="text-right font-amount">${fCur(billsSum)}</td>
          </tr>
          <tr style="background-color: #1F2937; color: #FFF;">
            <td colspan="7" class="text-center" style="color:#FFF; padding: 6px 0; font-size: 13.5px; font-weight: 700;">
              TOTAL EFECTIVO FÍSICO ENTREGADO: &nbsp;&nbsp;&nbsp; <strong>${fCur(physicalCashTotal)}</strong>
            </td>
          </tr>
        </tbody>
      </table>

      <!-- 3. Resumen Global de Recaudación -->
      <div class="bank-summary-card">
        <div class="bank-summary-title">Resumen Consolidado de Recaudación</div>
        <div class="bank-summary-row">
          <span>Total Transferencias BCP / Yape (Banco):</span>
          <span class="font-amount text-blue">${fCur(bankTotal)}</span>
        </div>
        <div class="bank-summary-row">
          <span>Total Efectivo Físico Entregado:</span>
          <span class="font-amount">${fCur(physicalCashTotal)}</span>
        </div>
        <div class="bank-summary-row total-row">
          <span>TOTAL GENERAL RECAUDADO (BANCO + EFECTIVO):</span>
          <span class="font-amount">${fCur(bankTotal + physicalCashTotal)}</span>
        </div>
      </div>

    </div>

  </div>

  <script>
    function downloadDirectPdf() {
      const btn = document.getElementById('btn-download-pdf');
      const originalText = btn.innerHTML;
      btn.innerHTML = '⏳ Generando PDF A4 (2 Hojas)...';
      btn.disabled = true;

      const element = document.getElementById('pdf-wrapper');
      const filename = 'Plastiluz_Caja_${formattedDateFilename}.pdf';

      if (typeof html2pdf !== 'undefined') {
        const opt = {
          margin: [0, 0, 0, 0],
          filename: filename,
          image: { type: 'jpeg', quality: 0.98 },
          html2canvas: { scale: 2, useCORS: true, logging: false },
          jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
          pagebreak: { mode: ['css', 'legacy'] }
        };
        html2pdf().set(opt).from(element).save().then(() => {
          btn.innerHTML = originalText;
          btn.disabled = false;
        }).catch(err => {
          console.error('html2pdf error:', err);
          window.print();
          btn.innerHTML = originalText;
          btn.disabled = false;
        });
      } else {
        window.print();
        btn.innerHTML = originalText;
        btn.disabled = false;
      }
    }

    // Auto-launch print dialog
    window.addEventListener('load', () => {
      setTimeout(() => {
        window.print();
      }, 400);
    });
  </script>
</body>
</html>`;

  const printWindow = window.open('', '_blank');
  if (printWindow) {
    printWindow.document.open();
    printWindow.document.write(htmlContent);
    printWindow.document.close();
  } else {
    alert('Por favor habilita las ventanas emergentes en tu navegador para ver e imprimir el reporte PDF.');
  }
}
