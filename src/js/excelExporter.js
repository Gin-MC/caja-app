import XLSX from 'xlsx-js-style';

/**
 * Exports the session data into a styled two-sheet Excel file matching Plastiluz's original structure.
 * @param {Object} session { date, documents, bankTransfers, cashCount, cashAdjustments, totals }
 */
export function exportToExcel(session) {
  const { date, documents, bankTransfers, cashCount, cashAdjustments = [], totals } = session;

  const dateObj = new Date(date);
  // Excel date serial number calculated in a timezone-independent manner to prevent 1-day offsets
  const year = dateObj.getFullYear();
  const month = dateObj.getMonth();
  const day = dateObj.getDate();
  const excelDateValue = Math.round((Date.UTC(year, month, day) - Date.UTC(1899, 11, 30)) / (24 * 60 * 60 * 1000));

  // Styles definitions
  const borderThin = {
    top: { style: 'thin', color: { rgb: 'D0D0D0' } },
    bottom: { style: 'thin', color: { rgb: 'D0D0D0' } },
    left: { style: 'thin', color: { rgb: 'D0D0D0' } },
    right: { style: 'thin', color: { rgb: 'D0D0D0' } }
  };

  const fillHeader = { fgColor: { rgb: 'E5E8E8' } }; // Soft light grey header
  const fillBCP = { fgColor: { rgb: 'D4E6F1' } }; // Soft blue for banco
  const fillCredit = { fgColor: { rgb: 'FCF3CF' } }; // Soft yellow for credito
  const fillNC = { fgColor: { rgb: 'FADBD8' } }; // Soft red for notas de credito

  // --- SHEET 1: caja ---
  const cajaRows = [];
  
  // Row 1: VENTA DEL DIA
  cajaRows.push(['VENTA DEL DIA', null, null, null, null, excelDateValue]);
  cajaRows.push([]); // empty spacer
  
  // Row 3: Headers
  cajaRows.push([
    'Ventas S/Doc', null, null, null, 
    'BOLETAS', null, null, null, 
    'FACTURA'
  ]);

  // Separate lists (NCs go in Boletas/Facturas columns directly)
  const nvs = documents.filter(d => d.type === 'NV');
  const boletas = documents.filter(d => d.type === 'BOLETA' || (d.type === 'NC' && d.serie.toUpperCase().startsWith('B')));
  const facturas = documents.filter(d => d.type === 'FACTURA' || (d.type === 'NC' && d.serie.toUpperCase().startsWith('F')));

  // Sorter: normal documents sorted ascending by document number, Credit Notes (NC) go at the end
  const docSorter = (a, b) => {
    if (a.type === 'NC' && b.type !== 'NC') return 1;
    if (a.type !== 'NC' && b.type === 'NC') return -1;
    const numA = parseInt(a.numero.replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(b.numero.replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  };

  nvs.sort(docSorter);
  boletas.sort(docSorter);
  facturas.sort(docSorter);

  const maxLen = Math.max(nvs.length, boletas.length, facturas.length);

  let sumNvs = 0;
  let sumBoletas = 0;
  let sumFacturas = 0;

  // Add items
  for (let i = 0; i < maxLen; i++) {
    const nv = nvs[i] || null;
    const bol = boletas[i] || null;
    const fac = facturas[i] || null;

    const row = [];
    
    // Ventas S/Doc (Col A: indicator, Col B: number, Col C: amount)
    if (nv) {
      row[0] = nv.metodoPago === 'banco' ? '(*)' : (nv.metodoPago === 'credito' ? nv.importe : null); // Col A
      row[1] = nv.numero; // Col B
      row[2] = nv.metodoPago === 'credito' ? 'credito' : nv.importe; // Col C
      sumNvs += nv.importe;
    } else {
      row[0] = null;
      row[1] = null;
      row[2] = null;
    }
    
    row[3] = null; // Col D divider spacer

    // Boletas (Col E: indicator, Col F: number, Col G: amount)
    if (bol) {
      row[4] = bol.type === 'NC' ? 'NC' : (bol.metodoPago === 'banco' ? '(*)' : (bol.metodoPago === 'credito' ? (bol.type === 'NC' ? -bol.importe : bol.importe) : null)); // Col E
      row[5] = bol.type === 'NC' ? `${bol.serie}-${bol.numero}` : bol.numero; // Col F
      const val = bol.type === 'NC' ? -bol.importe : bol.importe;
      row[6] = bol.metodoPago === 'credito' ? 'credito' : val; // Col G
      sumBoletas += val;
    } else {
      row[4] = null;
      row[5] = null;
      row[6] = null;
    }
    
    row[7] = null; // Col H divider spacer

    // Facturas (Col I: indicator, Col J: number, Col K: amount)
    if (fac) {
      row[8] = fac.type === 'NC' ? 'NC' : (fac.metodoPago === 'banco' ? '(*)' : (fac.metodoPago === 'credito' ? (fac.type === 'NC' ? -fac.importe : fac.importe) : null)); // Col I
      row[9] = fac.type === 'NC' ? `${fac.serie}-${fac.numero}` : fac.numero; // Col J
      const val = fac.type === 'NC' ? -fac.importe : fac.importe;
      row[10] = fac.metodoPago === 'credito' ? 'credito' : val; // Col K
      sumFacturas += val;
    } else {
      row[8] = null;
      row[9] = null;
      row[10] = null;
    }

    cajaRows.push(row);
  }

  // Row: TOTALS (aligned perfectly under the new columns structure!)
  cajaRows.push([]);
  cajaRows.push([
    null, 'TOTAL', sumNvs, null, // Col B: 'TOTAL', Col C: sum
    null, 'TOTAL', sumBoletas, null, // Col F: 'TOTAL', Col G: sum
    null, 'TOTAL', sumFacturas // Col J: 'TOTAL', Col K: sum
  ]);
  
  cajaRows.push([]);
  
  // Small Ventas a Crédito Table pushed to Column I & J below document totals
  const creditDocs = documents.filter(d => d.metodoPago === 'credito');
  
  const rowCrdTitle = [];
  rowCrdTitle[8] = 'VENTAS A CRÉDITO';
  cajaRows.push(rowCrdTitle);

  const rowCrdHeader = [];
  rowCrdHeader[8] = 'Nro. Doc';
  rowCrdHeader[9] = 'Monto';
  cajaRows.push(rowCrdHeader);

  if (creditDocs.length === 0) {
    const rowEmpty = [];
    rowEmpty[8] = 'Sin créditos';
    rowEmpty[9] = 0;
    cajaRows.push(rowEmpty);
  } else {
    creditDocs.forEach(d => {
      const rowDoc = [];
      rowDoc[8] = d.type === 'NC' ? `${d.serie}-${d.numero}` : `${d.serie || ''}-${d.numero}`;
      rowDoc[9] = d.type === 'NC' ? -d.importe : d.importe;
      cajaRows.push(rowDoc);
    });
  }
  
  cajaRows.push([]); // spacer row
  
  const salesTotal = sumNvs + sumBoletas + sumFacturas;

  // --- RECONCILIATION SUMMARY BLOCK IN EXACT SPECIFIED ORDER ---
  
  // Row 1: TOTAL VENTAS DEL DIA (sum of all vouchers, i.e., sumNvs + sumBoletas + sumFacturas)
  const rowTotVentas = [];
  rowTotVentas[2] = 'TOTAL VENTAS DEL DIA';
  rowTotVentas[6] = salesTotal;
  cajaRows.push(rowTotVentas);
  cajaRows.push([]); // spacer

  // Rows: Custom adjustments added manually (income / expense concepts)
  let adjTotal = 0;
  if (cashAdjustments.length > 0) {
    cashAdjustments.forEach(adj => {
      if (adj.description.trim() === '' && adj.amount === 0) return;
      const rowAdj = [];
      rowAdj[2] = adj.description.toUpperCase() || 'AJUSTE CAJA';
      rowAdj[6] = adj.amount;
      cajaRows.push(rowAdj);
      adjTotal += adj.amount;
    });
    cajaRows.push([]); // spacer
  }

  // Row 3: INGRESO DEL DIA (Net income = TOTAL VENTAS DEL DIA + sum of Adjustments)
  const rowIngreso = [];
  rowIngreso[2] = 'INGRESO DEL DIA';
  rowIngreso[6] = salesTotal + adjTotal;
  cajaRows.push(rowIngreso);
  cajaRows.push([]); // spacer

  // Row 4: TOTAL BANCO BCP
  const rowTotBanco = [];
  rowTotBanco[2] = 'TOTAL BANCO BCP';
  rowTotBanco[6] = totals.bankTotal;
  cajaRows.push(rowTotBanco);

  // Row 5: TOTAL DINERO ENTREGADO (Physical Cash Count)
  const rowTotEntregado = [];
  rowTotEntregado[2] = 'TOTAL DINERO ENTREGADO';
  rowTotEntregado[6] = totals.cashTotal;
  cajaRows.push(rowTotEntregado);
  cajaRows.push([]); // spacer

  // Row 6: TOTAL DINERO (BCP + Physical Cash)
  const rowTotDinero = [];
  rowTotDinero[2] = 'TOTAL DINERO';
  rowTotDinero[6] = totals.bankTotal + totals.cashTotal;
  cajaRows.push(rowTotDinero);
  cajaRows.push([]); // spacer

  // Row 7: sobra/falta
  const rowSobraFalta = [];
  rowSobraFalta[2] = 'sobra/falta';
  rowSobraFalta[6] = totals.difference;
  cajaRows.push(rowSobraFalta);

  // Convert array to worksheet
  const cajaSheet = XLSX.utils.aoa_to_sheet(cajaRows);

  // Apply Styling for 'caja' Sheet
  cajaSheet['!views'] = [{ showGridLines: true }];
  
  // Format Date in Row 1 (cell F1 is column index 5, row index 0)
  const cellF1 = cajaSheet['F1'];
  if (cellF1) {
    cellF1.z = 'yyyy-mm-dd';
    cellF1.s = {
      font: { name: 'Arial', sz: 11, bold: true, color: { rgb: '2B2521' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
  }

  // Row 1 VENTA DEL DIA label
  const cellA1 = cajaSheet['A1'];
  if (cellA1) {
    cellA1.s = {
      font: { name: 'Arial', sz: 12, bold: true, color: { rgb: 'E67E22' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    };
  }

  // Row 3 Headers
  const headerCols = ['A', 'E', 'I'];
  headerCols.forEach(col => {
    const cell = cajaSheet[`${col}3`];
    if (cell) {
      cell.s = {
        font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '2B2521' } },
        fill: fillHeader,
        alignment: { horizontal: 'center', vertical: 'center' },
        border: borderThin
      };
    }
  });

  // Apply borders, alignments and styles for all grid cells
  Object.keys(cajaSheet).forEach(key => {
    if (key.startsWith('!')) return;
    const cell = cajaSheet[key];
    const col = key.replace(/[0-9]/g, '');
    const row = parseInt(key.replace(/\D/g, ''), 10);

    // Apply default font (Strictly integer sizes to prevent Excel warning!)
    if (!cell.s) {
      cell.s = { font: { name: 'Arial', sz: 10 }, alignment: { vertical: 'center' } };
    }

    // Grid data styling (Row 4 to Row 4 + maxLen)
    if (row >= 4 && row < 4 + maxLen) {
      cell.s.border = borderThin;

      // Col A, E, I: Method indicators (asterisk (*)/indicators aligned to left)
      if (col === 'A' || col === 'E' || col === 'I') {
        cell.s.alignment = { horizontal: 'left', vertical: 'center' };
        cell.s.font = { name: 'Arial', sz: 9, bold: true };
        
        if (cell.v === '(*)') {
          cell.s.fill = fillBCP;
          cell.s.font.color = { rgb: '1B4F72' };
        } else if (typeof cell.v === 'number') {
          // Numerical credit amount goes to indicator column in template!
          cell.s.alignment = { horizontal: 'right', vertical: 'center' };
          cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '7E5109' } };
          cell.s.fill = fillCredit;
          cell.z = 'S/. #,##0.00';
        } else if (cell.v === 'NC') {
          cell.s.fill = fillNC;
          cell.s.font.color = { rgb: '78281F' };
        }
      }

      // Col B, F, J: Document numbers
      if (col === 'B' || col === 'F' || col === 'J') {
        cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        cell.s.font = { name: 'Arial', sz: 10, bold: true };
        
        // Highlight NC numbers
        if (cell.v && (String(cell.v).startsWith('BC01') || String(cell.v).startsWith('FC01'))) {
          cell.s.fill = fillNC;
          cell.s.font.color = { rgb: '900C3F' };
        }
      }

      // Col C, G, K: Importes
      if (col === 'C' || col === 'G' || col === 'K') {
        if (cell.v === 'credito') {
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
          cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '7E5109' } };
          cell.s.fill = fillCredit;
        } else {
          cell.s.alignment = { horizontal: 'right', vertical: 'center' };
          cell.s.font = { name: 'Arial', sz: 10, bold: true };
          cell.z = 'S/. #,##0.00';
          if (cell.v < 0) {
            cell.s.font.color = { rgb: 'C0392B' };
          }
        }
      }
    }

    // Total Row Styling (Row = 4 + maxLen + 1)
    if (row === 4 + maxLen + 1) {
      cell.s.border = borderThin;
      cell.s.font = { name: 'Arial', sz: 10, bold: true };
      cell.s.fill = fillHeader;
      
      if (col === 'B' || col === 'F' || col === 'J') {
        cell.s.alignment = { horizontal: 'center', vertical: 'center' };
      }
      if (col === 'C' || col === 'G' || col === 'K') {
        cell.s.alignment = { horizontal: 'right', vertical: 'center' };
        cell.z = 'S/. #,##0.00';
      }
    }

    // Ventas a Crédito small table styling (Column I and J below documents total)
    const startCreditRow = 4 + maxLen + 3;
    const endCreditRow = startCreditRow + 1 + Math.max(1, creditDocs.length);
    
    if (row >= startCreditRow && row <= endCreditRow) {
      if (col === 'I' || col === 'J') {
        cell.s.border = borderThin;
        if (row === startCreditRow) {
          // Title
          cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '7E5109' } };
          cell.s.fill = fillCredit;
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        } else if (row === startCreditRow + 1) {
          // Header
          cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '2B2521' } };
          cell.s.fill = fillHeader;
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        } else {
          // Data
          if (col === 'I') {
            cell.s.alignment = { horizontal: 'center', vertical: 'center' };
            cell.s.font = { name: 'Arial', sz: 10 };
          } else if (col === 'J') {
            cell.s.alignment = { horizontal: 'right', vertical: 'center' };
            cell.s.font = { name: 'Arial', sz: 10, bold: true };
            cell.z = 'S/. #,##0.00';
            if (cell.v < 0) {
              cell.s.font.color = { rgb: 'C0392B' };
            }
          }
        }
      }
    }

    // Daily reconciliation summary block styling
    if (row >= 4 + maxLen + 3) {
      if (col === 'C') {
        cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '2B2521' } };
        cell.s.alignment = { horizontal: 'left', vertical: 'center' };
        
        if (cell.v === 'TOTAL VENTAS DEL DIA' || cell.v === 'INGRESO DEL DIA' || cell.v === 'TOTAL DINERO' || cell.v === 'sobra/falta') {
          cell.s.font.sz = 11;
          cell.s.font.color = { rgb: 'E67E22' };
        }
      }

      if (col === 'G') {
        cell.s.font = { name: 'Arial', sz: 11, bold: true };
        cell.s.alignment = { horizontal: 'right', vertical: 'center' };
        cell.z = 'S/. #,##0.00';
        
        // Colormap difference field
        if (cajaSheet[`C${row}`]?.v === 'sobra/falta') {
          if (cell.v < -0.015) {
            cell.s.font.color = { rgb: 'C0392B' }; // Red
            cell.s.fill = fillNC;
          } else if (cell.v > 0.015) {
            cell.s.font.color = { rgb: '2980B9' }; // Blue
            cell.s.fill = fillBCP;
          } else {
            cell.s.font.color = { rgb: '27AE60' }; // Green
            cell.s.fill = { fgColor: { rgb: 'EAFAF1' } };
          }
        }
      }
    }
  });


  // --- SHEET 2: banco ---
  const bancoRows = [];

  // Row 1: VENTA DEL DIA
  bancoRows.push(['VENTA DEL DIA', null, null, null, null, excelDateValue]);
  bancoRows.push([]);

  // Row 3: Headers
  bancoRows.push(['FECHA', 'HORA', '#OPERACIÓN ', 'TIPO', 'NRO.DOC', 'MONTO', 'VERIFICADO']);

  let bankSum = 0;
  bankTransfers.forEach(t => {
    const actualAmt = t.type === 'NC' ? -t.importe : t.importe;
    
    // Parse custom date strings or keep them
    bancoRows.push([
      t.fecha,
      t.hora || null,
      t.nroOperacion || '',
      t.type,
      `${t.serie}-${t.numero}`,
      actualAmt,
      t.verificado || 'SI'
    ]);
    bankSum += actualAmt;
  });

  // Total Row
  bancoRows.push(['TOTAL YAPE Y TRASFERENCIA', null, null, null, null, bankSum]);
  bancoRows.push([]);

  // Physical cash count block
  bancoRows.push(['MONEDAS', null, null, null, 'BILLETES']);
  bancoRows.push(['VALOR', 'CANTIDAD', 'TOTAL', null, 'VALOR', 'CANTIDAD', 'TOTAL']);

  const coins = cashCount.filter(c => c.value < 10);
  const bills = cashCount.filter(c => c.value >= 10);

  const cashMaxLen = Math.max(coins.length, bills.length);

  let coinsSum = 0;
  let billsSum = 0;

  for (let i = 0; i < cashMaxLen; i++) {
    const coin = coins[i] || null;
    const bill = bills[i] || null;
    
    const row = [];
    
    if (coin) {
      row[0] = coin.value;
      row[1] = coin.quantity || null;
      row[2] = coin.total;
      coinsSum += coin.total;
    } else {
      row[0] = null;
      row[1] = null;
      row[2] = null;
    }

    row[3] = null; // empty column divider

    if (bill) {
      row[4] = bill.value;
      row[5] = bill.quantity || null;
      row[6] = bill.total;
      billsSum += bill.total;
    } else {
      row[4] = null;
      row[5] = null;
      row[6] = null;
    }

    bancoRows.push(row);
  }

  // Add bill and coin totals
  bancoRows.push([
    null, null, null, null, 
    'TOTAL', null, billsSum
  ]);
  bancoRows.push([
    'TOTAL', null, coinsSum
  ]);

  const bancoSheet = XLSX.utils.aoa_to_sheet(bancoRows);

  // Apply bank sheet styles
  bancoSheet['!views'] = [{ showGridLines: true }];

  // Format Date in Row 1 (cell F1 is column index 5)
  const cellBankF1 = bancoSheet['F1'];
  if (cellBankF1) {
    cellBankF1.z = 'yyyy-mm-dd';
    cellBankF1.s = {
      font: { name: 'Arial', sz: 11, bold: true, color: { rgb: '2B2521' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
  }

  // Row 1 VENTA DEL DIA label
  const cellBankA1 = bancoSheet['A1'];
  if (cellBankA1) {
    cellBankA1.s = {
      font: { name: 'Arial', sz: 12, bold: true, color: { rgb: 'E67E22' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    };
  }

  // Grid styling for Banco sheet
  Object.keys(bancoSheet).forEach(key => {
    if (key.startsWith('!')) return;
    const cell = bancoSheet[key];
    const col = key.replace(/[0-9]/g, '');
    const row = parseInt(key.replace(/\D/g, ''), 10);

    // Apply default font (Integer sizes only to prevent Excel XML warning!)
    if (!cell.s) {
      cell.s = { font: { name: 'Arial', sz: 10 }, alignment: { vertical: 'center' } };
    }

    // Row 3 (Headers BCP Transfers)
    if (row === 3) {
      cell.s.font = { name: 'Arial', sz: 10, bold: true, color: { rgb: '2B2521' } };
      cell.s.fill = fillHeader;
      cell.s.border = borderThin;
      cell.s.alignment = { horizontal: col === 'F' ? 'right' : 'center', vertical: 'center' };
    }

    // Bank data rows
    if (row >= 4 && row < 4 + bankTransfers.length) {
      cell.s.border = borderThin;
      
      if (col === 'F') { // Monto
        cell.s.alignment = { horizontal: 'right', vertical: 'center' };
        cell.s.font = { name: 'Arial', sz: 10, bold: true };
        cell.z = 'S/. #,##0.00';
        if (cell.v < 0) {
          cell.s.font.color = { rgb: 'C0392B' };
          cell.s.fill = fillNC;
        }
      } else {
        cell.s.alignment = { horizontal: 'center', vertical: 'center' };
      }

      if (col === 'C') { // # Operacion
        cell.s.font = { name: 'Arial', sz: 10, bold: true };
      }

      if (col === 'G' && cell.v === 'SI') {
        cell.s.fill = { fgColor: { rgb: 'EAFAF1' } };
        cell.s.font = { name: 'Arial', sz: 9, bold: true, color: { rgb: '27AE60' } };
      }
    }

    // Bank total row
    if (row === 4 + bankTransfers.length) {
      cell.s.font = { name: 'Arial', sz: 10, bold: true };
      cell.s.border = borderThin;
      cell.s.fill = fillHeader;
      
      if (col === 'F') {
        cell.s.alignment = { horizontal: 'right', vertical: 'center' };
        cell.z = 'S/. #,##0.00';
      }
    }

    // Cash Count section styling
    const startCashRow = 4 + bankTransfers.length + 2;
    
    // "MONEDAS" and "BILLETES" section titles
    if (row === startCashRow) {
      cell.s.font = { name: 'Arial', sz: 11, bold: true, color: { rgb: 'E67E22' } };
      cell.s.alignment = { horizontal: 'center', vertical: 'center' };
      cell.s.fill = fillHeader;
      cell.s.border = borderThin;
    }

    // Denominations headers
    if (row === startCashRow + 1) {
      cell.s.font = { name: 'Arial', sz: 9, bold: true, color: { rgb: '2B2521' } };
      cell.s.alignment = { horizontal: 'center', vertical: 'center' };
      cell.s.border = borderThin;
    }

    // Denominations inventory grid data (coins and bills)
    if (row >= startCashRow + 2 && row < startCashRow + 2 + cashMaxLen) {
      // Columns A, B, C (Coins) & E, F, G (Bills)
      if (col === 'A' || col === 'B' || col === 'C' || col === 'E' || col === 'F' || col === 'G') {
        cell.s.border = borderThin;
        
        if (col === 'A' || col === 'E') { // Valor denomination
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
          cell.s.font = { name: 'Arial', sz: 10, bold: true };
          cell.z = 'S/. #,##0.00';
        }
        if (col === 'B' || col === 'F') { // Qty
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
        if (col === 'C' || col === 'G') { // Total subtotal
          cell.s.alignment = { horizontal: 'right', vertical: 'center' };
          cell.s.font = { name: 'Arial', sz: 10, bold: true };
          cell.z = 'S/. #,##0.00';
        }
      }
    }

    // Inventory Totals rows
    if (row === startCashRow + 2 + cashMaxLen) {
      if (col === 'E' || col === 'G') {
        cell.s.font = { name: 'Arial', sz: 10, bold: true };
        cell.s.fill = fillHeader;
        cell.s.border = borderThin;
        if (col === 'G') {
          cell.s.alignment = { horizontal: 'right', vertical: 'center' };
          cell.z = 'S/. #,##0.00';
        } else {
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
      }
    }

    if (row === startCashRow + 2 + cashMaxLen + 1) {
      if (col === 'A' || col === 'C') {
        cell.s.font = { name: 'Arial', sz: 10, bold: true };
        cell.s.fill = fillHeader;
        cell.s.border = borderThin;
        if (col === 'C') {
          cell.s.alignment = { horizontal: 'right', vertical: 'center' };
          cell.z = 'S/. #,##0.00';
        } else {
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
      }
    }
  });

  // Set widths
  cajaSheet['!cols'] = [
    { wch: 15 }, { wch: 15 }, { wch: 14 }, { wch: 4 }, // Ventas S/Doc (A, B, C, D)
    { wch: 15 }, { wch: 15 }, { wch: 14 }, { wch: 4 }, // Boletas (E, F, G, H)
    { wch: 15 }, { wch: 15 }, { wch: 14 }              // Facturas (I, J, K)
  ];

  bancoSheet['!cols'] = [
    { wch: 14 }, { wch: 10 }, { wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 12 }
  ];

  // Create workbook and append sheets
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, cajaSheet, 'caja');
  XLSX.utils.book_append_sheet(wb, bancoSheet, 'banco');

  // Trigger download
  const dateStr = dateObj.toISOString().split('T')[0];
  XLSX.writeFile(wb, `CAJA_PLASTILUZ_${dateStr}.xlsx`);
}
