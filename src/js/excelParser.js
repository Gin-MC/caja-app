import * as XLSX from 'xlsx';

/**
 * Parses the daily sales Excel file (typically .xls or .xlsx).
 * @param {ArrayBuffer} arrayBuffer 
 * @returns {Object} { date, documents }
 */
export function parseDailyReport(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Convert worksheet to JSON rows
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: null });
  
  if (rawRows.length === 0) {
    throw new Error("El archivo está vacío o no tiene un formato válido.");
  }
  
  const documents = [];
  let parsedDate = null;

  // Process rows
  rawRows.forEach((row, index) => {
    // Standard headers mapping:
    // Comprobante, Tipo, Serie, Numero, Fecha, Codigo, Nombre, Importe
    const tipo = row['Tipo'] ? String(row['Tipo']).trim() : null;
    const serie = row['Serie'] ? String(row['Serie']).trim() : null;
    const numero = row['Numero'] ? String(row['Numero']).trim() : null;
    let importe = parseFloat(row['Importe']);
    
    if (!tipo || !numero) return; // skip invalid rows

    if (isNaN(importe)) {
      importe = 0;
    }

    // Try to parse fecha
    let fecha = row['Fecha'];
    if (fecha && !parsedDate) {
      if (typeof fecha === 'number') {
        parsedDate = excelDateToJSDate(fecha);
      } else {
        parsedDate = new Date(fecha);
      }
    }

    // Determine type label
    let docType = 'NV'; // Default
    if (tipo === '01') docType = 'FACTURA';
    else if (tipo === '03') docType = 'BOLETA';
    else if (tipo === 'NV' || tipo.toUpperCase() === 'NV') docType = 'NV';
    else if (tipo === '07') docType = 'NC'; // Nota de crédito

    // Create standard object
    documents.push({
      id: `${tipo}-${serie}-${numero}-${index}`,
      rawTipo: tipo,
      type: docType,
      serie: serie || '',
      numero: numero,
      fecha: typeof fecha === 'number' ? excelDateToJSDate(fecha).toISOString().split('T')[0] : (fecha || ''),
      codigo: row['Codigo'] ? String(row['Codigo']).trim() : '00000000',
      nombre: row['Nombre'] ? String(row['Nombre']).trim() : 'CLIENTES VARIOS',
      importe: importe,
      metodoPago: 'efectivo', // default
      verificado: 'NO', // for bank transfers
      nroOperacion: '',
      hora: ''
    });
  });

  // Sort documents by number ascending
  documents.sort((a, b) => {
    const numA = parseInt(a.numero.replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(b.numero.replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  });

  return {
    date: parsedDate || new Date(),
    documents: documents
  };
}

/**
 * Converts Excel serial date to JS Date object
 * @param {number} serial 
 * @returns {Date}
 */
export function excelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  
  const hours = Math.floor(total_seconds / 3600);
  total_seconds %= 3600;
  const minutes = Math.floor(total_seconds / 60);
  const seconds = total_seconds % 60;
  
  // Adjust for timezone offset
  const localDate = new Date(
    date_info.getUTCFullYear(),
    date_info.getUTCMonth(),
    date_info.getUTCDate(),
    hours,
    minutes,
    seconds
  );
  
  return localDate;
}
