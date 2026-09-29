/**
 * Helper to parse dates into comparable timestamps
 * Handles formats: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, or standard Date strings
 * @param {string|number|Date} str 
 * @returns {number}
 */
export function parseDateKey(str) {
  if (!str) return 0;
  if (typeof str === 'number') return str;
  if (str instanceof Date) return str.getTime();

  const clean = String(str).trim();
  const dmy = clean.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    const day = dmy[1].padStart(2, '0');
    const month = dmy[2].padStart(2, '0');
    const year = dmy[3];
    return new Date(`${year}-${month}-${day}T00:00:00Z`).getTime() || 0;
  }
  const ymd = clean.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymd) {
    const year = ymd[1];
    const month = ymd[2].padStart(2, '0');
    const day = ymd[3].padStart(2, '0');
    return new Date(`${year}-${month}-${day}T00:00:00Z`).getTime() || 0;
  }
  const ts = Date.parse(clean);
  return isNaN(ts) ? 0 : ts;
}

/**
 * Helper to parse time strings into seconds from midnight
 * Handles formats: HH:MM, HH:MM:SS, HH.MM, with optional AM/PM
 * @param {string} str 
 * @returns {number|null}
 */
export function parseTimeSeconds(str) {
  if (!str || typeof str !== 'string') return null;
  const clean = str.trim();
  if (clean === '') return null;
  
  const match = clean.match(/^(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*(am|pm)?/i);
  if (match) {
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const seconds = match[3] ? parseInt(match[3], 10) : 0;
    const modifier = match[4] ? match[4].toLowerCase() : null;
    if (modifier === 'pm' && hours < 12) hours += 12;
    if (modifier === 'am' && hours === 12) hours = 0;
    return hours * 3600 + minutes * 60 + seconds;
  }
  return null;
}

/**
 * Sorts bank transfers: Date ascending, then Time ascending
 * @param {Array} transfers 
 * @returns {Array} sorted copy of transfers
 */
export function sortBankTransfers(transfers) {
  if (!Array.isArray(transfers)) return [];
  return [...transfers].sort((a, b) => {
    // 1. Compare Date ascending
    const dateA = parseDateKey(a.fecha);
    const dateB = parseDateKey(b.fecha);
    if (dateA !== dateB) {
      return dateA - dateB;
    }

    // 2. Compare Time ascending
    const timeA = parseTimeSeconds(a.hora);
    const timeB = parseTimeSeconds(b.hora);

    if (timeA !== null && timeB !== null) {
      if (timeA !== timeB) return timeA - timeB;
      return String(a.hora || '').localeCompare(String(b.hora || ''));
    }
    // Items with recorded time come before items without time
    if (timeA !== null && timeB === null) return -1;
    if (timeA === null && timeB !== null) return 1;

    // 3. Fallback: document number ascending
    const numA = parseInt(String(a.numero || '').replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(String(b.numero || '').replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  });
}
