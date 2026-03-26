import { ICsvData } from '../models';

/**
 * Parses a CSV string into headers and rows.
 * Handles quoted fields (including fields containing commas and newlines).
 */
export function parseCsv(csvText: string, delimiter: string = ';'): ICsvData {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField: string = '';
  let inQuotes: boolean = false;
  let i: number = 0;

  while (i < csvText.length) {
    const char: string = csvText[i];
    const nextChar: string = i + 1 < csvText.length ? csvText[i + 1] : '';

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Escaped quote
          currentField += '"';
          i += 2;
          continue;
        } else {
          // End of quoted field
          inQuotes = false;
          i++;
          continue;
        }
      } else {
        currentField += char;
        i++;
        continue;
      }
    }

    // Not in quotes
    if (char === '"') {
      inQuotes = true;
      i++;
      continue;
    }

    if (char === delimiter) {
      currentRow.push(currentField.trim());
      currentField = '';
      i++;
      continue;
    }

    if (char === '\r' && nextChar === '\n') {
      currentRow.push(currentField.trim());
      currentField = '';
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }
      currentRow = [];
      i += 2;
      continue;
    }

    if (char === '\n' || char === '\r') {
      currentRow.push(currentField.trim());
      currentField = '';
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }
      currentRow = [];
      i++;
      continue;
    }

    currentField += char;
    i++;
  }

  // Handle last field/row
  if (currentField.length > 0 || currentRow.length > 0) {
    currentRow.push(currentField.trim());
    rows.push(currentRow);
  }

  if (rows.length === 0) {
    return { headers: [], rows: [] };
  }

  const headers: string[] = rows[0];
  const dataRows: string[][] = rows.slice(1).filter((row: string[]) => {
    // Filter out completely empty rows
    return row.some((cell: string) => cell.length > 0);
  });

  return { headers, rows: dataRows };
}

/**
 * Auto-detect the CSV delimiter by examining the first line.
 */
export function detectDelimiter(csvText: string): string {
  const firstLine: string = csvText.split(/\r?\n/)[0] || '';
  const semicolonCount: number = (firstLine.match(/;/g) || []).length;
  const commaCount: number = (firstLine.match(/,/g) || []).length;
  const tabCount: number = (firstLine.match(/\t/g) || []).length;

  if (tabCount > semicolonCount && tabCount > commaCount) {
    return '\t';
  }
  if (semicolonCount >= commaCount) {
    return ';';
  }
  return ',';
}
