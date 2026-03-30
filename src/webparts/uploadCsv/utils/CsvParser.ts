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
 * Decode raw CSV file bytes into a string.
 * Strategy:
 *  1. If a UTF-8 BOM (EF BB BF) is present, decode as UTF-8 and strip the BOM.
 *  2. If a UTF-16 LE BOM (FF FE) is present, decode as UTF-16LE and strip the BOM.
 *  3. Scan the raw bytes to check if they are valid UTF-8.
 *  4. If valid UTF-8, decode as UTF-8.
 *  5. Otherwise decode as Windows-1252 (the encoding Excel uses for CSV
 *     on Western-European Windows systems) using a manual byte map,
 *     so we don't depend on TextDecoder supporting that label.
 */
export function decodeCsvBytes(buffer: ArrayBuffer): string {
  const bytes: Uint8Array = new Uint8Array(buffer);

  // Check for UTF-8 BOM
  if (bytes.length >= 3
    && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    console.log('[CsvParser] decodeCsvBytes — UTF-8 BOM detected');
    const decoder: TextDecoder = new TextDecoder('utf-8');
    return decoder.decode(bytes.slice(3));
  }

  // Check for UTF-16 LE BOM
  if (bytes.length >= 2
    && bytes[0] === 0xFF && bytes[1] === 0xFE) {
    console.log('[CsvParser] decodeCsvBytes — UTF-16 LE BOM detected');
    const decoder: TextDecoder = new TextDecoder('utf-16le');
    return decoder.decode(bytes.slice(2));
  }

  // Check if bytes are valid UTF-8 by scanning manually
  if (_isValidUtf8(bytes)) {
    console.log('[CsvParser] decodeCsvBytes — valid UTF-8 (no BOM)');
    const decoder: TextDecoder = new TextDecoder('utf-8');
    return decoder.decode(bytes);
  }

  // Not valid UTF-8 → decode as Windows-1252
  console.log('[CsvParser] decodeCsvBytes — not valid UTF-8,',
    'decoding as Windows-1252');
  return _decodeWindows1252(bytes);
}

/**
 * Check if a byte array is valid UTF-8 by scanning the byte sequences.
 * Returns false as soon as an invalid sequence is found.
 */
/* tslint:disable:no-bitwise */
function _isValidUtf8(bytes: Uint8Array): boolean {
  let i: number = 0;
  while (i < bytes.length) {
    const b: number = bytes[i];
    if (b <= 0x7F) {
      // ASCII — always valid
      i++;
    } else if (b >= 0xC2 && b <= 0xDF) {
      // 2-byte sequence
      if (i + 1 >= bytes.length || (bytes[i + 1] & 0xC0) !== 0x80) {
        return false;
      }
      i += 2;
    } else if (b >= 0xE0 && b <= 0xEF) {
      // 3-byte sequence
      if (i + 2 >= bytes.length
        || (bytes[i + 1] & 0xC0) !== 0x80
        || (bytes[i + 2] & 0xC0) !== 0x80) {
        return false;
      }
      // Reject overlong sequences
      if (b === 0xE0 && bytes[i + 1] < 0xA0) {
        return false;
      }
      i += 3;
    } else if (b >= 0xF0 && b <= 0xF4) {
      // 4-byte sequence
      if (i + 3 >= bytes.length
        || (bytes[i + 1] & 0xC0) !== 0x80
        || (bytes[i + 2] & 0xC0) !== 0x80
        || (bytes[i + 3] & 0xC0) !== 0x80) {
        return false;
      }
      // Reject overlong sequences
      if (b === 0xF0 && bytes[i + 1] < 0x90) {
        return false;
      }
      i += 4;
    } else {
      // 0x80-0xBF (unexpected continuation), 0xC0-0xC1 (overlong),
      // 0xF5-0xFF (invalid) — all indicate non-UTF-8
      return false;
    }
  }
  return true;
}
/* tslint:enable:no-bitwise */

// tslint:disable-next-line:max-line-length
/** Windows-1252 code points for bytes 0x80–0x9F (the range that differs from ISO-8859-1). */
const WIN1252_MAP: number[] = [
  0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, // 80-87
  0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F, // 88-8F
  0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, // 90-97
  0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178  // 98-9F
];

/**
 * Decode bytes as Windows-1252 without relying on TextDecoder support.
 * Bytes 0x00-0x7F map to ASCII, 0x80-0x9F use the WIN1252_MAP,
 * and 0xA0-0xFF map directly to the same Unicode code point.
 */
function _decodeWindows1252(bytes: Uint8Array): string {
  const chars: string[] = [];
  for (let i: number = 0; i < bytes.length; i++) {
    const b: number = bytes[i];
    if (b < 0x80) {
      chars.push(String.fromCharCode(b));
    } else if (b >= 0x80 && b <= 0x9F) {
      chars.push(String.fromCharCode(WIN1252_MAP[b - 0x80]));
    } else {
      // 0xA0-0xFF: same code point in Unicode as in Windows-1252
      chars.push(String.fromCharCode(b));
    }
  }
  return chars.join('');
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
