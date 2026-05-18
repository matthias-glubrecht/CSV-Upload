import { IImportProgress } from '../models';

/**
 * Parse a taxonomy value that may be in SharePoint export format.
 *
 * SharePoint's "Export to Excel" writes taxonomy fields as:
 *   Single:  "<wssId>;#<label>"          e.g. "6;#Einsatz"
 *   Multi:   "<wssId>;#<label>;#<wssId>;#<label>"  e.g. "12;#Lehre;#13;#Forschung"
 *
 * SharePoint sometimes prefixes labels with an extra "#" in the
 * export (e.g. "22;##IT-Verfahren" instead of "22;#IT-Verfahren").
 * After splitting on ";#", this leaves a leading "#" on the label
 * which must be stripped.
 *
 * This method detects the pattern and returns an array of clean labels.
 * If the value is NOT in export format, it splits by ";" as usual.
 */
export function parseTaxonomyExportValue(csvValue: string): string[] {
  // Detect SharePoint export format: starts with digits followed by ;#
  if (/^\d+;#/.test(csvValue)) {
    const segments: string[] = csvValue.split(';#');
    const labels: string[] = [];
    for (let i: number = 0; i < segments.length; i++) {
      let seg: string = segments[i].trim();
      if (!seg) {
        continue;
      }
      // Skip segments that are purely numeric (wssId values)
      if (/^\d+$/.test(seg)) {
        continue;
      }
      // Strip leading "#" left over from SharePoint's "##" export artifact
      if (seg.charAt(0) === '#') {
        seg = seg.substring(1);
      }
      if (seg) {
        labels.push(seg);
      }
    }
    return labels;
  }
  // Plain format: split by semicolon
  return csvValue.split(';').map((s: string) => s.trim()).filter((s: string) => s !== '');
}

/**
 * Resolve special tokens in default values (e.g. [Heute] / [Today] for date fields).
 */
export function resolveDefaultValue(defaultValue: string, fieldType: string): string {
  if (!defaultValue) {
    return '';
  }
  const lower: string = defaultValue.toLowerCase().trim();

  // Date tokens
  if (fieldType === 'DateTime') {
    if (lower === '[heute]' || lower === '[today]') {
      return new Date().toISOString();
    }
    // Handle [Heute]+N / [Today]+N  (offset in days)
    const offsetMatch: RegExpMatchArray = lower.match(
      /^\[(heute|today)\]\s*([+-]\s*\d+)$/
    );
    if (offsetMatch) {
      const offsetDays: number = parseInt(offsetMatch[2].replace(/\s/g, ''), 10);
      const d: Date = new Date();
      d.setDate(d.getDate() + offsetDays);
      return d.toISOString();
    }
  }

  // Boolean tokens
  if (fieldType === 'Boolean') {
    if (lower === '1' || lower === 'true' || lower === 'ja' || lower === 'yes') {
      return 'true';
    }
    if (lower === '0' || lower === 'false' || lower === 'nein' || lower === 'no') {
      return 'false';
    }
  }

  return defaultValue;
}

/**
 * Extract a user-friendly error message from a SharePoint REST error
 * or any other Error object. Tries to pull out the odata.error message
 * value from the JSON body; falls back to error.message.
 */
export function extractErrorMessage(error: Error): string {
  const raw: string = error && error.message ? error.message : String(error);
  // Try to extract the odata.error JSON from PnPJS error messages
  // Format: "Error making HttpClient request ... ::> { JSON }"
  const jsonStart: number = raw.indexOf('::>');
  if (jsonStart >= 0) {
    const jsonPart: string = raw.substring(jsonStart + 3).trim();
    try {
      const parsed: { 'odata.error'?: { message?: { value?: string } } } =
        JSON.parse(jsonPart);
      if (parsed['odata.error'] &&
          parsed['odata.error'].message &&
          parsed['odata.error'].message.value) {
        return parsed['odata.error'].message.value;
      }
    } catch (e) {
      // JSON parse failed — fall through to raw message
    }
  }
  return raw;
}

/**
 * Create a fresh (idle) import progress object.
 */
export function resetProgress(): IImportProgress {
  return {
    total: 0,
    current: 0,
    created: 0,
    updated: 0,
    errors: 0,
    errorMessages: [],
    status: 'idle'
  };
}

/**
 * Replace `{0}`, `{1}`, ... placeholders in a template with the
 * supplied arguments. Used for localised strings that come from
 * the loc/*.js files.
 */
// tslint:disable-next-line:no-any
export function formatString(template: string, ...args: any[]): string {
  return template.replace(/\{(\d+)\}/g, (match: string, index: string) => {
    const i: number = parseInt(index, 10);
    return i < args.length ? String(args[i]) : match;
  });
}
