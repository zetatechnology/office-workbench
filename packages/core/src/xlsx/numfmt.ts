// A pragmatic subset of Excel number formatting: General, fixed/thousands/percent/scientific,
// dates and times, literal text and multi-section formats. Fractions and fill characters
// are approximated.

export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'yyyy/m/d',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'yyyy/m/d h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
  // Japanese-locale built-ins.
  27: 'ggge"年"m"月"d"日"',
  28: 'mm"月"dd"日"',
  29: 'mm"月"dd"日"',
  30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"',
  32: 'h"時"mm"分"',
  33: 'h"時"mm"分"ss"秒"',
  34: 'yyyy"年"m"月"',
  35: 'm"月"d"日"',
  36: 'ggge"年"m"月"d"日"',
  50: 'ggge"年"m"月"d"日"',
  51: 'yyyy"年"m"月"',
  52: 'm"月"d"日"',
  53: 'yyyy"年"m"月"',
  54: 'm"月"d"日"',
  55: 'yyyy"年"m"月"',
  56: 'm"月"d"日"',
  57: 'ggge"年"m"月"d"日"',
  58: 'ggge"年"m"月"d"日"',
};

interface Section {
  pattern: string;
  condition?: { op: string; value: number };
}

const formatCache = new Map<string, Section[]>();

function splitSections(format: string): Section[] {
  let cached = formatCache.get(format);
  if (cached) return cached;
  const parts: string[] = [];
  let current = '';
  let quoted = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') quoted = !quoted;
    if (ch === '\\' && !quoted) {
      current += ch + (format[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === ';' && !quoted) {
      parts.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  parts.push(current);
  cached = parts.map((pattern) => {
    const m = /\[(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)\]/.exec(pattern);
    return m ? { pattern, condition: { op: m[1], value: Number(m[2]) } } : { pattern };
  });
  formatCache.set(format, cached);
  return cached;
}

/** Removes bracketed codes ([Red], [$-411], [DBNum1], conditions) but keeps elapsed-time codes. */
function stripBrackets(pattern: string): string {
  return pattern.replace(/\[(?![hms]+\])[^\]]*\]/gi, '');
}

const DATE_TOKEN = /y|m|d|h|s|e|g|aaa|AM\/PM|A\/P|\[h+\]|\[m+\]|\[s+\]/i;

export function isDateFormat(format: string): boolean {
  if (/^general$/i.test(format.trim())) return false;
  const first = stripBrackets(splitSections(format)[0].pattern)
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/E[+-]/gi, ''); // Scientific notation, not an era year.
  return DATE_TOKEN.test(first);
}

export function formatValue(value: number, format: string, date1904 = false): string {
  if (!format || format === 'General') return formatGeneral(value);
  const sections = splitSections(format);
  const section = pickSection(sections, value);
  const pattern = stripBrackets(section.pattern);
  // Negative sections carry their own sign (e.g. "(#,##0)").
  const useAbs = sections.length > 1 && value < 0 && sections.indexOf(section) > 0 && !section.condition;
  const v = useAbs ? Math.abs(value) : value;
  if (pattern === '') return '';
  if (/^general$/i.test(pattern.trim())) return formatGeneral(v);
  if (isDateFormat(pattern)) return formatDate(v, pattern, date1904);
  return formatNumber(v, pattern);
}

function pickSection(sections: Section[], value: number): Section {
  if (sections.some((s) => s.condition)) {
    for (const s of sections) {
      if (!s.condition) return s;
      const { op, value: c } = s.condition;
      const ok =
        op === '<' ? value < c : op === '<=' ? value <= c : op === '>' ? value > c :
        op === '>=' ? value >= c : op === '=' ? value === c : value !== c;
      if (ok) return s;
    }
    return sections[sections.length - 1];
  }
  if (value < 0 && sections.length > 1) return sections[1];
  if (value === 0 && sections.length > 2) return sections[2];
  return sections[0];
}

export function formatGeneral(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e11) return String(value);
  const abs = Math.abs(value);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    const [mantissa, exp] = value.toExponential(5).split('e');
    const m = mantissa.replace(/\.?0+$/, '');
    const e = Number(exp);
    return `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  // Excel shows up to 11 significant characters in a default-width column.
  const digitsBefore = Math.max(1, Math.floor(Math.log10(abs)) + 1);
  const decimals = Math.max(0, 10 - digitsBefore);
  return String(Number(value.toFixed(decimals)));
}

/** Splits a pattern into literal text and placeholder runs, honouring quotes and escapes. */
function tokenize(pattern: string): { literal: boolean; text: string }[] {
  const tokens: { literal: boolean; text: string }[] = [];
  const push = (literal: boolean, t: string) => {
    const last = tokens[tokens.length - 1];
    if (last && last.literal === literal) last.text += t;
    else tokens.push({ literal, text: t });
  };
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '"') {
      const end = pattern.indexOf('"', i + 1);
      push(true, pattern.slice(i + 1, end < 0 ? undefined : end));
      i = end < 0 ? pattern.length : end;
    } else if (ch === '\\') {
      push(true, pattern[i + 1] ?? '');
      i++;
    } else if (ch === '_') {
      push(true, ' ');
      i++;
    } else if (ch === '*') {
      i++; // Fill character: ignored.
    } else {
      push(false, ch);
    }
  }
  return tokens;
}

function formatNumber(value: number, pattern: string): string {
  const tokens = tokenize(pattern);
  const code = tokens.filter((t) => !t.literal).map((t) => t.text).join('');
  if (code.includes('@')) {
    return tokens.map((t) => (t.literal ? t.text : t.text.replace('@', formatGeneral(value)))).join('');
  }
  const placeholders = /[0#?.,]/;
  if (!/[0#?]/.test(code)) {
    // Literal-only section, e.g. "-" for zero.
    return tokens.map((t) => t.text).join('');
  }
  if (code.includes('/')) return formatGeneral(value); // Fractions: not supported yet.

  const percentCount = (code.match(/%/g) ?? []).length;
  let v = value * Math.pow(100, percentCount);
  const sci = /E[+-]/i.exec(code);
  const numberPart = code.replace(/%/g, '');

  const [intPart, rest = ''] = numberPart.split(sci ? /E[+-]/i : /(?!)/);
  const [intCode, decCode = ''] = intPart.split('.');
  // Trailing commas scale by 1000.
  const trailing = /,+$/.exec(intCode);
  if (trailing && !sci) v /= Math.pow(1000, trailing[0].length);
  const cleanInt = intCode.replace(/,+$/, '');
  const thousands = cleanInt.includes(',');
  const decimals = (decCode.match(/[0#?]/g) ?? []).length;
  const minDecimals = (decCode.match(/0/g) ?? []).length;
  const minInt = (cleanInt.match(/0/g) ?? []).length;

  let body: string;
  if (sci) {
    const expDigits = (rest.match(/0/g) ?? []).length || 2;
    if (v === 0) {
      body = `${(0).toFixed(decimals)}E+${'0'.repeat(expDigits)}`;
    } else {
      const exp = Math.floor(Math.log10(Math.abs(v)));
      const mantissa = v / Math.pow(10, exp);
      const sign = exp < 0 ? '-' : '+';
      body = `${trimDecimals(Math.abs(mantissa).toFixed(decimals), minDecimals)}E${sign}${String(Math.abs(exp)).padStart(expDigits, '0')}`;
      if (mantissa < 0) body = '-' + body;
    }
  } else {
    const negative = v < 0;
    let fixed = trimDecimals(Math.abs(v).toFixed(decimals), minDecimals);
    const [intDigits, d] = fixed.split('.');
    let i = intDigits;
    if (i === '0' && minInt === 0) i = '';
    i = i.padStart(minInt, '0');
    if (thousands) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    fixed = d !== undefined ? `${i}.${d}` : i;
    body = (negative && Number(fixed.replace(/,/g, '')) !== 0 ? '-' : '') + fixed;
  }

  // Re-insert literals around the numeric body: text before the first and after the last placeholder.
  let prefix = '';
  let suffix = '';
  let seenPlaceholder = false;
  for (const t of tokens) {
    if (t.literal) {
      if (seenPlaceholder) suffix += t.text;
      else prefix += t.text;
      continue;
    }
    for (let i = 0; i < t.text.length; i++) {
      const ch = t.text[i];
      if (ch === '%') {
        suffix += '%';
        seenPlaceholder = true;
      } else if (sci && (ch === 'E' || ch === 'e') && (t.text[i + 1] === '+' || t.text[i + 1] === '-')) {
        i++; // The exponent is part of the numeric body.
      } else if (placeholders.test(ch)) {
        seenPlaceholder = true;
      } else if (seenPlaceholder) {
        suffix += ch;
      } else {
        prefix += ch;
      }
    }
  }
  return prefix + body + suffix;
}

function trimDecimals(fixed: string, minDecimals: number): string {
  if (!fixed.includes('.')) return fixed;
  const [i, decimals] = fixed.split('.');
  let d = decimals;
  while (d.length > minDecimals && d.endsWith('0')) d = d.slice(0, -1);
  return d.length ? `${i}.${d}` : i;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAYS_JA = ['日', '月', '火', '水', '木', '金', '土'];

interface DateParts {
  y: number;
  mo: number;
  d: number;
  wd: number;
  h: number;
  mi: number;
  s: number;
  ms: number;
}

/** Converts an Excel serial date to calendar parts (1900 system includes the 1900 leap-year bug). */
export function serialToDate(serial: number, date1904 = false): DateParts {
  let days = Math.floor(serial);
  const frac = serial - days;
  if (date1904) days += 1462;
  // Serial 60 is the fictitious 1900-02-29.
  if (!date1904 && days === 60) {
    return { y: 1900, mo: 2, d: 29, wd: 3, ...timeParts(frac) };
  }
  if (!date1904 && days < 60) days += 1;
  const ms = Date.UTC(1899, 11, 30) + days * 86_400_000;
  const dt = new Date(ms);
  return { y: dt.getUTCFullYear(), mo: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay(), ...timeParts(frac) };
}

function timeParts(frac: number): { h: number; mi: number; s: number; ms: number } {
  let totalMs = Math.round(frac * 86_400_000);
  if (totalMs >= 86_400_000) totalMs = 86_400_000 - 1;
  const h = Math.floor(totalMs / 3_600_000);
  const mi = Math.floor((totalMs % 3_600_000) / 60_000);
  const s = Math.floor((totalMs % 60_000) / 1000);
  return { h, mi, s, ms: totalMs % 1000 };
}

function formatDate(value: number, pattern: string, date1904: boolean): string {
  const p = serialToDate(value, date1904);
  const tokens = tokenize(pattern);
  const hasAmPm = /AM\/PM|A\/P/i.test(pattern);
  const out: string[] = [];
  // Collect date tokens so we can decide whether "m" means month or minute.
  const parts: { kind: 'lit' | 'tok'; text: string }[] = [];
  for (const t of tokens) {
    if (t.literal) {
      parts.push({ kind: 'lit', text: t.text });
      continue;
    }
    const re = /\[h+\]|\[m+\]|\[s+\]|AM\/PM|A\/P|yyyy|yy|y+|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|\.0+|ggg|gg|g|ee|e|aaaa|aaa|./gi;
    let m: RegExpExecArray | null;
    while ((m = re.exec(t.text))) {
      const tok = m[0];
      parts.push({ kind: /^(\[|[ymdhsge]|aaa|AM\/PM|A\/P|\.0)/i.test(tok) ? 'tok' : 'lit', text: tok });
    }
  }
  const tokIndexes = parts.map((x, i) => (x.kind === 'tok' ? i : -1)).filter((i) => i >= 0);
  const isMinute = (idx: number) => {
    const pos = tokIndexes.indexOf(idx);
    const prev = tokIndexes[pos - 1] !== undefined ? parts[tokIndexes[pos - 1]].text.toLowerCase() : '';
    const next = tokIndexes[pos + 1] !== undefined ? parts[tokIndexes[pos + 1]].text.toLowerCase() : '';
    return prev.startsWith('h') || prev.startsWith('[h') || next.startsWith('s');
  };
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  const hour12 = p.h % 12 === 0 ? 12 : p.h % 12;
  parts.forEach((part, idx) => {
    if (part.kind === 'lit') {
      out.push(part.text);
      return;
    }
    const tok = part.text;
    const lower = tok.toLowerCase();
    if (/^\[h+\]$/.test(lower)) out.push(String(Math.floor(value * 24)));
    else if (/^\[m+\]$/.test(lower)) out.push(String(Math.floor(value * 1440)));
    else if (/^\[s+\]$/.test(lower)) out.push(String(Math.floor(value * 86400)));
    else if (lower === 'am/pm') out.push(p.h < 12 ? 'AM' : 'PM');
    else if (lower === 'a/p') out.push(p.h < 12 ? 'A' : 'P');
    else if (lower.startsWith('y')) out.push(lower.length <= 2 ? pad(p.y % 100) : String(p.y));
    else if (lower === 'e' || lower === 'ee') out.push(String(p.y)); // Era years: approximated as Gregorian.
    else if (lower.startsWith('g')) out.push('');
    else if (lower.startsWith('m')) {
      if (lower.length <= 2 && isMinute(idx)) out.push(lower.length === 2 ? pad(p.mi) : String(p.mi));
      else if (lower === 'mmmmm') out.push(MONTHS[p.mo - 1][0]);
      else if (lower === 'mmmm') out.push(MONTHS[p.mo - 1]);
      else if (lower === 'mmm') out.push(MONTHS[p.mo - 1].slice(0, 3));
      else out.push(lower === 'mm' ? pad(p.mo) : String(p.mo));
    } else if (lower === 'aaaa') out.push(DAYS_JA[p.wd] + '曜日');
    else if (lower === 'aaa') out.push(DAYS_JA[p.wd]);
    else if (lower.startsWith('d')) {
      if (lower === 'dddd') out.push(DAYS[p.wd]);
      else if (lower === 'ddd') out.push(DAYS[p.wd].slice(0, 3));
      else out.push(lower === 'dd' ? pad(p.d) : String(p.d));
    } else if (lower.startsWith('h')) {
      const h = hasAmPm ? hour12 : p.h;
      out.push(lower.length === 2 ? pad(h) : String(h));
    } else if (lower.startsWith('s')) out.push(lower.length === 2 ? pad(p.s) : String(p.s));
    else if (lower.startsWith('.')) out.push('.' + pad(p.ms, 3).slice(0, lower.length - 1));
    else out.push(tok);
  });
  return out.join('');
}
