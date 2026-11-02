import { describe, expect, it } from 'vitest';
import { formatGeneral, formatValue, isDateFormat, serialToDate } from '../src/xlsx/numfmt';

describe('formatGeneral', () => {
  it.each([
    [0, '0'],
    [42, '42'],
    [-3.5, '-3.5'],
    [0.1 + 0.2, '0.3'],
    [1234567.891, '1234567.891'],
    [123456789012, '1.23457E+11'],
    [0.000000000123, '1.23E-10'],
  ])('%s → %s', (value, expected) => {
    expect(formatGeneral(value)).toBe(expected);
  });
});

describe('formatValue', () => {
  it.each([
    [1234.5, '#,##0', '1,235'],
    [1234.5, '#,##0.00', '1,234.50'],
    [-1234.5, '#,##0.00;[Red]-#,##0.00', '-1,234.50'],
    [-1234, '#,##0 ;(#,##0)', '(1,234)'],
    [0.875, '0.0%', '87.5%'],
    [0.5, '0%', '50%'],
    [12345.678, '0.00E+00', '1.23E+04'],
    [64290, '"¥"#,##0', '¥64,290'],
    [5, '0" 個"', '5 個'],
    [0, '#,##0;-#,##0;"-"', '-'],
    [3, '000', '003'],
    [1500000, '#,##0,"K"', '1,500K'],
  ])('%s with %s → %s', (value, format, expected) => {
    expect(formatValue(value, format)).toBe(expected);
  });

  it('formats dates and times', () => {
    expect(formatValue(46037, 'yyyy/m/d')).toBe('2026/1/15');
    expect(formatValue(46037, 'yyyy"年"m"月"d"日"')).toBe('2026年1月15日');
    expect(formatValue(46037.75, 'yyyy-mm-dd hh:mm')).toBe('2026-01-15 18:00');
    expect(formatValue(46037.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(formatValue(1.5, '[h]:mm:ss')).toBe('36:00:00');
    expect(formatValue(46037, 'mmm d, yyyy')).toBe('Jan 15, 2026');
    expect(formatValue(46037, 'aaa')).toBe('木');
  });

  it('detects date formats', () => {
    expect(isDateFormat('yyyy/m/d')).toBe(true);
    expect(isDateFormat('[$-411]ggge"年"m"月"d"日"')).toBe(true);
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('"day"0')).toBe(false);
    expect(isDateFormat('[Red]0.00')).toBe(false);
  });
});

describe('serialToDate', () => {
  it('handles the 1900 leap-year bug', () => {
    expect(serialToDate(1)).toMatchObject({ y: 1900, mo: 1, d: 1 });
    expect(serialToDate(59)).toMatchObject({ y: 1900, mo: 2, d: 28 });
    expect(serialToDate(60)).toMatchObject({ y: 1900, mo: 2, d: 29 });
    expect(serialToDate(61)).toMatchObject({ y: 1900, mo: 3, d: 1 });
  });

  it('supports the 1904 date system', () => {
    expect(serialToDate(0, true)).toMatchObject({ y: 1904, mo: 1, d: 1 });
  });
});
