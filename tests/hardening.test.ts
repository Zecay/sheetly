import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readCSV } from '../src/core/csv';
import { readXlsx } from '../src/core/xlsx';
import { Workbook } from '../src/core/model';
import { toJSON } from '../src/core/serialize';

describe('hardening', () => {
  it('csv keeps leading zeros, long digit strings and "=" text', () => {
    const s = readCSV('00123,1234567890123456789,=A1\n').sheets[0];
    expect(s.get(0, 0)?.v).toBe('00123');
    expect(s.get(0, 1)?.v).toBe('1234567890123456789');
    expect(s.get(0, 2)?.v).toBe('=A1');
  });
  it('csv errors when a quote is never closed', () => {
    expect(() => readCSV('a,"b\nc')).toThrow(/quote/);
  });
  it('xlsx with shared-string cells but no sharedStrings part throws', () => {
    const wbXml = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>';
    const sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>';
    const zip = zipSync({ 'xl/workbook.xml': strToU8(wbXml), 'xl/_rels/workbook.xml.rels': strToU8(rels), 'xl/worksheets/sheet1.xml': strToU8(sheet) });
    expect(() => readXlsx(zip)).toThrow();
  });
  it('native format never writes NaN as null', () => {
    const wb = new Workbook();
    wb.setValue(wb.sheets[0], 0, 0, NaN);
    expect(toJSON(wb)).toContain('#NUM!');
  });
});
