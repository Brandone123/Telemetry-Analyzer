/**
 * Surgical xlsx editor: opens a workbook (JSZip), allows updating individual cells
 * by sheet name + column letter while preserving every other byte of the file
 * (native pivot tables, native charts, drawings, conditional formatting, styles…).
 *
 * Used for the SLA Tracker template: we update today's date column in Comm Issue,
 * and append today's row to NO COM HISTORY. Native charts referencing those cells
 * automatically refresh in Excel; pivot tables refresh on file open.
 */
import JSZip from "jszip";

export interface SheetMeta {
  name: string;
  rId: string;
  filePath: string; // e.g. xl/worksheets/sheet3.xml
}

export class XlsxTemplate {
  private zip: JSZip;
  private sheets: SheetMeta[] = [];
  private sheetXmlCache = new Map<string, string>();

  private constructor(zip: JSZip) {
    this.zip = zip;
  }

  static async load(buffer: Buffer): Promise<XlsxTemplate> {
    const zip = await JSZip.loadAsync(buffer);
    const t = new XlsxTemplate(zip);
    await t.parseSheetIndex();
    return t;
  }

  private async parseSheetIndex(): Promise<void> {
    const wbXml = await this.zip.file("xl/workbook.xml")!.async("string");
    const relsXml = await this.zip.file("xl/_rels/workbook.xml.rels")!.async("string");
    const rels = new Map<string, string>();
    for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
      rels.set(m[1]!, m[2]!);
    }
    for (const m of wbXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
      const target = rels.get(m[2]!) ?? "";
      const filePath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
      this.sheets.push({ name: m[1]!, rId: m[2]!, filePath });
    }
  }

  getSheet(name: string): SheetMeta | undefined {
    return this.sheets.find((s) => s.name === name);
  }

  async getSheetXml(name: string): Promise<string> {
    if (this.sheetXmlCache.has(name)) return this.sheetXmlCache.get(name)!;
    const meta = this.getSheet(name);
    if (!meta) throw new Error(`Sheet not found: ${name}`);
    const xml = await this.zip.file(meta.filePath)!.async("string");
    this.sheetXmlCache.set(name, xml);
    return xml;
  }

  setSheetXml(name: string, xml: string): void {
    const meta = this.getSheet(name);
    if (!meta) throw new Error(`Sheet not found: ${name}`);
    this.sheetXmlCache.set(name, xml);
    this.zip.file(meta.filePath, xml);
  }

  /**
   * Read header values from row 1 of a sheet. Returns map of column letter -> value (string).
   * For numeric headers (e.g. date serials), value is the number as string.
   */
  async getHeaderRow(name: string): Promise<Map<string, string>> {
    const xml = await this.getSheetXml(name);
    const result = new Map<string, string>();
    const rowMatch = xml.match(/<row[^>]+r="1"[^>]*>([\s\S]*?)<\/row>/);
    if (!rowMatch) return result;
    const content = rowMatch[1]!;
    // shared strings, if any
    let sharedStrings: string[] | null = null;
    for (const c of content.matchAll(/<c\s+r="([A-Z]+)1"([^>]*)>([\s\S]*?)<\/c>/g)) {
      const ref = c[1]!;
      const attrs = c[2]!;
      const inner = c[3]!;
      const isShared = /\bt="s"/.test(attrs);
      const isInline = /\bt="(inlineStr|str)"/.test(attrs);
      const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
      const tMatch = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/);
      let value = "";
      if (isShared && vMatch) {
        if (!sharedStrings) sharedStrings = await this.getSharedStrings();
        const idx = parseInt(vMatch[1]!, 10);
        value = sharedStrings[idx] ?? "";
      } else if (isInline && tMatch) {
        value = tMatch[1]!;
      } else if (vMatch) {
        value = vMatch[1]!;
      }
      result.set(ref, value);
    }
    return result;
  }

  private cachedSharedStrings: string[] | null = null;
  private async getSharedStrings(): Promise<string[]> {
    if (this.cachedSharedStrings) return this.cachedSharedStrings;
    const file = this.zip.file("xl/sharedStrings.xml");
    if (!file) return (this.cachedSharedStrings = []);
    const xml = await file.async("string");
    const out: string[] = [];
    for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      const inner = m[1]!;
      const tMatches = [...inner.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
      out.push(tMatches.map((tm) => tm[1]!).join(""));
    }
    this.cachedSharedStrings = out;
    return out;
  }

  /**
   * Update existing numeric/value cells within a single column over a given row range.
   * Cells must already exist (typical for a pre-formatted tracker). Replaces the inner
   * <v> while preserving cell attributes (style index s="...").
   */
  async updateColumnValues(
    sheetName: string,
    columnLetter: string,
    valuesByRow: Map<number, number>,
  ): Promise<{ updated: number; missing: number }> {
    if (valuesByRow.size === 0) return { updated: 0, missing: 0 };
    let xml = await this.getSheetXml(sheetName);

    // Build a lookup map: "COL<row>" → numeric value  (e.g. "AD42" → 1)
    const refToVal = new Map<string, number>();
    for (const [rowNum, val] of valuesByRow) {
      refToVal.set(`${columnLetter}${rowNum}`, val);
    }

    let updated = 0;

    // Single pass through the XML: match every <c r="..."> and replace when in map.
    xml = xml.replace(
      /<c\s+r="([^"]+)"([^>]*?)(?:\s*\/>|>([\s\S]*?)<\/c>)/g,
      (match, ref: string, attrs: string) => {
        if (!refToVal.has(ref)) return match;
        const val = refToVal.get(ref)!;
        refToVal.delete(ref);          // mark as found
        const cleanAttrs = attrs.replace(/\s+t="[^"]*"/g, "");
        updated++;
        return `<c r="${ref}"${cleanAttrs}><v>${val}</v></c>`;
      },
    );

    this.setSheetXml(sheetName, xml);
    return { updated, missing: refToVal.size };
  }

  /**
   * Replace one cell with a formula. Preserves style attribute. The given `value` is
   * stored as the cached result so Excel shows it before recompute.
   */
  async updateCellFormula(
    sheetName: string,
    cellRef: string,
    formula: string,
    cachedValue: number | string,
  ): Promise<boolean> {
    let xml = await this.getSheetXml(sheetName);
    const cellRe = new RegExp(
      `<c\\s+r="${cellRef}"([^>]*?)(?:\\s*/>|>([\\s\\S]*?)</c>)`,
      "",
    );
    const m = xml.match(cellRe);
    let attrs = "";
    if (m) {
      attrs = m[1]!.replace(/\s+t="[^"]*"/g, "");
      const replacement = `<c r="${cellRef}"${attrs}><f>${escapeXml(formula)}</f><v>${cachedValue}</v></c>`;
      xml = xml.replace(cellRe, replacement);
      this.setSheetXml(sheetName, xml);
      return true;
    }
    // Cell doesn't exist; insert into existing row or create new row
    const rowNum = parseInt(cellRef.match(/\d+$/)![0], 10);
    const rowRe = new RegExp(`<row[^>]+r="${rowNum}"[^>]*>([\\s\\S]*?)</row>`);
    const rm = xml.match(rowRe);
    const newCell = `<c r="${cellRef}"><f>${escapeXml(formula)}</f><v>${cachedValue}</v></c>`;
    if (rm) {
      xml = xml.replace(rowRe, (whole) => whole.replace("</row>", newCell + "</row>"));
    } else {
      // Insert a new row before </sheetData>
      const newRow = `<row r="${rowNum}">${newCell}</row>`;
      xml = xml.replace("</sheetData>", `${newRow}</sheetData>`);
    }
    this.setSheetXml(sheetName, xml);
    return false;
  }

  /**
   * Set or insert a cell holding a string (inline string).
   */
  async setCellString(sheetName: string, cellRef: string, text: string): Promise<void> {
    let xml = await this.getSheetXml(sheetName);
    const cellRe = new RegExp(
      `<c\\s+r="${cellRef}"([^>]*?)(?:\\s*/>|>([\\s\\S]*?)</c>)`,
      "",
    );
    const m = xml.match(cellRe);
    const safe = escapeXml(text);
    if (m) {
      let attrs = m[1]!.replace(/\s+t="[^"]*"/g, "");
      attrs += ' t="inlineStr"';
      const replacement = `<c r="${cellRef}"${attrs}><is><t>${safe}</t></is></c>`;
      xml = xml.replace(cellRe, replacement);
    } else {
      const rowNum = parseInt(cellRef.match(/\d+$/)![0], 10);
      const newCell = `<c r="${cellRef}" t="inlineStr"><is><t>${safe}</t></is></c>`;
      const rowRe = new RegExp(`<row[^>]+r="${rowNum}"[^>]*>([\\s\\S]*?)</row>`);
      if (rowRe.test(xml)) {
        xml = xml.replace(rowRe, (whole) => whole.replace("</row>", newCell + "</row>"));
      } else {
        xml = xml.replace("</sheetData>", `<row r="${rowNum}">${newCell}</row></sheetData>`);
      }
    }
    this.setSheetXml(sheetName, xml);
  }

  /**
   * Set or insert a numeric cell.
   */
  async setCellNumber(sheetName: string, cellRef: string, value: number): Promise<void> {
    let xml = await this.getSheetXml(sheetName);
    const cellRe = new RegExp(
      `<c\\s+r="${cellRef}"([^>]*?)(?:\\s*/>|>([\\s\\S]*?)</c>)`,
      "",
    );
    const m = xml.match(cellRe);
    if (m) {
      const attrs = m[1]!.replace(/\s+t="[^"]*"/g, "");
      const replacement = `<c r="${cellRef}"${attrs}><v>${value}</v></c>`;
      xml = xml.replace(cellRe, replacement);
    } else {
      const rowNum = parseInt(cellRef.match(/\d+$/)![0], 10);
      const newCell = `<c r="${cellRef}"><v>${value}</v></c>`;
      const rowRe = new RegExp(`<row[^>]+r="${rowNum}"[^>]*>([\\s\\S]*?)</row>`);
      if (rowRe.test(xml)) {
        xml = xml.replace(rowRe, (whole) => whole.replace("</row>", newCell + "</row>"));
      } else {
        xml = xml.replace("</sheetData>", `<row r="${rowNum}">${newCell}</row></sheetData>`);
      }
    }
    this.setSheetXml(sheetName, xml);
  }

  /**
   * Blank a cell's content while keeping the cell element and its style. Used to
   * clear rows that fall outside a rolling window after it shrinks.
   */
  async clearCell(sheetName: string, cellRef: string): Promise<void> {
    let xml = await this.getSheetXml(sheetName);
    const cellRe = new RegExp(
      `<c\\s+r="${cellRef}"([^>]*?)(?:\\s*/>|>([\\s\\S]*?)</c>)`,
      "",
    );
    const m = xml.match(cellRe);
    if (!m) return;
    const attrs = m[1]!.replace(/\s+t="[^"]*"/g, "");
    xml = xml.replace(cellRe, `<c r="${cellRef}"${attrs}/>`);
    this.setSheetXml(sheetName, xml);
  }

  /**
   * Read the first category (`c:cat`) data range of a chart, e.g. for
   * ref="'NO COM HISTORY'!$E$4:$E$8" returns { startRow: 4, endRow: 8 }.
   * Used to size a rolling window to exactly the cells the chart plots.
   */
  async getChartCategoryRange(
    chartNum: number,
  ): Promise<{ startRow: number; endRow: number } | null> {
    const f = this.zip.file(`xl/charts/chart${chartNum}.xml`);
    if (!f) return null;
    const xml = await f.async("string");
    const m = xml.match(
      /<c:cat>[\s\S]*?<c:f>[^<]*?\$[A-Z]+\$(\d+):\$[A-Z]+\$(\d+)<\/c:f>/,
    );
    if (!m) return null;
    return { startRow: parseInt(m[1]!, 10), endRow: parseInt(m[2]!, 10) };
  }

  /**
   * Overwrite the cached points of a chart's category (`c:cat`) and value
   * (`c:val`) series. Native charts embed a cache of the plotted values; a
   * surgical cell edit does not touch it, so viewers (and Excel before a
   * recalc) keep showing the old chart. Rewriting the cache makes the rendered
   * chart match the new data immediately. The `c:f` cell references and any
   * `c:formatCode` are preserved.
   */
  async updateChartSeriesCache(
    chartNum: number,
    categories: Array<string | number>,
    values: number[],
  ): Promise<void> {
    const path = `xl/charts/chart${chartNum}.xml`;
    const f = this.zip.file(path);
    if (!f) return;
    let xml = await f.async("string");
    xml = xml.replace(
      /(<c:cat>)([\s\S]*?)(<\/c:cat>)/,
      (_, a: string, inner: string, b: string) =>
        a + replaceCachePoints(inner, categories.map(String)) + b,
    );
    xml = xml.replace(
      /(<c:val>)([\s\S]*?)(<\/c:val>)/,
      (_, a: string, inner: string, b: string) =>
        a + replaceCachePoints(inner, values.map((v) => String(v))) + b,
    );
    this.zip.file(path, xml);
  }

  /**
   * Read every cell in a single column of a sheet (rows >= 2). Returns a map
   * of row number → raw string value (numbers come through as their text form).
   */
  async getColumnValuesByRow(sheetName: string, columnLetter: string): Promise<Map<number, string>> {
    const xml = await this.getSheetXml(sheetName);
    const result = new Map<number, string>();
    const re = new RegExp(
      `<c\\s+r="${columnLetter}(\\d+)"[^>]*?(?:\\s*/>|>([\\s\\S]*?)<\\/c>)`,
      "g",
    );
    for (const m of xml.matchAll(re)) {
      const row = parseInt(m[1]!, 10);
      const inner = m[2] ?? "";
      const v = inner.match(/<v>([\s\S]*?)<\/v>/);
      const t = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/);
      result.set(row, v?.[1] ?? t?.[1] ?? "");
    }
    return result;
  }

  /**
   * Return the current end column letter of a pivot cache's worksheetSource ref,
   * e.g. "AC" for ref="A1:AC1048576". null if the cache or ref is missing.
   */
  async getPivotCacheEndColumn(cacheNum: number): Promise<string | null> {
    const path = `xl/pivotCache/pivotCacheDefinition${cacheNum}.xml`;
    const f = this.zip.file(path);
    if (!f) return null;
    const xml = await f.async("string");
    const m = xml.match(/<worksheetSource\b[^/]*ref="[A-Z]+\d+:([A-Z]+)\d+"/);
    return m ? m[1]! : null;
  }

  /**
   * Extend a pivot cache definition's worksheetSource range to a new end column
   * and append placeholder cacheField entries. Excel will rebuild shared items
   * and records on open thanks to refreshOnLoad="1".
   */
  async extendPivotCacheToColumn(
    cacheNum: number,
    newEndColumnLetter: string,
    newFieldNames: string[],
  ): Promise<void> {
    const path = `xl/pivotCache/pivotCacheDefinition${cacheNum}.xml`;
    const f = this.zip.file(path);
    if (!f) return;
    let xml = await f.async("string");
    // Update worksheetSource ref
    xml = xml.replace(
      /<worksheetSource\b([^/]*?)ref="([^"]+)"([^/]*?)\/>/,
      (_, b, ref, c) => {
        const m = ref.match(/^([A-Z]+)1:([A-Z]+)(\d+)$/);
        if (!m) return `<worksheetSource${b}ref="A1:${newEndColumnLetter}1048576"${c}/>`;
        return `<worksheetSource${b}ref="${m[1]}1:${newEndColumnLetter}${m[3]}"${c}/>`;
      },
    );
    if (newFieldNames.length > 0) {
      const countMatch = xml.match(/<cacheFields\s+count="(\d+)"/);
      if (countMatch) {
        const oldCount = parseInt(countMatch[1]!, 10);
        const newCount = oldCount + newFieldNames.length;
        xml = xml.replace(/<cacheFields\s+count="\d+"/, `<cacheFields count="${newCount}"`);
      }
      const fieldsXml = newFieldNames.map((n) =>
        `<cacheField name="${escapeXml(n)}" numFmtId="0">` +
        `<sharedItems containsString="0" containsBlank="1" containsNumber="1" containsInteger="1" minValue="0" maxValue="1"/>` +
        `</cacheField>`,
      ).join("");
      xml = xml.replace("</cacheFields>", fieldsXml + "</cacheFields>");
    }
    // Make sure refresh-on-load is on
    if (!/refreshOnLoad="1"/.test(xml)) {
      if (/refreshOnLoad="\d"/.test(xml)) {
        xml = xml.replace(/refreshOnLoad="\d"/, 'refreshOnLoad="1"');
      } else {
        xml = xml.replace(
          /<pivotCacheDefinition\b([^>]*?)>/,
          (_, attrs) => `<pivotCacheDefinition${attrs} refreshOnLoad="1">`,
        );
      }
    }
    this.zip.file(path, xml);
  }

  /**
   * Move a pivot table's column-axis pivotField to a new field index. The existing
   * axisCol pivotField (with its <items>) is moved; any intermediate empty fields
   * are appended. Updates pivotFields count and <colFields> field x.
   */
  async retargetPivotColumnField(
    pivotNum: number,
    newFieldIndex: number,
  ): Promise<void> {
    const path = `xl/pivotTables/pivotTable${pivotNum}.xml`;
    const f = this.zip.file(path);
    if (!f) return;
    let xml = await f.async("string");

    const axisColRe = /<pivotField\s+axis="axisCol"[^>]*>([\s\S]*?)<\/pivotField>/;
    const axisMatch = xml.match(axisColRe);
    if (!axisMatch) return;
    const itemsContent = axisMatch[1]!;

    // Replace the old axisCol pivotField with a regular one (preserve count)
    xml = xml.replace(axisColRe, '<pivotField showAll="0"/>');

    const countMatch = xml.match(/<pivotFields\s+count="(\d+)"/);
    const oldCount = countMatch ? parseInt(countMatch[1]!, 10) : 0;
    const addEmpty = Math.max(0, newFieldIndex - oldCount);
    const newCount = oldCount + addEmpty + 1; // +1 for the new axisCol at the end

    const padding = '<pivotField showAll="0"/>'.repeat(addEmpty);
    const newAxisField = `<pivotField axis="axisCol" showAll="0">${itemsContent}</pivotField>`;
    xml = xml.replace("</pivotFields>", padding + newAxisField + "</pivotFields>");
    xml = xml.replace(/<pivotFields\s+count="\d+"/, `<pivotFields count="${newCount}"`);

    // Update <colFields> ... <field x="OLD"/></colFields> — last field is the column dim
    xml = xml.replace(
      /(<colFields\b[^>]*>(?:<field\s+x="-?\d+"\/>)*?)<field\s+x="\d+"\/>(<\/colFields>)/,
      (_, head, tail) => `${head}<field x="${newFieldIndex}"/>${tail}`,
    );

    this.zip.file(path, xml);
  }

  /**
   * Replace text inside the <c:title> block of a chart by regex.
   */
  async updateChartTitle(
    chartNum: number,
    replacements: Array<{ from: RegExp; to: string }>,
  ): Promise<void> {
    const path = `xl/charts/chart${chartNum}.xml`;
    const f = this.zip.file(path);
    if (!f) return;
    let xml = await f.async("string");
    const titleRe = /<c:title>[\s\S]*?<\/c:title>/;
    const m = xml.match(titleRe);
    if (!m) return;
    let block = m[0];
    for (const r of replacements) block = block.replace(r.from, r.to);
    if (block !== m[0]) {
      xml = xml.replace(titleRe, block);
      this.zip.file(path, xml);
    }
  }

  /**
   * Mark every pivot cache definition with refreshOnLoad="1" so Excel refreshes
   * the pivot tables when the file is opened.
   */
  async setPivotCachesRefreshOnLoad(): Promise<void> {
    const files = Object.keys(this.zip.files).filter((f) =>
      /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(f),
    );
    for (const f of files) {
      let xml = await this.zip.file(f)!.async("string");
      if (/refreshOnLoad="1"/.test(xml)) continue;
      if (/refreshOnLoad="\d"/.test(xml)) {
        xml = xml.replace(/refreshOnLoad="\d"/, 'refreshOnLoad="1"');
      } else {
        xml = xml.replace(
          /<pivotCacheDefinition\b([^>]*?)>/,
          (_, attrs) => `<pivotCacheDefinition${attrs} refreshOnLoad="1">`,
        );
      }
      this.zip.file(f, xml);
    }
  }

  /** Force full calc on open so Excel recomputes formulas. */
  async setFullCalcOnLoad(): Promise<void> {
    const f = this.zip.file("xl/workbook.xml");
    if (!f) return;
    let xml = await f.async("string");
    if (/<calcPr\b/.test(xml)) {
      if (/fullCalcOnLoad="\d"/.test(xml)) {
        xml = xml.replace(/fullCalcOnLoad="\d"/, 'fullCalcOnLoad="1"');
      } else {
        xml = xml.replace(/<calcPr\b([^/>]*)\/?>/, (_, attrs) => `<calcPr${attrs} fullCalcOnLoad="1"/>`);
      }
    } else {
      xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    this.zip.file("xl/workbook.xml", xml);
  }

  async toBuffer(): Promise<Buffer> {
    return await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  }
}

/**
 * Replace the cached points inside a chart `c:cat`/`c:val` block: rewrites
 * `c:ptCount` and the list of `c:pt` entries, preserving everything else
 * (the `c:f` reference, `c:formatCode`, cache element type, etc.).
 */
function replaceCachePoints(inner: string, vals: string[]): string {
  let out = inner.replace(
    /<c:ptCount\s+val="\d+"\s*\/>/,
    `<c:ptCount val="${vals.length}"/>`,
  );
  out = out.replace(/<c:pt\b[\s\S]*?<\/c:pt>/g, "");
  const pts = vals
    .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`)
    .join("");
  out = out.replace(/(<c:ptCount\s+val="\d+"\s*\/>)/, `$1${pts}`);
  return out;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Convert a 0-based column index to a column letter (A, B, …, Z, AA, AB, …). */
export function colLetter(idx: number): string {
  let s = "";
  let n = idx;
  while (true) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return s;
}

/** Convert a JS Date to an Excel serial number (1900 system). */
export function dateToExcelSerial(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86_400_000);
}
