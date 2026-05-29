/* global Excel */

/**
 * Turn a pandas dtype string into an Excel number-format string.
 * Datetime columns are handled separately (values are converted to Excel
 * serials), so here we only need the display format for them.
 */
function numberFormatForDtype(dtype) {
  const d = (dtype || "").toLowerCase();
  if (d.startsWith("datetime")) return "yyyy-mm-dd hh:mm";
  if (d.startsWith("timedelta")) return "[h]:mm:ss";
  if (d.startsWith("int") || d.startsWith("uint")) return "#,##0";
  if (d.startsWith("float")) return "#,##0.00";
  if (d === "bool" || d.startsWith("bool")) return "General";
  return "General"; // object / string / category
}

/** Excel stores dates as serial numbers: days since 1899-12-30. */
function isoToExcelSerial(iso) {
  const ms = Date.parse(iso);
  if (Number.isNaN(ms)) return iso; // leave unparseable strings as-is
  return ms / 86400000 + 25569; // 25569 = days between 1899-12-30 and 1970-01-01
}

/**
 * Coerce the JSON rows coming from the kernel into values Excel accepts,
 * converting datetime columns to serials and nulls to blanks.
 */
function sanitizeRows(rows, dtypes) {
  const isDate = dtypes.map((t) => (t || "").toLowerCase().startsWith("datetime"));
  return rows.map((row) =>
    row.map((cell, c) => {
      if (cell === null || cell === undefined) return "";
      if (isDate[c]) return isoToExcelSerial(cell);
      return cell;
    })
  );
}

/**
 * Write a DataFrame payload to the workbook and format it.
 *
 * @param {object} payload  - {columns, dtypes, rows, target?}
 * @param {object} [override] - {anchor, sheet, table} to override payload.target
 * @returns {Promise<{address: string, rows: number, cols: number}>}
 */
export async function writeDataFrame(payload, override = {}) {
  const { columns, dtypes, rows } = payload;
  const target = { ...(payload.target || {}), ...override };

  const anchor = target.anchor || "A1";
  const sheetName = target.sheet || null;
  const asTable = target.table !== false;

  const nCols = columns.length;
  const nRows = rows.length;
  const body = sanitizeRows(rows, dtypes);

  return Excel.run(async (ctx) => {
    // --- resolve the worksheet -----------------------------------------
    let sheet;
    if (sheetName) {
      sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      if (sheet.isNullObject) sheet = ctx.workbook.worksheets.add(sheetName);
    } else {
      sheet = ctx.workbook.worksheets.getActiveWorksheet();
    }
    sheet.activate();

    // --- write header + body in one shot -------------------------------
    const start = sheet.getRange(anchor);
    const fullRange = start.getResizedRange(nRows, nCols - 1); // +1 header row included via nRows? see below
    // start is 1x1. Total rows we need = 1 header + nRows. So delta = nRows.
    // (getResizedRange(deltaRows, deltaCols) keeps the anchor and grows.)
    fullRange.values = [columns, ...body];

    // --- per-column number formats on the data body --------------------
    if (nRows > 0) {
      const dataRange = start.getOffsetRange(1, 0).getResizedRange(nRows - 1, nCols - 1);
      const colFormats = dtypes.map(numberFormatForDtype);
      const fmtGrid = Array.from({ length: nRows }, () => colFormats.slice());
      dataRange.numberFormat = fmtGrid;
    }

    // --- styling: Table, or manual header treatment --------------------
    if (asTable) {
      const table = sheet.tables.add(fullRange, true /* hasHeaders */);
      table.style = "TableStyleMedium2";
      table.getHeaderRowRange().format.font.bold = true;
    } else {
      const header = start.getResizedRange(0, nCols - 1);
      header.format.font.bold = true;
      header.format.font.color = "#0B0B0C";
      header.format.fill.color = "#E8B500";
      header.format.horizontalAlignment = "Left";
      fullRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
      fullRange.format.borders.getItem("InsideHorizontal").color = "#E5E5E5";
    }

    fullRange.format.autofitColumns();
    fullRange.select();

    fullRange.load("address");
    await ctx.sync();

    return { address: fullRange.address, rows: nRows, cols: nCols };
  });
}

/** Read the user's current selection back as a 2D array (handy for round-tripping). */
export async function readSelection() {
  return Excel.run(async (ctx) => {
    const rng = ctx.workbook.getSelectedRange();
    rng.load(["values", "address"]);
    await ctx.sync();
    return { values: rng.values, address: rng.address };
  });
}
