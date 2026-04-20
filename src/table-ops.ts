import type { docs_v1 } from '@googleapis/docs';
import type { BatchUpdateResult } from './docs-client';

type DocsAPI = docs_v1.Docs;

/** Identifies a single cell within a table by row, column, and the table's start index. */
export interface TableCellLocation {
  /** Zero-based row index within the table. */
  rowIndex: number;
  /** Zero-based column index within the table. */
  columnIndex: number;
  /** The document body index where the table element begins. */
  tableStartIndex: number;
}

/** Describes a rectangular block of cells within a table. */
export interface TableRange {
  /** The upper-left cell of the range. */
  location: TableCellLocation;
  /** Number of rows the range spans. */
  rowSpan: number;
  /** Number of columns the range spans. */
  columnSpan: number;
}

export interface TableCellStyleInput {
  /** Background color in hex format (e.g. "#FF0000"). */
  backgroundColor?: string;
  /** Padding for all four sides, in PT. */
  paddingTop?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  paddingRight?: number;
  /** Content alignment within the cell. */
  contentAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
}

export function createTableOps(docs: DocsAPI) {
  return {
    /**
     * Inserts an empty row into a table.
     * @param documentId - The ID of the document containing the table.
     * @param cell - A cell in the row adjacent to where the new row will be inserted.
     * @param insertBelow - If true, insert below the reference row; otherwise above.
     * @returns A promise resolving to the batch update result.
     */
    insertTableRow: (documentId: string, cell: TableCellLocation, insertBelow: boolean) =>
      insertTableRow(docs, documentId, cell, insertBelow),

    /**
     * Inserts an empty column into a table.
     * @param documentId - The ID of the document containing the table.
     * @param cell - A cell in the column adjacent to where the new column will be inserted.
     * @param insertRight - If true, insert to the right of the reference column; otherwise to the left.
     * @returns A promise resolving to the batch update result.
     */
    insertTableColumn: (documentId: string, cell: TableCellLocation, insertRight: boolean) =>
      insertTableColumn(docs, documentId, cell, insertRight),

    /**
     * Deletes an entire row from a table.
     * @param documentId - The ID of the document containing the table.
     * @param cell - A cell in the row to delete.
     * @returns A promise resolving to the batch update result.
     */
    deleteTableRow: (documentId: string, cell: TableCellLocation) =>
      deleteTableRow(docs, documentId, cell),

    /**
     * Deletes an entire column from a table.
     * @param documentId - The ID of the document containing the table.
     * @param cell - A cell in the column to delete.
     * @returns A promise resolving to the batch update result.
     */
    deleteTableColumn: (documentId: string, cell: TableCellLocation) =>
      deleteTableColumn(docs, documentId, cell),

    /**
     * Merges a rectangular range of table cells into one.
     * @param documentId - The ID of the document containing the table.
     * @param range - The rectangular range of cells to merge.
     * @returns A promise resolving to the batch update result.
     */
    mergeTableCells: (documentId: string, range: TableRange) =>
      mergeTableCells(docs, documentId, range),

    /**
     * Unmerges previously merged cells back into individual cells.
     * @param documentId - The ID of the document containing the table.
     * @param range - The rectangular range covering the merged cells to split.
     * @returns A promise resolving to the batch update result.
     */
    unmergeTableCells: (documentId: string, range: TableRange) =>
      unmergeTableCells(docs, documentId, range),

    /**
     * Updates the visual style of table cells (background color, padding, alignment).
     * @param documentId - The ID of the document containing the table.
     * @param range - The rectangular range of cells to style.
     * @param style - The style properties to apply.
     * @returns A promise resolving to the batch update result.
     */
    updateTableCellStyle: (documentId: string, range: TableRange, style: TableCellStyleInput) =>
      updateTableCellStyle(docs, documentId, range, style),

    /**
     * Updates column properties such as width for specified columns.
     * @param documentId - The ID of the document containing the table.
     * @param tableStartIndex - The document index where the table begins.
     * @param columnIndices - Zero-based column indices to update.
     * @param widthPt - The column width in points.
     * @returns A promise resolving to the batch update result.
     */
    updateTableColumnProperties: (
      documentId: string,
      tableStartIndex: number,
      columnIndices: number[],
      widthPt: number,
    ) => updateTableColumnProperties(docs, documentId, tableStartIndex, columnIndices, widthPt),

    /**
     * Updates the row style (minimum height) for specified rows.
     * @param documentId - The ID of the document containing the table.
     * @param tableStartIndex - The document index where the table begins.
     * @param rowIndices - Zero-based row indices to update.
     * @param minHeightPt - The minimum row height in points.
     * @returns A promise resolving to the batch update result.
     */
    updateTableRowStyle: (
      documentId: string,
      tableStartIndex: number,
      rowIndices: number[],
      minHeightPt: number,
    ) => updateTableRowStyle(docs, documentId, tableStartIndex, rowIndices, minHeightPt),

    /**
     * Pins N header rows at the top of a table so they repeat on every page.
     * @param documentId - The ID of the document containing the table.
     * @param tableStartIndex - The document index where the table begins.
     * @param pinnedCount - Number of rows to pin. Pass 0 to unpin all.
     * @returns A promise resolving to the batch update result.
     */
    pinTableHeaderRows: (documentId: string, tableStartIndex: number, pinnedCount: number) =>
      pinTableHeaderRows(docs, documentId, tableStartIndex, pinnedCount),
  };
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

function toApiCellLocation(cell: TableCellLocation): docs_v1.Schema$TableCellLocation {
  return {
    rowIndex: cell.rowIndex,
    columnIndex: cell.columnIndex,
    tableStartLocation: { index: cell.tableStartIndex },
  };
}

function toApiTableRange(range: TableRange): docs_v1.Schema$TableRange {
  return {
    tableCellLocation: toApiCellLocation(range.location),
    rowSpan: range.rowSpan,
    columnSpan: range.columnSpan,
  };
}

function hexToColor(hex: string): docs_v1.Schema$OptionalColor {
  const cleaned = hex.replace('#', '');
  const r = parseInt(cleaned.substring(0, 2), 16) / 255;
  const g = parseInt(cleaned.substring(2, 4), 16) / 255;
  const b = parseInt(cleaned.substring(4, 6), 16) / 255;
  return { color: { rgbColor: { red: r, green: g, blue: b } } };
}

async function batchUpdate(
  docs: DocsAPI,
  documentId: string,
  requests: docs_v1.Schema$Request[],
): Promise<BatchUpdateResult> {
  const res = await docs.documents.batchUpdate({
    documentId,
    requestBody: { requests },
  });

  return {
    documentId,
    replyCount: res.data.replies?.length ?? 0,
  };
}

// ---------------------------------------------------------------------------
// Table row / column insertion & deletion
// ---------------------------------------------------------------------------

async function insertTableRow(
  docs: DocsAPI,
  documentId: string,
  cell: TableCellLocation,
  insertBelow: boolean,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      insertTableRow: {
        tableCellLocation: toApiCellLocation(cell),
        insertBelow,
      },
    },
  ]);
}

async function insertTableColumn(
  docs: DocsAPI,
  documentId: string,
  cell: TableCellLocation,
  insertRight: boolean,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      insertTableColumn: {
        tableCellLocation: toApiCellLocation(cell),
        insertRight,
      },
    },
  ]);
}

async function deleteTableRow(
  docs: DocsAPI,
  documentId: string,
  cell: TableCellLocation,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteTableRow: {
        tableCellLocation: toApiCellLocation(cell),
      },
    },
  ]);
}

async function deleteTableColumn(
  docs: DocsAPI,
  documentId: string,
  cell: TableCellLocation,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteTableColumn: {
        tableCellLocation: toApiCellLocation(cell),
      },
    },
  ]);
}

// ---------------------------------------------------------------------------
// Cell merging / unmerging
// ---------------------------------------------------------------------------

async function mergeTableCells(
  docs: DocsAPI,
  documentId: string,
  range: TableRange,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      mergeTableCells: {
        tableRange: toApiTableRange(range),
      },
    },
  ]);
}

async function unmergeTableCells(
  docs: DocsAPI,
  documentId: string,
  range: TableRange,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      unmergeTableCells: {
        tableRange: toApiTableRange(range),
      },
    },
  ]);
}

// ---------------------------------------------------------------------------
// Cell / column / row style updates
// ---------------------------------------------------------------------------

async function updateTableCellStyle(
  docs: DocsAPI,
  documentId: string,
  range: TableRange,
  style: TableCellStyleInput,
): Promise<BatchUpdateResult> {
  const tableCellStyle: docs_v1.Schema$TableCellStyle = {};
  const fields: string[] = [];

  if (style.backgroundColor !== undefined) {
    tableCellStyle.backgroundColor = hexToColor(style.backgroundColor);
    fields.push('backgroundColor');
  }
  if (style.paddingTop !== undefined) {
    tableCellStyle.paddingTop = { magnitude: style.paddingTop, unit: 'PT' };
    fields.push('paddingTop');
  }
  if (style.paddingBottom !== undefined) {
    tableCellStyle.paddingBottom = { magnitude: style.paddingBottom, unit: 'PT' };
    fields.push('paddingBottom');
  }
  if (style.paddingLeft !== undefined) {
    tableCellStyle.paddingLeft = { magnitude: style.paddingLeft, unit: 'PT' };
    fields.push('paddingLeft');
  }
  if (style.paddingRight !== undefined) {
    tableCellStyle.paddingRight = { magnitude: style.paddingRight, unit: 'PT' };
    fields.push('paddingRight');
  }
  if (style.contentAlignment !== undefined) {
    tableCellStyle.contentAlignment = style.contentAlignment;
    fields.push('contentAlignment');
  }

  if (!fields.length) {
    return { documentId, replyCount: 0 };
  }

  return batchUpdate(docs, documentId, [
    {
      updateTableCellStyle: {
        tableRange: toApiTableRange(range),
        tableCellStyle,
        fields: fields.join(','),
      },
    },
  ]);
}

async function updateTableColumnProperties(
  docs: DocsAPI,
  documentId: string,
  tableStartIndex: number,
  columnIndices: number[],
  widthPt: number,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      updateTableColumnProperties: {
        tableStartLocation: { index: tableStartIndex },
        columnIndices,
        tableColumnProperties: {
          width: { magnitude: widthPt, unit: 'PT' },
          widthType: 'FIXED_WIDTH',
        },
        fields: 'width,widthType',
      },
    },
  ]);
}

async function updateTableRowStyle(
  docs: DocsAPI,
  documentId: string,
  tableStartIndex: number,
  rowIndices: number[],
  minHeightPt: number,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      updateTableRowStyle: {
        tableStartLocation: { index: tableStartIndex },
        rowIndices,
        tableRowStyle: {
          minRowHeight: { magnitude: minHeightPt, unit: 'PT' },
        },
        fields: 'minRowHeight',
      },
    },
  ]);
}

// ---------------------------------------------------------------------------
// Header row pinning
// ---------------------------------------------------------------------------

async function pinTableHeaderRows(
  docs: DocsAPI,
  documentId: string,
  tableStartIndex: number,
  pinnedCount: number,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      pinTableHeaderRows: {
        tableStartLocation: { index: tableStartIndex },
        pinnedHeaderRowsCount: pinnedCount,
      },
    },
  ]);
}
