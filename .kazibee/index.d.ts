import { docs_v1 } from '@googleapis/docs';

export interface Env {
	CLIENT_ID: string;
	CLIENT_SECRET: string;
	REFRESH_TOKEN: string;
}
export interface ParagraphNode {
	type: "paragraph";
	startIndex: number;
	endIndex: number;
	text: string;
	heading?: string;
	isBullet?: boolean;
}
export interface TableNode {
	type: "table";
	startIndex: number;
	endIndex: number;
	rowCount: number;
	columnCount: number;
}
export interface SectionBreakNode {
	type: "sectionBreak";
	startIndex: number;
	endIndex: number;
}
export type ShadowNode = ParagraphNode | TableNode | SectionBreakNode;
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;
export type PendingChange = {
	type: "insert";
	index: number;
	text: string;
} | {
	type: "delete";
	startIndex: number;
	endIndex: number;
} | {
	type: "replaceAll";
	find: string;
	replace: string;
} | {
	type: "insertHeading";
	index: number;
	text: string;
	level: HeadingLevel;
} | {
	type: "insertBullet";
	index: number;
	items: string[];
};
export declare class ShadowDocument {
	readonly documentId: string;
	readonly title: string;
	private nodes;
	private originalNodes;
	private pendingChanges;
	private constructor();
	/** Build a ShadowDocument from a DocumentStructure returned by getDocumentStructure(). */
	static fromStructure(structure: DocumentStructure): ShadowDocument;
	/** Returns a shallow copy of the current node list. */
	getNodes(): readonly ShadowNode[];
	/** Returns accumulated pending changes in order. */
	getPendingChanges(): readonly PendingChange[];
	/** Clears the pending changes list (call after a successful commit). */
	clearPendingChanges(): void;
	/** Insert text at the given document index. Shifts all downstream indices. */
	insertText(index: number, text: string): void;
	/** Delete a content range. Shifts all downstream indices backward. */
	deleteRange(startIndex: number, endIndex: number): void;
	/** Find and replace all occurrences of `find` with `replace` across paragraph text. */
	replaceText(find: string, replace: string): void;
	/** Insert a heading at the given index. */
	insertHeading(index: number, text: string, level: HeadingLevel): void;
	/** Insert a bullet list at the given index. */
	insertBulletList(index: number, items: string[]): void;
	/** Returns a plain-text representation of the current document state. */
	render(): string;
	/** Returns a simple before/after diff for paragraphs that changed. */
	diff(): string;
	private applyInsert;
	private applyDelete;
	private applyReplaceAll;
	private applyInsertParagraph;
	private applyInsertBulletList;
	/** Find the node index that contains the given document index. Returns -1 if none. */
	private findNodeContaining;
	/** Find the position where a new node should be inserted for the given index. */
	private findInsertPosition;
	/** Shift all nodes after position `afterIdx` by `delta`. */
	private shiftNodesAfter;
}
export interface DocumentStyleInput {
	marginTop?: number;
	marginBottom?: number;
	marginLeft?: number;
	marginRight?: number;
	pageWidth?: number;
	pageHeight?: number;
	landscape?: boolean;
	useEvenPageHeaderFooter?: boolean;
	background?: {
		red: number;
		green: number;
		blue: number;
	};
}
export interface SectionStyleInput {
	columnCount?: number;
	contentDirection?: "LEFT_TO_RIGHT" | "RIGHT_TO_LEFT";
	marginTop?: number;
	marginBottom?: number;
	marginLeft?: number;
	marginRight?: number;
}
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
	contentAlignment?: "TOP" | "MIDDLE" | "BOTTOM";
}
export interface DocumentInfo {
	documentId: string;
	title: string;
	revisionId: string;
}
export interface BatchUpdateResult {
	documentId: string;
	replyCount: number;
}
export interface IndexRange {
	startIndex: number;
	endIndex: number;
}
export interface TextMatchRange {
	startIndex: number;
	endIndex: number;
	text: string;
}
export interface DocumentSection {
	type: "paragraph" | "table" | "sectionBreak";
	startIndex: number;
	endIndex: number;
	text?: string;
	heading?: string;
	rowCount?: number;
	columnCount?: number;
}
export interface DocumentStructure {
	documentId: string;
	title: string;
	sections: DocumentSection[];
}
export interface TextStyleInput {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strikethrough?: boolean;
	linkUrl?: string;
	fontSize?: number;
}
export interface ParagraphStyleInput {
	namedStyleType?: "NORMAL_TEXT" | "TITLE" | "SUBTITLE" | "HEADING_1" | "HEADING_2" | "HEADING_3" | "HEADING_4" | "HEADING_5" | "HEADING_6";
	alignment?: "START" | "CENTER" | "END" | "JUSTIFIED";
	lineSpacing?: number;
}
export interface ImageSize {
	width: number;
	height: number;
}
/**
 * Compiles an array of pending shadow-document changes into Google Docs API
 * batch-update requests. Index-based operations are sorted from highest index
 * to lowest so earlier mutations never shift indices that later operations
 * depend on. `replaceAll` operations are index-independent and are emitted
 * first (order among them does not matter).
 */
export declare function compileRequests(changes: PendingChange[]): docs_v1.Schema$Request[];
declare function main(env: Env): {
	updateDocumentStyle: (documentId: string, style: DocumentStyleInput) => Promise<BatchUpdateResult>;
	insertSectionBreak: (documentId: string, type?: "NEXT_PAGE" | "CONTINUOUS", index?: number) => Promise<BatchUpdateResult>;
	updateSectionStyle: (documentId: string, range: IndexRange, style: SectionStyleInput) => Promise<BatchUpdateResult>;
	deleteContentRange: (documentId: string, range: IndexRange) => Promise<BatchUpdateResult>;
	deleteParagraphBullets: (documentId: string, range: IndexRange) => Promise<BatchUpdateResult>;
	deleteNamedRange: (documentId: string, namedRangeId: string) => Promise<BatchUpdateResult>;
	addDocumentTab: (documentId: string, title?: string) => Promise<BatchUpdateResult>;
	deleteTab: (documentId: string, tabId: string) => Promise<BatchUpdateResult>;
	insertPerson: (documentId: string, email: string, index?: number) => Promise<BatchUpdateResult>;
	insertInlineImage: (documentId: string, uri: string, width?: number, height?: number, index?: number) => Promise<BatchUpdateResult>;
	replaceImage: (documentId: string, imageObjectId: string, uri: string) => Promise<BatchUpdateResult>;
	deletePositionedObject: (documentId: string, objectId: string) => Promise<BatchUpdateResult>;
	createHeader: (documentId: string, type?: string, sectionBreakIndex?: number) => Promise<BatchUpdateResult>;
	createFooter: (documentId: string, type?: string, sectionBreakIndex?: number) => Promise<BatchUpdateResult>;
	deleteHeader: (documentId: string, headerId: string) => Promise<BatchUpdateResult>;
	deleteFooter: (documentId: string, footerId: string) => Promise<BatchUpdateResult>;
	createFootnote: (documentId: string, index?: number) => Promise<BatchUpdateResult>;
	insertTableRow: (documentId: string, cell: TableCellLocation, insertBelow: boolean) => Promise<BatchUpdateResult>;
	insertTableColumn: (documentId: string, cell: TableCellLocation, insertRight: boolean) => Promise<BatchUpdateResult>;
	deleteTableRow: (documentId: string, cell: TableCellLocation) => Promise<BatchUpdateResult>;
	deleteTableColumn: (documentId: string, cell: TableCellLocation) => Promise<BatchUpdateResult>;
	mergeTableCells: (documentId: string, range: TableRange) => Promise<BatchUpdateResult>;
	unmergeTableCells: (documentId: string, range: TableRange) => Promise<BatchUpdateResult>;
	updateTableCellStyle: (documentId: string, range: TableRange, style: TableCellStyleInput) => Promise<BatchUpdateResult>;
	updateTableColumnProperties: (documentId: string, tableStartIndex: number, columnIndices: number[], widthPt: number) => Promise<BatchUpdateResult>;
	updateTableRowStyle: (documentId: string, tableStartIndex: number, rowIndices: number[], minHeightPt: number) => Promise<BatchUpdateResult>;
	pinTableHeaderRows: (documentId: string, tableStartIndex: number, pinnedCount: number) => Promise<BatchUpdateResult>;
	createDocument: (title: string) => Promise<DocumentInfo>;
	getDocument: (documentId: string) => Promise<DocumentInfo>;
	getDocumentText: (documentId: string) => Promise<string>;
	getDocumentStructure: (documentId: string) => Promise<DocumentStructure>;
	appendText: (documentId: string, text: string) => Promise<BatchUpdateResult>;
	insertText: (documentId: string, text: string, index?: number) => Promise<BatchUpdateResult>;
	insertHeading: (documentId: string, text: string, level: 1 | 2 | 3 | 4 | 5 | 6, index?: number) => Promise<BatchUpdateResult>;
	insertBulletList: (documentId: string, items: string[], index?: number) => Promise<BatchUpdateResult>;
	insertTable: (documentId: string, rows: number, columns: number, index?: number) => Promise<BatchUpdateResult>;
	replaceAllText: (documentId: string, findText: string, replaceText: string) => Promise<BatchUpdateResult>;
	replaceAllTextBatch: (documentId: string, replacements: Array<{
		findText: string;
		replaceText: string;
		matchCase?: boolean;
	}>) => Promise<BatchUpdateResult>;
	findTextRanges: (documentId: string, query: string, matchCase?: boolean) => Promise<TextMatchRange[]>;
	applyTextStyle: (documentId: string, range: IndexRange, style: TextStyleInput) => Promise<BatchUpdateResult>;
	applyParagraphStyle: (documentId: string, range: IndexRange, style: ParagraphStyleInput) => Promise<BatchUpdateResult>;
	createNamedRange: (documentId: string, name: string, range: IndexRange) => Promise<BatchUpdateResult>;
	replaceNamedRange: (documentId: string, name: string, text: string) => Promise<BatchUpdateResult>;
	insertPageBreak: (documentId: string, index?: number) => Promise<BatchUpdateResult>;
	batchUpdate: (documentId: string, requests: import("@googleapis/docs").docs_v1.Schema$Request[]) => Promise<BatchUpdateResult>;
	safeBatchUpdate: (documentId: string, requests: import("@googleapis/docs").docs_v1.Schema$Request[]) => Promise<BatchUpdateResult>;
	openDocument: (documentId: string) => Promise<ShadowDocument>;
	previewChanges: (shadow: ShadowDocument) => string;
	diffChanges: (shadow: ShadowDocument) => string;
	commitChanges: (shadow: ShadowDocument) => Promise<BatchUpdateResult>;
};

export {
	main as default,
};

export {};
