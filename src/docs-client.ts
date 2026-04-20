import { docs as createDocs, type docs_v1 } from '@googleapis/docs';
import type { OAuth2Client } from 'google-auth-library';
import { createTableOps } from './table-ops';
import { createMediaOps } from './media-ops';
import { createStyleContentOps } from './style-content-ops';
import { ShadowDocument } from './shadow-document';
import { compileRequests } from './request-compiler';

type DocsAPI = docs_v1.Docs;

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
  type: 'paragraph' | 'table' | 'sectionBreak';
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
  namedStyleType?:
    | 'NORMAL_TEXT'
    | 'TITLE'
    | 'SUBTITLE'
    | 'HEADING_1'
    | 'HEADING_2'
    | 'HEADING_3'
    | 'HEADING_4'
    | 'HEADING_5'
    | 'HEADING_6';
  alignment?: 'START' | 'CENTER' | 'END' | 'JUSTIFIED';
  lineSpacing?: number;
}

export function createDocsClient(auth: OAuth2Client) {
  const docs = createDocs({ version: 'v1', auth });

  return {
    /**
     * Creates a new Google Doc.
     */
    createDocument: (title: string) => createDocument(docs, title),

    /**
     * Returns basic metadata for a Google Doc.
     */
    getDocument: (documentId: string) => getDocument(docs, documentId),

    /**
     * Returns flattened plain text for a Google Doc.
     */
    getDocumentText: (documentId: string) => getDocumentText(docs, documentId),

    /**
     * Returns a simplified document structure (paragraphs, headings, tables).
     */
    getDocumentStructure: (documentId: string) => getDocumentStructure(docs, documentId),

    /**
     * Appends text at the end of the document.
     */
    appendText: (documentId: string, text: string) => appendText(docs, documentId, text),

    /**
     * Inserts text at a specific index. If no index is provided, inserts at end.
     */
    insertText: (documentId: string, text: string, index?: number) =>
      insertText(docs, documentId, text, index),

    /**
     * Inserts heading text and applies heading style.
     */
    insertHeading: (
      documentId: string,
      text: string,
      level: 1 | 2 | 3 | 4 | 5 | 6,
      index?: number,
    ) => insertHeading(docs, documentId, text, level, index),

    /**
     * Inserts a bullet list for the provided items.
     */
    insertBulletList: (documentId: string, items: string[], index?: number) =>
      insertBulletList(docs, documentId, items, index),

    /**
     * Inserts an empty table at the target index or at the end.
     */
    insertTable: (documentId: string, rows: number, columns: number, index?: number) =>
      insertTable(docs, documentId, rows, columns, index),

    /**
     * Replaces all exact occurrences of a string in the document.
     */
    replaceAllText: (documentId: string, findText: string, replaceText: string) =>
      replaceAllText(docs, documentId, findText, replaceText),

    /**
     * Runs multiple replace-all operations in one batch update.
     */
    replaceAllTextBatch: (
      documentId: string,
      replacements: Array<{ findText: string; replaceText: string; matchCase?: boolean }>,
    ) => replaceAllTextBatch(docs, documentId, replacements),

    /**
     * Finds all text matches and returns their document index ranges.
     */
    findTextRanges: (documentId: string, query: string, matchCase?: boolean) =>
      findTextRanges(docs, documentId, query, matchCase),

    /**
     * Applies text style to a specific range.
     */
    applyTextStyle: (documentId: string, range: IndexRange, style: TextStyleInput) =>
      applyTextStyle(docs, documentId, range, style),

    /**
     * Applies paragraph style to a specific range.
     */
    applyParagraphStyle: (documentId: string, range: IndexRange, style: ParagraphStyleInput) =>
      applyParagraphStyle(docs, documentId, range, style),

    /**
     * Creates a named range for future targeted edits.
     */
    createNamedRange: (documentId: string, name: string, range: IndexRange) =>
      createNamedRange(docs, documentId, name, range),

    /**
     * Replaces content inside one named range by name.
     */
    replaceNamedRange: (documentId: string, name: string, text: string) =>
      replaceNamedRange(docs, documentId, name, text),

    /**
     * Inserts a page break at a specific index or at the end.
     */
    insertPageBreak: (documentId: string, index?: number) =>
      insertPageBreak(docs, documentId, index),

    /**
     * Runs raw Docs API requests in a single batch.
     */
    batchUpdate: (documentId: string, requests: docs_v1.Schema$Request[]) =>
      batchUpdate(docs, documentId, requests),

    /**
     * Runs raw Docs API requests and retries once with normalized indexes if needed.
     */
    safeBatchUpdate: (documentId: string, requests: docs_v1.Schema$Request[]) =>
      safeBatchUpdate(docs, documentId, requests),

    /**
     * Opens a document and returns a mutable ShadowDocument for offline editing.
     */
    openDocument: async (documentId: string): Promise<ShadowDocument> => {
      const structure = await getDocumentStructure(docs, documentId);
      return ShadowDocument.fromStructure(structure);
    },

    /**
     * Returns a plain-text preview of the shadow document's current state.
     */
    previewChanges: (shadow: ShadowDocument): string => shadow.render(),

    /**
     * Returns a before/after diff of pending changes in the shadow document.
     */
    diffChanges: (shadow: ShadowDocument): string => shadow.diff(),

    /**
     * Compiles pending shadow changes into API requests, sends them, then
     * clears the pending changes. Callers should use openDocument() again
     * to obtain a fresh shadow that reflects the remote state.
     */
    commitChanges: async (shadow: ShadowDocument): Promise<BatchUpdateResult> => {
      const requests = compileRequests([...shadow.getPendingChanges()]);
      const result = await batchUpdate(docs, shadow.documentId, requests);
      shadow.clearPendingChanges();
      return result;
    },

    ...createTableOps(docs),
    ...createMediaOps(docs),
    ...createStyleContentOps(docs),
  };
}

async function createDocument(docs: DocsAPI, title: string): Promise<DocumentInfo> {
  const res = await docs.documents.create({ requestBody: { title } });
  return mapDocument(res.data);
}

async function getDocument(docs: DocsAPI, documentId: string): Promise<DocumentInfo> {
  const res = await docs.documents.get({ documentId });
  return mapDocument(res.data);
}

async function getDocumentText(docs: DocsAPI, documentId: string): Promise<string> {
  const res = await docs.documents.get({ documentId });
  return extractText(res.data.body?.content ?? []);
}

async function getDocumentStructure(docs: DocsAPI, documentId: string): Promise<DocumentStructure> {
  const res = await docs.documents.get({ documentId });
  const doc = res.data;
  const content = doc.body?.content ?? [];

  const sections: DocumentSection[] = content.map((element) => {
    if (element.table) {
      const rows = element.table.tableRows ?? [];
      const rowCount = rows.length;
      const columnCount = rows[0]?.tableCells?.length ?? 0;
      return {
        type: 'table',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex ?? 0,
        rowCount,
        columnCount,
      };
    }

    if (element.sectionBreak) {
      return {
        type: 'sectionBreak',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex ?? 0,
      };
    }

    const paragraph = element.paragraph;
    const paragraphText = paragraph?.elements?.map((p) => p.textRun?.content ?? '').join('') ?? '';

    return {
      type: 'paragraph',
      startIndex: element.startIndex ?? 0,
      endIndex: element.endIndex ?? 0,
      text: paragraphText,
      heading: paragraph?.paragraphStyle?.namedStyleType ?? undefined,
    };
  });

  return {
    documentId: doc.documentId ?? '',
    title: doc.title ?? '',
    sections,
  };
}

async function appendText(docs: DocsAPI, documentId: string, text: string): Promise<BatchUpdateResult> {
  return insertText(docs, documentId, text);
}

async function insertText(
  docs: DocsAPI,
  documentId: string,
  text: string,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  return batchUpdate(docs, documentId, [{ insertText: { location: { index: targetIndex }, text } }]);
}

async function insertHeading(
  docs: DocsAPI,
  documentId: string,
  text: string,
  level: 1 | 2 | 3 | 4 | 5 | 6,
  index?: number,
): Promise<BatchUpdateResult> {
  const headingText = text.endsWith('\n') ? text : `${text}\n`;
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));
  const endIndex = targetIndex + headingText.length;

  return batchUpdate(docs, documentId, [
    { insertText: { location: { index: targetIndex }, text: headingText } },
    {
      updateParagraphStyle: {
        range: { startIndex: targetIndex, endIndex },
        paragraphStyle: { namedStyleType: `HEADING_${level}` },
        fields: 'namedStyleType',
      },
    },
  ]);
}

async function insertBulletList(
  docs: DocsAPI,
  documentId: string,
  items: string[],
  index?: number,
): Promise<BatchUpdateResult> {
  const listText = `${items.join('\n')}\n`;
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));
  const endIndex = targetIndex + listText.length;

  return batchUpdate(docs, documentId, [
    { insertText: { location: { index: targetIndex }, text: listText } },
    {
      createParagraphBullets: {
        range: { startIndex: targetIndex, endIndex },
        bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
      },
    },
  ]);
}

async function insertTable(
  docs: DocsAPI,
  documentId: string,
  rows: number,
  columns: number,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  return batchUpdate(docs, documentId, [
    { insertTable: { rows, columns, location: { index: targetIndex } } },
  ]);
}

async function replaceAllText(
  docs: DocsAPI,
  documentId: string,
  findText: string,
  replaceText: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      replaceAllText: {
        containsText: {
          text: findText,
          matchCase: true,
        },
        replaceText,
      },
    },
  ]);
}

async function replaceAllTextBatch(
  docs: DocsAPI,
  documentId: string,
  replacements: Array<{ findText: string; replaceText: string; matchCase?: boolean }>,
): Promise<BatchUpdateResult> {
  const requests: docs_v1.Schema$Request[] = replacements.map((entry) => ({
    replaceAllText: {
      containsText: {
        text: entry.findText,
        matchCase: entry.matchCase ?? true,
      },
      replaceText: entry.replaceText,
    },
  }));

  return batchUpdate(docs, documentId, requests);
}

async function findTextRanges(
  docs: DocsAPI,
  documentId: string,
  query: string,
  matchCase = true,
): Promise<TextMatchRange[]> {
  const res = await docs.documents.get({ documentId });
  const { fullText, indexMap } = buildTextIndexMap(res.data.body?.content ?? []);
  if (!query) return [];

  const haystack = matchCase ? fullText : fullText.toLowerCase();
  const needle = matchCase ? query : query.toLowerCase();
  const matches: TextMatchRange[] = [];

  let from = 0;
  while (from < haystack.length) {
    const at = haystack.indexOf(needle, from);
    if (at === -1) break;

    const startIndex = indexMap[at];
    const endMapIndex = at + needle.length - 1;
    const endIndex = indexMap[endMapIndex] + 1;

    matches.push({
      startIndex,
      endIndex,
      text: fullText.slice(at, at + needle.length),
    });

    from = at + needle.length;
  }

  return matches;
}

async function applyTextStyle(
  docs: DocsAPI,
  documentId: string,
  range: IndexRange,
  style: TextStyleInput,
): Promise<BatchUpdateResult> {
  const textStyle: docs_v1.Schema$TextStyle = {};
  const fields: string[] = [];

  if (style.bold !== undefined) {
    textStyle.bold = style.bold;
    fields.push('bold');
  }
  if (style.italic !== undefined) {
    textStyle.italic = style.italic;
    fields.push('italic');
  }
  if (style.underline !== undefined) {
    textStyle.underline = style.underline;
    fields.push('underline');
  }
  if (style.strikethrough !== undefined) {
    textStyle.strikethrough = style.strikethrough;
    fields.push('strikethrough');
  }
  if (style.linkUrl) {
    textStyle.link = { url: style.linkUrl };
    fields.push('link');
  }
  if (style.fontSize !== undefined) {
    textStyle.weightedFontFamily = undefined;
    textStyle.fontSize = { magnitude: style.fontSize, unit: 'PT' };
    fields.push('fontSize');
  }

  if (!fields.length) {
    return { documentId, replyCount: 0 };
  }

  return batchUpdate(docs, documentId, [
    {
      updateTextStyle: {
        range,
        textStyle,
        fields: fields.join(','),
      },
    },
  ]);
}

async function applyParagraphStyle(
  docs: DocsAPI,
  documentId: string,
  range: IndexRange,
  style: ParagraphStyleInput,
): Promise<BatchUpdateResult> {
  const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
  const fields: string[] = [];

  if (style.namedStyleType) {
    paragraphStyle.namedStyleType = style.namedStyleType;
    fields.push('namedStyleType');
  }
  if (style.alignment) {
    paragraphStyle.alignment = style.alignment;
    fields.push('alignment');
  }
  if (style.lineSpacing !== undefined) {
    paragraphStyle.lineSpacing = style.lineSpacing;
    fields.push('lineSpacing');
  }

  if (!fields.length) {
    return { documentId, replyCount: 0 };
  }

  return batchUpdate(docs, documentId, [
    {
      updateParagraphStyle: {
        range,
        paragraphStyle,
        fields: fields.join(','),
      },
    },
  ]);
}

async function createNamedRange(
  docs: DocsAPI,
  documentId: string,
  name: string,
  range: IndexRange,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      createNamedRange: {
        name,
        range,
      },
    },
  ]);
}

async function replaceNamedRange(
  docs: DocsAPI,
  documentId: string,
  name: string,
  text: string,
): Promise<BatchUpdateResult> {
  const doc = await docs.documents.get({ documentId });
  const ranges = Object.values(doc.data.namedRanges ?? {}).flatMap((entry) =>
    (entry.namedRanges ?? []).filter((named) => named.name === name),
  );

  const requests: docs_v1.Schema$Request[] = ranges.flatMap((named) => {
    const allRanges = named.ranges ?? [];
    return allRanges
      .filter((range): range is docs_v1.Schema$Range =>
        typeof range.startIndex === 'number' && typeof range.endIndex === 'number',
      )
      .sort((a, b) => (b.startIndex ?? 0) - (a.startIndex ?? 0))
      .flatMap((range) => {
        const startIndex = range.startIndex as number;
        const endIndex = range.endIndex as number;
        return [
          { deleteContentRange: { range: { startIndex, endIndex } } },
          { insertText: { location: { index: startIndex }, text } },
        ];
      });
  });

  return batchUpdate(docs, documentId, requests);
}

async function insertPageBreak(
  docs: DocsAPI,
  documentId: string,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  return batchUpdate(docs, documentId, [
    {
      insertPageBreak: {
        location: { index: targetIndex },
      },
    },
  ]);
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

async function safeBatchUpdate(
  docs: DocsAPI,
  documentId: string,
  requests: docs_v1.Schema$Request[],
): Promise<BatchUpdateResult> {
  try {
    return await batchUpdate(docs, documentId, requests);
  } catch {
    const endIndex = await getDocumentEndIndex(docs, documentId);
    const normalized = normalizeRequests(requests, endIndex);
    return batchUpdate(docs, documentId, normalized);
  }
}

function normalizeRequests(
  requests: docs_v1.Schema$Request[],
  endIndex: number,
): docs_v1.Schema$Request[] {
  const clamp = (value: number) => Math.max(1, Math.min(value, endIndex));

  return requests.map((req) => {
    const next: docs_v1.Schema$Request = { ...req };

    if (next.insertText?.location?.index !== undefined) {
      next.insertText = {
        ...next.insertText,
        location: { index: clamp(next.insertText.location.index ?? 1) },
      };
    }

    if (next.insertPageBreak?.location?.index !== undefined) {
      next.insertPageBreak = {
        ...next.insertPageBreak,
        location: { index: clamp(next.insertPageBreak.location.index ?? 1) },
      };
    }

    if (next.deleteContentRange?.range) {
      const start = clamp(next.deleteContentRange.range.startIndex ?? 1);
      const end = clamp(next.deleteContentRange.range.endIndex ?? endIndex);
      next.deleteContentRange = {
        ...next.deleteContentRange,
        range: { startIndex: Math.min(start, end), endIndex: Math.max(start, end) },
      };
    }

    if (next.updateTextStyle?.range) {
      const start = clamp(next.updateTextStyle.range.startIndex ?? 1);
      const end = clamp(next.updateTextStyle.range.endIndex ?? endIndex);
      next.updateTextStyle = {
        ...next.updateTextStyle,
        range: { startIndex: Math.min(start, end), endIndex: Math.max(start, end) },
      };
    }

    if (next.updateParagraphStyle?.range) {
      const start = clamp(next.updateParagraphStyle.range.startIndex ?? 1);
      const end = clamp(next.updateParagraphStyle.range.endIndex ?? endIndex);
      next.updateParagraphStyle = {
        ...next.updateParagraphStyle,
        range: { startIndex: Math.min(start, end), endIndex: Math.max(start, end) },
      };
    }

    if (next.createParagraphBullets?.range) {
      const start = clamp(next.createParagraphBullets.range.startIndex ?? 1);
      const end = clamp(next.createParagraphBullets.range.endIndex ?? endIndex);
      next.createParagraphBullets = {
        ...next.createParagraphBullets,
        range: { startIndex: Math.min(start, end), endIndex: Math.max(start, end) },
      };
    }

    if (next.insertTable?.location?.index !== undefined) {
      next.insertTable = {
        ...next.insertTable,
        location: { index: clamp(next.insertTable.location.index ?? 1) },
      };
    }

    return next;
  });
}

async function getDocumentEndIndex(docs: DocsAPI, documentId: string): Promise<number> {
  const res = await docs.documents.get({ documentId });
  const content = res.data.body?.content ?? [];

  for (let i = content.length - 1; i >= 0; i -= 1) {
    const endIndex = content[i]?.endIndex;
    if (typeof endIndex === 'number') {
      return Math.max(1, endIndex - 1);
    }
  }

  return 1;
}

function mapDocument(data: docs_v1.Schema$Document): DocumentInfo {
  return {
    documentId: data.documentId ?? '',
    title: data.title ?? '',
    revisionId: data.revisionId ?? '',
  };
}

function extractText(content: docs_v1.Schema$StructuralElement[]): string {
  const parts: string[] = [];

  for (const element of content) {
    const paragraph = element.paragraph;
    if (!paragraph?.elements) continue;

    for (const paragraphElement of paragraph.elements) {
      const text = paragraphElement.textRun?.content;
      if (text) parts.push(text);
    }
  }

  return parts.join('');
}

function buildTextIndexMap(content: docs_v1.Schema$StructuralElement[]): {
  fullText: string;
  indexMap: number[];
} {
  let fullText = '';
  const indexMap: number[] = [];

  for (const element of content) {
    const paragraph = element.paragraph;
    if (!paragraph?.elements) continue;

    for (const paragraphElement of paragraph.elements) {
      const runText = paragraphElement.textRun?.content;
      const startIndex = paragraphElement.startIndex;
      if (!runText || typeof startIndex !== 'number') continue;

      fullText += runText;
      for (let i = 0; i < runText.length; i += 1) {
        indexMap.push(startIndex + i);
      }
    }
  }

  return { fullText, indexMap };
}
