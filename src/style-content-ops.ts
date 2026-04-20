import type { docs_v1 } from '@googleapis/docs';
import type { BatchUpdateResult, IndexRange } from './docs-client';

type DocsAPI = docs_v1.Docs;

export type { IndexRange };

export interface DocumentStyleInput {
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  pageWidth?: number;
  pageHeight?: number;
  landscape?: boolean;
  useEvenPageHeaderFooter?: boolean;
  background?: { red: number; green: number; blue: number };
}

export interface SectionStyleInput {
  columnCount?: number;
  contentDirection?: 'LEFT_TO_RIGHT' | 'RIGHT_TO_LEFT';
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
}

/**
 * Creates the style and content operations surface for Google Docs.
 * Handles document styling, section management, content deletion, tabs, and person mentions.
 *
 * @param docs - Authenticated Google Docs API instance
 * @returns Object containing all style and content operation methods
 */
export function createStyleContentOps(docs: DocsAPI) {
  return {
    /**
     * Sets document-level page style including margins, page size, orientation, and background color.
     * @param documentId - The ID of the document to update.
     * @param style - Document style properties to apply.
     * @returns A promise resolving to the batch update result.
     */
    updateDocumentStyle: (documentId: string, style: DocumentStyleInput) =>
      updateDocumentStyle(docs, documentId, style),

    /**
     * Inserts a section break at the given index or at the end of the document.
     * @param documentId - The ID of the document.
     * @param type - The section break type. Defaults to 'NEXT_PAGE'.
     * @param index - The insertion index. If omitted, inserts at the end.
     * @returns A promise resolving to the batch update result.
     */
    insertSectionBreak: (
      documentId: string,
      type?: 'NEXT_PAGE' | 'CONTINUOUS',
      index?: number,
    ) => insertSectionBreak(docs, documentId, type, index),

    /**
     * Updates section-level properties such as column count, content direction, and section margins.
     * @param documentId - The ID of the document.
     * @param range - The index range of the section to update.
     * @param style - Section style properties to apply.
     * @returns A promise resolving to the batch update result.
     */
    updateSectionStyle: (documentId: string, range: IndexRange, style: SectionStyleInput) =>
      updateSectionStyle(docs, documentId, range, style),

    /**
     * Deletes content between startIndex and endIndex.
     * @param documentId - The ID of the document.
     * @param range - The index range of content to delete.
     * @returns A promise resolving to the batch update result.
     */
    deleteContentRange: (documentId: string, range: IndexRange) =>
      deleteContentRange(docs, documentId, range),

    /**
     * Removes bullet formatting from paragraphs within the given range.
     * @param documentId - The ID of the document.
     * @param range - The index range from which to remove bullets.
     * @returns A promise resolving to the batch update result.
     */
    deleteParagraphBullets: (documentId: string, range: IndexRange) =>
      deleteParagraphBullets(docs, documentId, range),

    /**
     * Deletes a named range by its ID.
     * @param documentId - The ID of the document.
     * @param namedRangeId - The ID of the named range to delete.
     * @returns A promise resolving to the batch update result.
     */
    deleteNamedRange: (documentId: string, namedRangeId: string) =>
      deleteNamedRange(docs, documentId, namedRangeId),

    /**
     * Adds a new document tab with an optional title.
     * @param documentId - The ID of the document.
     * @param title - Optional title for the new tab.
     * @returns A promise resolving to the batch update result.
     */
    addDocumentTab: (documentId: string, title?: string) =>
      addDocumentTab(docs, documentId, title),

    /**
     * Deletes a document tab by its tab ID.
     * @param documentId - The ID of the document.
     * @param tabId - The ID of the tab to delete.
     * @returns A promise resolving to the batch update result.
     */
    deleteTab: (documentId: string, tabId: string) => deleteTab(docs, documentId, tabId),

    /**
     * Inserts an @mention of a person by email at the given index or at the end.
     * @param documentId - The ID of the document.
     * @param email - The email address of the person to mention.
     * @param index - The insertion index. If omitted, inserts at the end.
     * @returns A promise resolving to the batch update result.
     */
    insertPerson: (documentId: string, email: string, index?: number) =>
      insertPerson(docs, documentId, email, index),
  };
}

async function updateDocumentStyle(
  docs: DocsAPI,
  documentId: string,
  style: DocumentStyleInput,
): Promise<BatchUpdateResult> {
  const documentStyle: docs_v1.Schema$DocumentStyle = {};
  const fields: string[] = [];

  if (style.marginTop !== undefined) {
    documentStyle.marginTop = { magnitude: style.marginTop, unit: 'PT' };
    fields.push('marginTop');
  }
  if (style.marginBottom !== undefined) {
    documentStyle.marginBottom = { magnitude: style.marginBottom, unit: 'PT' };
    fields.push('marginBottom');
  }
  if (style.marginLeft !== undefined) {
    documentStyle.marginLeft = { magnitude: style.marginLeft, unit: 'PT' };
    fields.push('marginLeft');
  }
  if (style.marginRight !== undefined) {
    documentStyle.marginRight = { magnitude: style.marginRight, unit: 'PT' };
    fields.push('marginRight');
  }
  if (style.pageWidth !== undefined || style.pageHeight !== undefined) {
    const pageSize: docs_v1.Schema$Size = {};
    if (style.pageWidth !== undefined) {
      pageSize.width = { magnitude: style.pageWidth, unit: 'PT' };
    }
    if (style.pageHeight !== undefined) {
      pageSize.height = { magnitude: style.pageHeight, unit: 'PT' };
    }
    documentStyle.pageSize = pageSize;
    fields.push('pageSize');
  }
  if (style.landscape !== undefined) {
    // Landscape orientation is controlled via pageSize dimensions.
    // When landscape is true, we set width > height; the API infers orientation from dimensions.
    // If explicit width/height were not provided, use standard US Letter dimensions swapped.
    if (!documentStyle.pageSize) {
      const portrait = { width: 612, height: 792 }; // US Letter in PT
      documentStyle.pageSize = style.landscape
        ? { width: { magnitude: portrait.height, unit: 'PT' }, height: { magnitude: portrait.width, unit: 'PT' } }
        : { width: { magnitude: portrait.width, unit: 'PT' }, height: { magnitude: portrait.height, unit: 'PT' } };
      if (!fields.includes('pageSize')) {
        fields.push('pageSize');
      }
    }
  }
  if (style.useEvenPageHeaderFooter !== undefined) {
    documentStyle.useEvenPageHeaderFooter = style.useEvenPageHeaderFooter;
    fields.push('useEvenPageHeaderFooter');
  }
  if (style.background) {
    documentStyle.background = {
      color: {
        color: {
          rgbColor: {
            red: style.background.red,
            green: style.background.green,
            blue: style.background.blue,
          },
        },
      },
    };
    fields.push('background');
  }

  if (!fields.length) {
    return { documentId, replyCount: 0 };
  }

  return batchUpdate(docs, documentId, [
    {
      updateDocumentStyle: {
        documentStyle,
        fields: fields.join(','),
      },
    },
  ]);
}

async function insertSectionBreak(
  docs: DocsAPI,
  documentId: string,
  type?: 'NEXT_PAGE' | 'CONTINUOUS',
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));
  const sectionType = type ?? 'NEXT_PAGE';

  return batchUpdate(docs, documentId, [
    {
      insertSectionBreak: {
        location: { index: targetIndex },
        sectionType,
      },
    },
  ]);
}

async function updateSectionStyle(
  docs: DocsAPI,
  documentId: string,
  range: IndexRange,
  style: SectionStyleInput,
): Promise<BatchUpdateResult> {
  const sectionStyle: docs_v1.Schema$SectionStyle = {};
  const fields: string[] = [];

  if (style.columnCount !== undefined) {
    sectionStyle.columnProperties = Array.from({ length: style.columnCount }, () => ({}));
    fields.push('columnProperties');
  }
  if (style.contentDirection !== undefined) {
    sectionStyle.contentDirection = style.contentDirection;
    fields.push('contentDirection');
  }
  if (style.marginTop !== undefined) {
    sectionStyle.marginTop = { magnitude: style.marginTop, unit: 'PT' };
    fields.push('marginTop');
  }
  if (style.marginBottom !== undefined) {
    sectionStyle.marginBottom = { magnitude: style.marginBottom, unit: 'PT' };
    fields.push('marginBottom');
  }
  if (style.marginLeft !== undefined) {
    sectionStyle.marginLeft = { magnitude: style.marginLeft, unit: 'PT' };
    fields.push('marginLeft');
  }
  if (style.marginRight !== undefined) {
    sectionStyle.marginRight = { magnitude: style.marginRight, unit: 'PT' };
    fields.push('marginRight');
  }

  if (!fields.length) {
    return { documentId, replyCount: 0 };
  }

  return batchUpdate(docs, documentId, [
    {
      updateSectionStyle: {
        range,
        sectionStyle,
        fields: fields.join(','),
      },
    },
  ]);
}

async function deleteContentRange(
  docs: DocsAPI,
  documentId: string,
  range: IndexRange,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteContentRange: {
        range: {
          startIndex: range.startIndex,
          endIndex: range.endIndex,
        },
      },
    },
  ]);
}

async function deleteParagraphBullets(
  docs: DocsAPI,
  documentId: string,
  range: IndexRange,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteParagraphBullets: {
        range: {
          startIndex: range.startIndex,
          endIndex: range.endIndex,
        },
      },
    },
  ]);
}

async function deleteNamedRange(
  docs: DocsAPI,
  documentId: string,
  namedRangeId: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteNamedRange: {
        namedRangeId,
      },
    },
  ]);
}

async function addDocumentTab(
  docs: DocsAPI,
  documentId: string,
  title?: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      // @ts-expect-error - addDocumentTab is supported by the API but not yet in @googleapis/docs types
      addDocumentTab: {
        ...(title ? { tab: { tabProperties: { title } } } : {}),
      },
    },
  ]);
}

async function deleteTab(
  docs: DocsAPI,
  documentId: string,
  tabId: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      // @ts-expect-error - deleteTab is supported by the API but not yet in @googleapis/docs types
      deleteTab: {
        tabId,
      },
    },
  ]);
}

async function insertPerson(
  docs: DocsAPI,
  documentId: string,
  email: string,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  return batchUpdate(docs, documentId, [
    {
      // @ts-expect-error - insertPerson is supported by the API but not yet in @googleapis/docs types
      insertPerson: {
        location: { index: targetIndex },
        personProperties: {
          email,
        },
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
