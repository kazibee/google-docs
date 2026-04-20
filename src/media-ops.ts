import type { docs_v1 } from '@googleapis/docs';
import type { BatchUpdateResult } from './docs-client';

type DocsAPI = docs_v1.Docs;

export interface ImageSize {
  width: number;
  height: number;
}

/**
 * Creates the media operations surface for Google Docs.
 * Handles images, headers, footers, and footnotes.
 *
 * @param docs - Authenticated Google Docs API instance
 * @returns Object containing all media operation methods
 */
export function createMediaOps(docs: DocsAPI) {
  return {
    /**
     * Inserts an inline image from a public URL at the given index or at the end of the document.
     *
     * @param documentId - The ID of the target document
     * @param uri - Public URL of the image to insert
     * @param width - Optional image width in PT units
     * @param height - Optional image height in PT units
     * @param index - Optional insertion index (defaults to end of document)
     * @returns The batch update result with document ID and reply count
     */
    insertInlineImage: (
      documentId: string,
      uri: string,
      width?: number,
      height?: number,
      index?: number,
    ) => insertInlineImage(docs, documentId, uri, width, height, index),

    /**
     * Replaces an existing image by its object ID with a new image from a public URL.
     *
     * @param documentId - The ID of the target document
     * @param imageObjectId - The object ID of the image to replace
     * @param uri - Public URL of the replacement image
     * @returns The batch update result with document ID and reply count
     */
    replaceImage: (documentId: string, imageObjectId: string, uri: string) =>
      replaceImage(docs, documentId, imageObjectId, uri),

    /**
     * Deletes a positioned (floating) object from the document.
     *
     * @param documentId - The ID of the target document
     * @param objectId - The object ID of the positioned object to delete
     * @returns The batch update result with document ID and reply count
     */
    deletePositionedObject: (documentId: string, objectId: string) =>
      deletePositionedObject(docs, documentId, objectId),

    /**
     * Creates a page header in the document.
     *
     * @param documentId - The ID of the target document
     * @param type - The header type (defaults to 'DEFAULT')
     * @param sectionBreakIndex - Optional index of the section break whose header to create
     * @returns The batch update result with document ID and reply count
     */
    createHeader: (documentId: string, type?: string, sectionBreakIndex?: number) =>
      createHeader(docs, documentId, type, sectionBreakIndex),

    /**
     * Creates a page footer in the document.
     *
     * @param documentId - The ID of the target document
     * @param type - The footer type (defaults to 'DEFAULT')
     * @param sectionBreakIndex - Optional index of the section break whose footer to create
     * @returns The batch update result with document ID and reply count
     */
    createFooter: (documentId: string, type?: string, sectionBreakIndex?: number) =>
      createFooter(docs, documentId, type, sectionBreakIndex),

    /**
     * Deletes a header from the document by its header ID.
     *
     * @param documentId - The ID of the target document
     * @param headerId - The ID of the header to delete
     * @returns The batch update result with document ID and reply count
     */
    deleteHeader: (documentId: string, headerId: string) =>
      deleteHeader(docs, documentId, headerId),

    /**
     * Deletes a footer from the document by its footer ID.
     *
     * @param documentId - The ID of the target document
     * @param footerId - The ID of the footer to delete
     * @returns The batch update result with document ID and reply count
     */
    deleteFooter: (documentId: string, footerId: string) =>
      deleteFooter(docs, documentId, footerId),

    /**
     * Inserts a footnote reference at the given index or at the end of the document.
     *
     * @param documentId - The ID of the target document
     * @param index - Optional insertion index (defaults to end of document)
     * @returns The batch update result with document ID and reply count
     */
    createFootnote: (documentId: string, index?: number) =>
      createFootnote(docs, documentId, index),
  };
}

async function insertInlineImage(
  docs: DocsAPI,
  documentId: string,
  uri: string,
  width?: number,
  height?: number,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  const request: docs_v1.Schema$InsertInlineImageRequest = {
    uri,
    location: { index: targetIndex },
  };

  if (width !== undefined || height !== undefined) {
    request.objectSize = {};
    if (width !== undefined) {
      request.objectSize.width = { magnitude: width, unit: 'PT' };
    }
    if (height !== undefined) {
      request.objectSize.height = { magnitude: height, unit: 'PT' };
    }
  }

  return batchUpdate(docs, documentId, [{ insertInlineImage: request }]);
}

async function replaceImage(
  docs: DocsAPI,
  documentId: string,
  imageObjectId: string,
  uri: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      replaceImage: {
        imageObjectId,
        uri,
      },
    },
  ]);
}

async function deletePositionedObject(
  docs: DocsAPI,
  documentId: string,
  objectId: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deletePositionedObject: {
        objectId,
      },
    },
  ]);
}

async function createHeader(
  docs: DocsAPI,
  documentId: string,
  type?: string,
  sectionBreakIndex?: number,
): Promise<BatchUpdateResult> {
  const request: docs_v1.Schema$CreateHeaderRequest = {
    type: type ?? 'DEFAULT',
  };

  if (sectionBreakIndex !== undefined) {
    request.sectionBreakLocation = { index: sectionBreakIndex };
  }

  return batchUpdate(docs, documentId, [{ createHeader: request }]);
}

async function createFooter(
  docs: DocsAPI,
  documentId: string,
  type?: string,
  sectionBreakIndex?: number,
): Promise<BatchUpdateResult> {
  const request: docs_v1.Schema$CreateFooterRequest = {
    type: type ?? 'DEFAULT',
  };

  if (sectionBreakIndex !== undefined) {
    request.sectionBreakLocation = { index: sectionBreakIndex };
  }

  return batchUpdate(docs, documentId, [{ createFooter: request }]);
}

async function deleteHeader(
  docs: DocsAPI,
  documentId: string,
  headerId: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteHeader: {
        headerId,
      },
    },
  ]);
}

async function deleteFooter(
  docs: DocsAPI,
  documentId: string,
  footerId: string,
): Promise<BatchUpdateResult> {
  return batchUpdate(docs, documentId, [
    {
      deleteFooter: {
        footerId,
      },
    },
  ]);
}

async function createFootnote(
  docs: DocsAPI,
  documentId: string,
  index?: number,
): Promise<BatchUpdateResult> {
  const targetIndex = index ?? (await getDocumentEndIndex(docs, documentId));

  return batchUpdate(docs, documentId, [
    {
      createFootnote: {
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
