import type { docs_v1 } from '@googleapis/docs';
import type { PendingChange, HeadingLevel } from './shadow-document';

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Compiles an array of pending shadow-document changes into Google Docs API
 * batch-update requests. Index-based operations are sorted from highest index
 * to lowest so earlier mutations never shift indices that later operations
 * depend on. `replaceAll` operations are index-independent and are emitted
 * first (order among them does not matter).
 */
export function compileRequests(changes: PendingChange[]): docs_v1.Schema$Request[] {
  const replaceAlls: docs_v1.Schema$Request[] = [];
  const indexedEntries: { index: number; requests: docs_v1.Schema$Request[] }[] = [];

  for (const change of changes) {
    switch (change.type) {
      case 'replaceAll':
        replaceAlls.push(compileReplaceAll(change.find, change.replace));
        break;
      case 'insert':
        indexedEntries.push({
          index: change.index,
          requests: compileInsert(change.index, change.text),
        });
        break;
      case 'delete':
        indexedEntries.push({
          index: change.startIndex,
          requests: compileDelete(change.startIndex, change.endIndex),
        });
        break;
      case 'insertHeading':
        indexedEntries.push({
          index: change.index,
          requests: compileInsertHeading(change.index, change.text, change.level),
        });
        break;
      case 'insertBullet':
        indexedEntries.push({
          index: change.index,
          requests: compileInsertBullet(change.index, change.items),
        });
        break;
    }
  }

  // Sort indexed operations by descending index to prevent shifting
  indexedEntries.sort((a, b) => b.index - a.index);

  return [
    ...replaceAlls,
    ...indexedEntries.flatMap((entry) => entry.requests),
  ];
}

// ---------------------------------------------------------------------------
// Per-change compilers
// ---------------------------------------------------------------------------

function compileInsert(index: number, text: string): docs_v1.Schema$Request[] {
  return [
    { insertText: { location: { index }, text } },
  ];
}

function compileDelete(startIndex: number, endIndex: number): docs_v1.Schema$Request[] {
  return [
    { deleteContentRange: { range: { startIndex, endIndex } } },
  ];
}

function compileReplaceAll(find: string, replace: string): docs_v1.Schema$Request {
  return {
    replaceAllText: {
      containsText: { text: find, matchCase: true },
      replaceText: replace,
    },
  };
}

function compileInsertHeading(
  index: number,
  text: string,
  level: HeadingLevel,
): docs_v1.Schema$Request[] {
  const endIndex = index + text.length;

  return [
    { insertText: { location: { index }, text } },
    {
      updateParagraphStyle: {
        range: { startIndex: index, endIndex },
        paragraphStyle: { namedStyleType: `HEADING_${level}` },
        fields: 'namedStyleType',
      },
    },
  ];
}

function compileInsertBullet(index: number, items: string[]): docs_v1.Schema$Request[] {
  const listText = `${items.join('\n')}\n`;
  const endIndex = index + listText.length;

  return [
    { insertText: { location: { index }, text: listText } },
    {
      createParagraphBullets: {
        range: { startIndex: index, endIndex },
        bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
      },
    },
  ];
}
