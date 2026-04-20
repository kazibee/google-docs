import type { DocumentStructure, DocumentSection } from './docs-client';

// ---------------------------------------------------------------------------
// Node types - local representation of document structural elements
// ---------------------------------------------------------------------------

export interface ParagraphNode {
  type: 'paragraph';
  startIndex: number;
  endIndex: number;
  text: string;
  heading?: string;
  isBullet?: boolean;
}

export interface TableNode {
  type: 'table';
  startIndex: number;
  endIndex: number;
  rowCount: number;
  columnCount: number;
}

export interface SectionBreakNode {
  type: 'sectionBreak';
  startIndex: number;
  endIndex: number;
}

export type ShadowNode = ParagraphNode | TableNode | SectionBreakNode;

// ---------------------------------------------------------------------------
// Heading level type
// ---------------------------------------------------------------------------

export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

// ---------------------------------------------------------------------------
// PendingChange - consumed by the RequestCompiler (Task #2)
// ---------------------------------------------------------------------------

export type PendingChange =
  | { type: 'insert'; index: number; text: string }
  | { type: 'delete'; startIndex: number; endIndex: number }
  | { type: 'replaceAll'; find: string; replace: string }
  | { type: 'insertHeading'; index: number; text: string; level: HeadingLevel }
  | { type: 'insertBullet'; index: number; items: string[] };

// ---------------------------------------------------------------------------
// ShadowDocument
// ---------------------------------------------------------------------------

export class ShadowDocument {
  readonly documentId: string;
  readonly title: string;
  private nodes: ShadowNode[];
  private originalNodes: ShadowNode[];
  private pendingChanges: PendingChange[] = [];

  private constructor(documentId: string, title: string, nodes: ShadowNode[]) {
    this.documentId = documentId;
    this.title = title;
    this.nodes = nodes;
    this.originalNodes = deepCloneNodes(nodes);
  }

  /** Build a ShadowDocument from a DocumentStructure returned by getDocumentStructure(). */
  static fromStructure(structure: DocumentStructure): ShadowDocument {
    const nodes: ShadowNode[] = structure.sections.map(sectionToNode);
    return new ShadowDocument(structure.documentId, structure.title, nodes);
  }

  // -----------------------------------------------------------------------
  // Accessors
  // -----------------------------------------------------------------------

  /** Returns a shallow copy of the current node list. */
  getNodes(): readonly ShadowNode[] {
    return this.nodes;
  }

  /** Returns accumulated pending changes in order. */
  getPendingChanges(): readonly PendingChange[] {
    return this.pendingChanges;
  }

  /** Clears the pending changes list (call after a successful commit). */
  clearPendingChanges(): void {
    this.pendingChanges = [];
    this.originalNodes = deepCloneNodes(this.nodes);
  }

  // -----------------------------------------------------------------------
  // Mutation operations
  // -----------------------------------------------------------------------

  /** Insert text at the given document index. Shifts all downstream indices. */
  insertText(index: number, text: string): void {
    this.pendingChanges.push({ type: 'insert', index, text });
    this.applyInsert(index, text);
  }

  /** Delete a content range. Shifts all downstream indices backward. */
  deleteRange(startIndex: number, endIndex: number): void {
    if (endIndex <= startIndex) return;
    this.pendingChanges.push({ type: 'delete', startIndex, endIndex });
    this.applyDelete(startIndex, endIndex);
  }

  /** Find and replace all occurrences of `find` with `replace` across paragraph text. */
  replaceText(find: string, replace: string): void {
    if (!find) return;
    this.pendingChanges.push({ type: 'replaceAll', find, replace });
    this.applyReplaceAll(find, replace);
  }

  /** Insert a heading at the given index. */
  insertHeading(index: number, text: string, level: HeadingLevel): void {
    const headingText = text.endsWith('\n') ? text : `${text}\n`;
    this.pendingChanges.push({ type: 'insertHeading', index, text: headingText, level });
    this.applyInsertParagraph(index, headingText, `HEADING_${level}`, false);
  }

  /** Insert a bullet list at the given index. */
  insertBulletList(index: number, items: string[]): void {
    this.pendingChanges.push({ type: 'insertBullet', index, items });
    this.applyInsertBulletList(index, items);
  }

  // -----------------------------------------------------------------------
  // Render & Diff
  // -----------------------------------------------------------------------

  /** Returns a plain-text representation of the current document state. */
  render(): string {
    const parts: string[] = [];
    for (const node of this.nodes) {
      if (node.type === 'paragraph') {
        if (node.heading && node.heading !== 'NORMAL_TEXT') {
          const marker = headingMarker(node.heading);
          parts.push(`${marker} ${node.text}`);
        } else if (node.isBullet) {
          parts.push(`- ${node.text}`);
        } else {
          parts.push(node.text);
        }
      } else if (node.type === 'table') {
        parts.push(`[Table ${node.rowCount}x${node.columnCount}]\n`);
      } else {
        parts.push('---\n');
      }
    }
    return parts.join('');
  }

  /** Returns a simple before/after diff for paragraphs that changed. */
  diff(): string {
    const originalTextMap = new Map<number, string>();
    for (const node of this.originalNodes) {
      if (node.type === 'paragraph') {
        originalTextMap.set(node.startIndex, node.text);
      }
    }

    const lines: string[] = [];
    for (const node of this.nodes) {
      if (node.type !== 'paragraph') continue;
      const originalText = originalTextMap.get(node.startIndex);

      if (originalText === undefined) {
        // New paragraph
        lines.push(`+ [${node.startIndex}] ${node.text.trimEnd()}`);
      } else if (originalText !== node.text) {
        // Changed paragraph
        lines.push(`- [${node.startIndex}] ${originalText.trimEnd()}`);
        lines.push(`+ [${node.startIndex}] ${node.text.trimEnd()}`);
        originalTextMap.delete(node.startIndex);
      } else {
        originalTextMap.delete(node.startIndex);
      }
    }

    // Paragraphs that were in original but no longer exist
    for (const [idx, text] of originalTextMap) {
      lines.push(`- [${idx}] ${text.trimEnd()}`);
    }

    return lines.length > 0 ? lines.join('\n') : '(no changes)';
  }

  // -----------------------------------------------------------------------
  // Internal mutation helpers
  // -----------------------------------------------------------------------

  private applyInsert(index: number, text: string): void {
    const length = text.length;
    const targetNodeIdx = this.findNodeContaining(index);

    if (targetNodeIdx !== -1 && this.nodes[targetNodeIdx].type === 'paragraph') {
      const node = this.nodes[targetNodeIdx] as ParagraphNode;
      const offset = index - node.startIndex;
      node.text = node.text.slice(0, offset) + text + node.text.slice(offset);
      node.endIndex += length;
      this.shiftNodesAfter(targetNodeIdx, length);
    } else {
      // Insert as a new paragraph node at the target position
      const newNode: ParagraphNode = {
        type: 'paragraph',
        startIndex: index,
        endIndex: index + length,
        text,
      };
      const insertPos = this.findInsertPosition(index);
      this.nodes.splice(insertPos, 0, newNode);
      this.shiftNodesAfter(insertPos, length);
    }
  }

  private applyDelete(startIndex: number, endIndex: number): void {
    const length = endIndex - startIndex;
    const nodesToRemove: number[] = [];

    for (let i = 0; i < this.nodes.length; i++) {
      const node = this.nodes[i];

      // Node fully within delete range: remove it
      if (node.startIndex >= startIndex && node.endIndex <= endIndex) {
        nodesToRemove.push(i);
        continue;
      }

      // Node partially overlaps with delete range
      if (node.startIndex < endIndex && node.endIndex > startIndex) {
        if (node.type === 'paragraph') {
          const pNode = node as ParagraphNode;
          const clipStart = Math.max(startIndex, pNode.startIndex) - pNode.startIndex;
          const clipEnd = Math.min(endIndex, pNode.endIndex) - pNode.startIndex;
          pNode.text = pNode.text.slice(0, clipStart) + pNode.text.slice(clipEnd);
          const removed = clipEnd - clipStart;
          pNode.endIndex -= removed;
        }
        continue;
      }

      // Node is entirely after delete range: shift backward
      if (node.startIndex >= endIndex) {
        node.startIndex -= length;
        node.endIndex -= length;
      }
    }

    // Remove fully deleted nodes in reverse order to preserve indices
    for (let i = nodesToRemove.length - 1; i >= 0; i--) {
      this.nodes.splice(nodesToRemove[i], 1);
    }
  }

  private applyReplaceAll(find: string, replace: string): void {
    const diff = replace.length - find.length;
    let cumulativeShift = 0;

    for (const node of this.nodes) {
      node.startIndex += cumulativeShift;
      node.endIndex += cumulativeShift;

      if (node.type === 'paragraph') {
        const pNode = node as ParagraphNode;
        let count = 0;
        let idx = pNode.text.indexOf(find);
        while (idx !== -1) {
          count++;
          idx = pNode.text.indexOf(find, idx + find.length);
        }

        if (count > 0) {
          pNode.text = pNode.text.replaceAll(find, replace);
          const totalShift = count * diff;
          pNode.endIndex += totalShift;
          cumulativeShift += totalShift;
        }
      }
    }
  }

  private applyInsertParagraph(
    index: number,
    text: string,
    heading: string | undefined,
    isBullet: boolean,
  ): void {
    const length = text.length;
    const newNode: ParagraphNode = {
      type: 'paragraph',
      startIndex: index,
      endIndex: index + length,
      text,
      heading,
      isBullet: isBullet || undefined,
    };

    const insertPos = this.findInsertPosition(index);
    // If inserting into the middle of an existing paragraph, split it
    if (insertPos > 0) {
      const prev = this.nodes[insertPos - 1];
      if (prev.type === 'paragraph' && prev.startIndex < index && prev.endIndex > index) {
        const pNode = prev as ParagraphNode;
        const offset = index - pNode.startIndex;
        const tailText = pNode.text.slice(offset);
        pNode.text = pNode.text.slice(0, offset);
        pNode.endIndex = index;

        const tailNode: ParagraphNode = {
          type: 'paragraph',
          startIndex: index + length,
          endIndex: index + length + tailText.length,
          text: tailText,
          heading: pNode.heading,
          isBullet: pNode.isBullet,
        };

        this.nodes.splice(insertPos, 0, newNode, tailNode);
        this.shiftNodesAfter(insertPos + 1, length);
        return;
      }
    }

    this.nodes.splice(insertPos, 0, newNode);
    this.shiftNodesAfter(insertPos, length);
  }

  private applyInsertBulletList(index: number, items: string[]): void {
    const bulletNodes: ParagraphNode[] = [];
    let currentIndex = index;

    for (const item of items) {
      const itemText = item.endsWith('\n') ? item : `${item}\n`;
      bulletNodes.push({
        type: 'paragraph',
        startIndex: currentIndex,
        endIndex: currentIndex + itemText.length,
        text: itemText,
        isBullet: true,
      });
      currentIndex += itemText.length;
    }

    const totalLength = currentIndex - index;
    const insertPos = this.findInsertPosition(index);

    // If inserting into the middle of an existing paragraph, split it
    if (insertPos > 0) {
      const prev = this.nodes[insertPos - 1];
      if (prev.type === 'paragraph' && prev.startIndex < index && prev.endIndex > index) {
        const pNode = prev as ParagraphNode;
        const offset = index - pNode.startIndex;
        const tailText = pNode.text.slice(offset);
        pNode.text = pNode.text.slice(0, offset);
        pNode.endIndex = index;

        const tailNode: ParagraphNode = {
          type: 'paragraph',
          startIndex: index + totalLength,
          endIndex: index + totalLength + tailText.length,
          text: tailText,
          heading: pNode.heading,
          isBullet: pNode.isBullet,
        };

        this.nodes.splice(insertPos, 0, ...bulletNodes, tailNode);
        this.shiftNodesAfter(insertPos + bulletNodes.length, totalLength);
        return;
      }
    }

    this.nodes.splice(insertPos, 0, ...bulletNodes);
    this.shiftNodesAfter(insertPos + bulletNodes.length - 1, totalLength);
  }

  /** Find the node index that contains the given document index. Returns -1 if none. */
  private findNodeContaining(index: number): number {
    for (let i = 0; i < this.nodes.length; i++) {
      const node = this.nodes[i];
      if (index >= node.startIndex && index < node.endIndex) {
        return i;
      }
    }
    return -1;
  }

  /** Find the position where a new node should be inserted for the given index. */
  private findInsertPosition(index: number): number {
    for (let i = 0; i < this.nodes.length; i++) {
      if (this.nodes[i].startIndex >= index) {
        return i;
      }
    }
    return this.nodes.length;
  }

  /** Shift all nodes after position `afterIdx` by `delta`. */
  private shiftNodesAfter(afterIdx: number, delta: number): void {
    for (let i = afterIdx + 1; i < this.nodes.length; i++) {
      this.nodes[i].startIndex += delta;
      this.nodes[i].endIndex += delta;
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sectionToNode(section: DocumentSection): ShadowNode {
  switch (section.type) {
    case 'table':
      return {
        type: 'table',
        startIndex: section.startIndex,
        endIndex: section.endIndex,
        rowCount: section.rowCount ?? 0,
        columnCount: section.columnCount ?? 0,
      };
    case 'sectionBreak':
      return {
        type: 'sectionBreak',
        startIndex: section.startIndex,
        endIndex: section.endIndex,
      };
    default:
      return {
        type: 'paragraph',
        startIndex: section.startIndex,
        endIndex: section.endIndex,
        text: section.text ?? '',
        heading: section.heading,
      };
  }
}

function deepCloneNodes(nodes: ShadowNode[]): ShadowNode[] {
  return nodes.map((node) => ({ ...node }));
}

function headingMarker(heading: string): string {
  const match = heading.match(/^HEADING_(\d)$/);
  if (!match) return '';
  const level = parseInt(match[1], 10);
  return '#'.repeat(level);
}
