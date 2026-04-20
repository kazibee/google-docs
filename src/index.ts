import { createAuthClient, type Env } from './auth';
import { createDocsClient } from './docs-client';

export type { Env } from './auth';
export type {
  DocumentInfo,
  BatchUpdateResult,
  IndexRange,
  TextMatchRange,
  DocumentSection,
  DocumentStructure,
  TextStyleInput,
  ParagraphStyleInput,
} from './docs-client';
export type { TableCellLocation, TableRange, TableCellStyleInput } from './table-ops';
export type { ImageSize } from './media-ops';
export type { DocumentStyleInput, SectionStyleInput } from './style-content-ops';
export { ShadowDocument } from './shadow-document';
export type {
  ShadowNode,
  ParagraphNode,
  TableNode,
  SectionBreakNode,
  PendingChange,
} from './shadow-document';
export { compileRequests } from './request-compiler';

export default function main(env: Env) {
  const auth = createAuthClient(env);
  return createDocsClient(auth);
}
