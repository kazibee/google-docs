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

export default function main(env: Env) {
  const auth = createAuthClient(env);
  return createDocsClient(auth);
}
