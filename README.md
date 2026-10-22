# @kazibee/google-docs

Google Docs tool for kazibee. Create documents, read content, and edit text with batch updates.

## Install

In the Kazibee desktop app, open **Settings → Tools → Add Extension → Install** and enter
`github:kazibee/google-docs`. The app clones the latest commit and records its SHA; use
**Update** on the extension to move to a newer commit.

## Login

Open **Settings → Tools → google-docs → Commands** and run `login`. It needs a Google OAuth
app's `CLIENT_ID` and `CLIENT_SECRET`: set them in the extension's **Environment** tab, or pass
them as the command's arguments (`CLIENT_ID CLIENT_SECRET`).

## API

- `createDocument(title)`
- `getDocument(documentId)`
- `getDocumentText(documentId)`
- `getDocumentStructure(documentId)`
- `appendText(documentId, text)`
- `insertText(documentId, text, index?)`
- `insertHeading(documentId, text, level, index?)`
- `insertBulletList(documentId, items, index?)`
- `insertTable(documentId, rows, columns, index?)`
- `replaceAllText(documentId, findText, replaceText)`
- `replaceAllTextBatch(documentId, replacements)`
- `findTextRanges(documentId, query, matchCase?)`
- `applyTextStyle(documentId, range, style)`
- `applyParagraphStyle(documentId, range, style)`
- `createNamedRange(documentId, name, range)`
- `replaceNamedRange(documentId, name, text)`
- `insertPageBreak(documentId, index?)`
- `batchUpdate(documentId, requests)`
- `safeBatchUpdate(documentId, requests)`

## Usage

```javascript
const doc = await tools["google-docs"].createDocument("Weekly Notes");
await tools["google-docs"].insertHeading(doc.documentId, "Weekly Notes", 1);
await tools["google-docs"].insertBulletList(doc.documentId, [
  "Ship filter API",
  "Review release notes"
]);
await tools["google-docs"].replaceAllTextBatch(doc.documentId, [
  { findText: "filter", replaceText: "filtering" }
]);
const text = await tools["google-docs"].getDocumentText(doc.documentId);
```
