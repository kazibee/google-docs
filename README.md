# @kazibee/google-docs

Google Docs tool for kazibee. Create documents, read content, and edit text with batch updates.

## Install

```bash
kazibee install google-docs github:kazibee/google-docs
```

Install globally with `-g`:

```bash
kazibee install -g google-docs github:kazibee/google-docs
```

## Login

```bash
kazibee google-docs login
```

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
