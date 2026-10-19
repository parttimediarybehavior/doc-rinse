# Renderer validation results (public summary)

Method: docs/roundtrip/PROTOCOL.md (frozen before the run). Each input was cleaned with default settings; original and cleaned copies were converted to PDF with LibreOffice 24.2.7.2 and compared by extracted text (pdftotext -layout) and page count. Inputs: 81 public documents from US government/university websites and 130 Apache POI test documents (hashes in inputs.json; the documents themselves are not redistributed).

| Outcome | Count |
|---|---|
| Same extracted text and page count (PASS) | 187 |
| Different (REVIEW_REQUIRED) | 6 |
| LibreOffice could not convert | 3 |
| Rejected before cleaning (unparseable / encrypted / fuzz inputs) | 15 |

**Strict rate: 187/196 cleaned outputs = 95.4%.** Not a pixel comparison, not Microsoft Word, one LibreOffice version.

A first run of an earlier version scored 177/196 (90.3%); it exposed two bugs (images referenced only from headers/footers/embedded-object previews were deleted; text in shapes was treated as invisible). Both were fixed; 10 files moved to PASS and none regressed.

## Every non-pass

| File | Outcome | Cause |
|---|---|---|
| `9f24b5cb_SVS-Validations-Single-and-Multi-Project-April-2015-tracked-PreRelease.docx` | REVIEW_REQUIRED | Accepting tracked changes changes text and pagination (isolated: tracked-only run reproduces it). |
| `58067.docx` | REVIEW_REQUIRED | Tracked-change removal (deleted revision text no longer rendered). |
| `Bug54771a.docx` | REVIEW_REQUIRED | Removing company metadata changes a content control bound to it, which displays that value (isolated: identity-only run reproduces it). |
| `ExternalEntityInText.docx` | CONVERSION_FAILED | LibreOffice failed to convert both original and cleaned file. |
| `Tika-792.docx` | REVIEW_REQUIRED | Tracked-change removal. |
| `bug59058.docx` | REVIEW_REQUIRED | Removing explicitly hidden text (incl. a hidden page-break marker) shifts a heading on page 26 by 18.75pt (isolated: hidden-text-only run reproduces it; 6 other actions pass). |
| `clusterfuzz-testcase-minimized-POIXWPFFuzzer-6733884933668864.docx` | CONVERSION_FAILED | LibreOffice failed on the ORIGINAL; the cleaned file converts. |
| `delins.docx` | REVIEW_REQUIRED | Tracked deletions accepted; deleted text no longer rendered. |
| `unicode-path.docx` | CONVERSION_FAILED | LibreOffice failed on the ORIGINAL; the cleaned file converts. |
