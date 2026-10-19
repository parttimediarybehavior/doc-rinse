# Doc Rinse

Inspect Office files locally, choose supported cleanup actions, and inspect the result again. Doc Rinse is an experimental single-file HTML tool; it also has a Node command-line interface. It does not establish authorship, prove how a document was written, or guarantee removal of every private detail.

## Open it locally

Build from source with Node.js, then open the generated file in your browser:

```sh
node build.js
# Open dist/doc-rinse.html
```

Drop a `.docx`, `.xlsx` or `.pptx` file onto the page, or choose the embedded synthetic sample. No account, service or upload is needed. Keep the original, select cleanup options carefully, and review the downloaded copy in your Office application before sharing it. The older development artifact name was `palimpsest.html`; the release build uses `doc-rinse.html`.

## What it inspects and cleans

| Format | Inspection | Cleanup |
|---|---|---|
| DOCX | Stored author and organization fields, timestamps, revision identifiers, tracked deletions, comments, hidden-text candidates, image metadata and recoverable cropped image content | Selected metadata and content removals, followed by re-analysis. Some actions are selected by default; review them. White/tiny text removal is off by default. Image crop baking requires the browser. |
| XLSX | Hidden/very-hidden sheets, comments and author records, external relationships, defined names, formula caches, calculation chains and author metadata | Opt-in identity, comments and thumbnail cleanup. Structural candidates, formulas, caches, hidden sheets and external links are retained. |
| PPTX | Hidden slides, speaker notes, comments/authors, metadata, thumbnails and media with no incoming package relationship | Opt-in identity, comments, speaker-note parts and thumbnail cleanup. Hidden slides, external links and unreferenced media are retained. |

Spreadsheet/presentation cleanup defaults to no selected actions. Signed packages and packages with unresolved relationships are rejected by that cleanup path. Note masters, embedded files and other unlisted stores can still contain private data. Inspection does not execute formulas, macros or embedded programs.

Revision identifiers and language tags are application metadata, not a reliable account of human work. Stored editing time can be missing, stale or reset. A pattern may suggest a hypothesis; it does not identify an author, keyboard, location, paste event or AI use.

## What has been verified

On the frozen DOCX test set, LibreOffice 24.2.7.2 rendered **187 of 196 generated cleaned files (95.4%) with equal extracted text and page count**, using the preregistered comparison. Fifteen other inputs were rejected before cleanup, out of 211 total inputs. The baseline passed 177/196; ten failures were resolved with no baseline pass regressing. The revised run also preserves original basenames to avoid a filename-field test artifact.

The nine remaining non-passes are documented in [the renderer review](docs/roundtrip/RESULTS.md). Hidden-text removal can shift layout; identity cleanup can change text bound to company metadata; accepting tracked changes changes content and pagination. Three inputs could not be compared because their original files failed to load in LibreOffice. These remain non-passes.

This is one renderer/environment and a text/page-count comparison, not a pixel-perfect layout check or a Microsoft Word compatibility guarantee. Retaining white/tiny text by default contributes to the improvement. The built-in parsed-word check cannot verify image text, layout or content the parser missed. Re-analysis by the same engine is not independent proof of sanitization.

XLSX/PPTX have synthetic package tests and two small LibreOffice-generated renderer examples, not an equivalent corpus validation. Real-browser release smoke testing is a separate check. Do not treat Doc Rinse as guaranteed anonymization, legal sanitization, or a substitute for reviewing what you share.

## Verify the offline behavior yourself

The build embeds its JavaScript, styles and one synthetic sample. URL-based sample loading has been removed. The page includes a Content Security Policy before its executable content:

```text
default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'
```

This policy blocks network connections and external resource loading in browsers that enforce it, while allowing the bundled inline code/styles and local image previews. Inline script is allowed because the application ships as one file; CSP is not a guarantee against every possible bug. The tool contains no analytics. Browser extensions and the browser's own services are outside its control.

Open the file locally with developer tools' Network panel recording, load an embedded sample, inspect it, clean it and download the result. There should be no HTTP(S) requests from the tool. Repeat with the browser offline. You can also inspect the HTML source for the policy and run the static release checks below. Static checks do not replace a real-browser test.

## Command line

```sh
node cli.js file.docx --json
node cli.js file.docx --scrub copy.docx
node cli.js deck.pptx --scrub copy.pptx --only identity,comments,notes,thumbnail --json
```

`--all` explicitly selects every supported action, including content removal and DOCX white/tiny-text removal. For XLSX/PPTX, `--scrub` without `--only` or `--all` removes nothing. The CLI does not bake image crops; it reports that limitation.

## Development and checks

No third-party runtime packages are required. Python 3 runs the CLI test driver; LibreOffice, `pdfinfo` and `pdftotext` are needed for the separate renderer harness.

```sh
node test.js
node docx-safety-test.js
node office-test.js
python3 tests/cli-test.py
node build.js
node tests/office-ui-test.js
node tests/release-test.js
```

The UI test uses a minimal DOM harness, not an actual browser or layout engine. Renderer protocols, input hashes and review records are under `docs/roundtrip/`. Corpus files, third-party fixtures and local validation outputs are not release assets. Corpus files and older internal reports are not included.

## Sample provenance

The release build embeds only `samples/00-synthetic-everything.docx`, generated entirely by this project's `test.js`. Reproduce it with `node test.js --emit samples/00-synthetic-everything.docx`. It is original synthetic test data under the project MIT license, including its generated image bytes. See [samples/README.md](samples/README.md). Other development sample documents must not be included in release archives.

## License

MIT, copyright 2026 mothwake. See [LICENSE](LICENSE). The license covers this project's code and original synthetic samples, not third-party documents used during local validation.
