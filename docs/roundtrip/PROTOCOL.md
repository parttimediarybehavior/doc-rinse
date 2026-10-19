# Renderer comparison protocol — frozen before the corpus run

2026-10-18. Frozen before the corpus run. (Public copy: one internal path reference removed from this line; rules unchanged.)

Inventory: all 81 corpus and 130 fixture DOCX paths and SHA-256 hashes in inputs.json, including malformed inputs. Expected historical analyzable subset is 196; reconcile actual preparation counts rather than assume that denominator. No input is silently skipped. Original files remain unchanged and untracked.

Run the default CLI scrub, without `--all` or selected-fix overrides, for every input. Record errors and output hashes. For each successfully generated cleaned file, convert BOTH original and cleaned to PDF with the same LibreOffice version, fonts and isolated profile. Extract with `pdftotext -layout -enc UTF-8`; obtain page count with `pdfinfo`. Per-conversion timeout: 90 seconds. Keep conversion logs, PDFs and extracted text locally, outside git.

Automatic PASS requires both conversions, positive equal page counts, and equal extracted text after CRLF normalization and removal of trailing spaces/tabs on each line. Preserve line order, page separators, case and internal whitespace. No word sorting, fuzzy score or dropping empty pages. This is a text/page-count check, not proof of pixel-perfect layout or Microsoft Word behavior.

Any difference is REVIEW_REQUIRED, never silently accepted. An intended-removal exception can be accepted only with a per-file explanation connecting the exact diff to a selected cleanup action and independently checked original XML/rendered content. Page changes require explicit explanation too. No automatic subtraction of the engine's own claimed removed text. Review unresolved cases as failures for the ≥95% gate. Conversion errors, timeouts, missing outputs and unparsable page counts are NOT passes. Original parse failures are reported separately from generated-output failures; publish both denominators.

Report exact file names and categories for all non-passes. Strict rate = automatic passes / generated cleaned outputs; reviewed rate includes individually justified exceptions only. If none convert because the environment cannot run LibreOffice, report ENVIRONMENT_BLOCKED, not a product accuracy failure or 0% verified correctness. Preserve the original failure and run in Claude's normal supported environment, without permission bypasses.

The initial single-file environment smoke test failed twice (exit 1, no PDF; javaldx warning). The second used writable runtime/profile locations and disabled dconf. No corpus-render outcome was inspected before freezing this protocol.

## V2 run amendment — frozen before the revised run

Use `--out .validation/roundtrip-v2` for both prepare and render; never overwrite the original run. Same211 input hashes and same strict comparator/denominator rules. Cleaned files retain their original basename in a separate input-clean directory, eliminating the known FILENAME-field naming artifact. This is a disclosed harness correction, not a relaxed comparison. Preparation records product commit and source/harness hashes. Default cleanup now preserves white/tiny text candidates and images referenced by any package relationship. No new renderer result inspected when adding this amendment.
