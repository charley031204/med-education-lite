# Initial migration

## Scope

- New independent plugin: `med-education-lite`.
- First skill: `pdf-to-rawmd`; existing local installation and original repository remain unchanged.
- Remove permanent source maps; preserve visible output, source fidelity, actual figures, fresh conversion contexts and independent full review.
- Use deterministic drafts, compact batched input, selective enlargement and targeted edits.
- Publish a private GitHub repository and commit/push the validated implementation. No installation or release tag was requested.

## Ownership

- Main: packaging, integration, validation report, Git publication.
- `lite_runtime`: scripts, runtime tests, dependencies.
- `lite_docs`: skill instructions and conditional references.
- Independent review/forward test follows implementation.

## Checks

- Runtime: 6 tests and 5 subtests passed in the Main validation run.
- Portable schema, Claude package/marketplace and skill validation passed. Version omission produces an intentional Claude strict-mode warning.
- Independent code review: found issues corrected and rechecked; no remaining blocking findings.
- PDF 3 pages and PPTX 2 slides: independently checked all content, 5 tables/105 original cells and 6 figures/PNGs; no fidelity defects found.
- Local installed skill: all 11 source files unchanged by SHA-256 comparison.
- No actual source documents, rendered pages, local dependencies or temporary extraction files enter Git.
- Detailed scope and limits: [validation/pdf-to-rawmd.md](validation/pdf-to-rawmd.md).

## Study pipeline migration

- Six skills now cover RAW, question-only Word, transcript polishing, PDF/Word study notes, Obsidian explanations and single-question answers.
- User requirements and command names are reflected in [pipeline.md](pipeline.md); detailed verification is in [validation/pipeline.md](validation/pipeline.md).
- Runtime structure uses local question/transcript indexes, bounded evidence packets, deterministic assembly and a shared DOCX/PDF source.
- Final integrated run: 29 tests and 10 subtests passed with Word conversion enabled. Independent integration review found no remaining blocking issues.
- Original PDF immutability is an explicit invariant with byte-preservation and overwrite-rejection regressions.
- Existing installed source files: 57 checked, zero differences. Personal source materials and actual outputs remain outside Git.
