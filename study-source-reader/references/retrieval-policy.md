# Retrieval policy

## Data path

`catalog -> current validated revision -> selected retrieval profile/chunks.json -> relevant source.md region -> assets -> retained normalized PDF for general visual precision -> current identical original PPTX for PPT-specific facts`

The helper uses deterministic substring, heading, English-token, and Chinese n-gram scoring. It has no embeddings, vector database, network service, parser, or MinerU subprocess. A revision-level `chunks_path` selects v2 profiled chunks; the root `chunks.json` remains a read-only v1 fallback. Returned chunk location fields are copied only when those fields genuinely exist. Catalog/cache lookup remains authoritative across legacy `<Course>/Sources` or `课程资料库/<Course>/<Document>` publications and the current `课程资料/<Course>/<Document>` layout; `笔记` is not used as an ingested-source store.

## Document resolution

Explicit course/title filters take priority. With several remaining candidates, a unique strong title/course/filename match may resolve the source; otherwise return `AMBIGUOUS` with candidates. Never choose an obviously ambiguous document silently.

## Precision boundary

Concepts such as “what is convolution?” normally use cached chunks and Markdown. Requests for exact equations, variable identity, Greek characters, signs/operators, indices, table cells, circuit/waveform/figure annotations, or layout require visual evidence because scanned STEM notation can be misrecognized.

Use an extracted asset when it contains the needed evidence. For normalized PPTX revisions, a hash-validated cache PDF can support ordinary formula, diagram, waveform, and layout verification. It cannot establish animation/build order, hidden slides, transitions, speaker notes, editable object structure, or information lost during PDF rendering; those questions require the original PPTX. For originals, prefer the catalog's current source binding, then deterministically try same-hash aliases, retained source, and immutable extraction provenance. A candidate must still exist and match the revision SHA-256. If no suitable evidence qualifies, report that boundary instead of fabricating precision.

For ChatGPT handoff, upload `source.md` first and `chunks.json` for large repeated retrieval. Include the original document when formulas or visuals dominate. Do not create another AI package.
