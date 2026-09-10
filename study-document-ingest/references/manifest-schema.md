# Manifest schema

Each cache entry contains:

```text
cache/<cache-key>/
  source.md
  assets/
  normalized/                 # quality-mode PPTX only
    source.pdf
  extraction.json
  manifest.json
  retrieval/
    <retrieval-profile-fingerprint>/
      chunks.json
      retrieval.json
```

`source.md` is normalized only so local MinerU image links point under `assets/`. `extraction.json` stores a compact, token-redacted process record without Markdown or embedded chunks. Each `chunks.json` is deterministically regenerated from cached Markdown by MinerU's installed chunking module; `retrieval.json` binds its hash and chunk count to the extraction cache key and retrieval profile.

The v2 extraction manifest records schema/document/revision/cache identifiers; immutable first-materialization path, filename, size, mtime, and SHA-256; course/title context; extraction profile and fingerprint; MinerU version/fingerprint; requested and observed backend data; OCR/split settings; UTC timestamps; artifact paths/status/warnings; optional retained-source path; validation counts; and SHA-256/size for `source.md`, assets, and any retained normalized derivative. Retrieval settings and chunks are intentionally excluded from extraction identity and artifact hashes.

For quality-mode PPTX, `original_source` identifies the canonical PPTX and its immutable extraction-time binding, while `normalized_source` identifies the cache-only PDF derivative, its SHA-256, backend/version, slide/page counts, and conversion duration. `normalized_pdf_sha256` is also recorded at top level. The catalog—not the immutable cache manifest—remains authoritative for the current source binding after a move or rename.

Fields not exposed by MinerU, such as the actual model in some result paths, remain `null`. They are never inferred as facts.

The catalog is the authority for current logical-document aliases, `current_source_binding`, revisions, each revision's selected retrieval profile/chunks path, and the current publication path/tree fingerprint. A cache manifest records the context that first materialized that shared extraction. Publication does not emit a second manifest into the vault; the visible copy contains only `source.md` and `assets/`.

Fatal extraction validation includes missing/empty/invalid-UTF-8 Markdown, path escape, inconsistent hashes, and catastrophic image-reference loss. Retrieval validation independently rejects missing/invalid/empty chunks or mismatched profile/source/chunks hashes. A small number of unresolved local images is a warning recorded in the extraction manifest.

Validated schema-v1 entries remain readable. On first reuse, the wrapper verifies the v1 artifact hashes, projects its combined profile onto extraction-only fields, creates a v2 extraction entry, and regenerates the requested chunks locally without MinerU extraction.
