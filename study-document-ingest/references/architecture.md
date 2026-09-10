# Architecture and invariants

## Ingestion flow

`source -> SHA-256/extraction profile -> extraction cache lookup -> optional quality-PPTX local PDF normalization -> isolated MinerU staging -> source.md/assets validation -> extraction cache commit`

Then, independently:

`cached source.md -> retrieval profile -> MinerU chunking module -> retrieval/<fingerprint>/chunks.json -> optional human-only publication swap -> catalog commit`

The source file is opened read-only. Changing filename, mtime, or directory cannot masquerade as changed content because source bytes are hashed. A profile fingerprint prevents incompatible extraction settings from sharing a cache entry.

## Identity

- `source_sha256`: SHA-256 of source bytes.
- `extraction_profile_fingerprint`: SHA-256 of parsing settings only: engine, API, model, OCR policy, language policy, split/parser policy, and—only for normalized PPTX—input modality plus normalizer backend/version/schema.
- `retrieval_profile_fingerprint`: SHA-256 of retrieval-only settings: chunk enablement, chunk size, chunker identity/hash, and future lexical options.
- `cache_key`: SHA-256 of deterministic JSON containing `source_sha256` and `extraction_profile_fingerprint`; filenames and retrieval settings are excluded.
- `document_id`: stable logical catalog identity. An existing source-path alias wins; otherwise a unique identical-content document is rebound even when its inferred filename/title changed. Ambiguous identical-content matches still require context or create a separate identity.
- `revision_id`: derived from the cache key. Modified bytes or a changed compatible-profile field creates a new revision while old cache directories remain.

MinerU version and a fingerprint of its relevant files are provenance only. They do not invalidate historical validated caches. Bump the extraction-profile schema only for parsing incompatibility, the retrieval-profile schema only for derivation incompatibility, and the manifest schema only for layout changes.

Changing only a retrieval setting never changes `cache_key` or invokes MinerU extraction. It selects or regenerates a retrieval artifact from validated `source.md`. Valid v1 caches are projected onto the extraction-only profile and migrated into the v2 layout without a MinerU call.

Quality-mode PPTX never uses the direct PPTX cloud path. It selects PowerPoint native export first and LibreOffice headless second, writes only inside staging, validates the PDF structurally and against a reliable slide count when available, then sends that PDF through the ordinary quality extraction path. The original PPTX hash remains source identity. A successful derivative is retained at `normalized/source.pdf`; it is not copied into the human-facing Obsidian course library. PDF inputs and non-quality modes retain their existing direct routing.

## Obsidian publication boundary

The required publication root is `课程资料`. A publication contains only `source.md` and, when present, `assets/`:

```text
<vault>/
  笔记/                             # reserved for user-authored or assisted notes
    <Course>/
  课程资料/
    <Course>/
      <Document>/
        source.md
        assets/
```

The ingest pipeline never creates or writes `笔记`. No other visible top-level study hierarchy is used. Catalog data, manifests, provenance, normalized derivatives, retrieval chunks, and runtime state remain under the external state directory. The reader resolves through that state, not through the publication copy.

A stored publication-tree fingerprint permits atomic replacement only when the existing `source.md`/`assets` tree still matches the last pipeline commit and contains no unexpected entries. An untracked collision or a user-modified publication directory is never overwritten; a new document receives a deterministic suffixed directory, while a modified known publication causes a safe refusal.

Legacy migration planning is read-only. A `<Course>/Sources/<Document>` folder is eligible only when one catalog document, its current published revision, the validated extraction cache, a matching legacy vault manifest, exact source/asset hashes, an absent destination, and a clean visible tree all agree. Any missing or ambiguous evidence is refused. Moving or deleting legacy material requires a later explicit approval.

## Provenance and current source bindings

The extraction manifest is immutable provenance for the path that first materialized that cache. Current paths live in the catalog. Every successful ingest refreshes `current_source_binding` and keeps content-hash-qualified aliases. The reader checks the current binding first, then aliases by newest `last_seen_at` and normalized path, and verifies source bytes before precision-sensitive use. Missing or changed paths are skipped; if no identical source exists, exact visual verification is reported unavailable.

## MinerU policy

- Quality: `--engine cloud --api auto --model vlm`, default retrieval chunk size 2000.
- Fast: `--engine auto --api auto --model pipeline`; it may use cloud when the optional local engine is absent.
- Private: PDF only, requires `pymupdf4llm`, invokes `--engine local`, and rejects unsupported inputs before subprocess construction. It cannot silently fall back to cloud.

MinerU's installed `scripts/chunking.py` remains the only chunker. The wrapper calls its `chunk_markdown` function locally against cached Markdown; it does not copy or reimplement that algorithm.

PDF page-count preflight reads MinerU's installed backend limits and only adds MinerU `--split` when necessary. A verified page-limit failure can receive one native-split retry. MinerU retains responsibility for splitting and network retry/backoff. A very sparse quality result can receive at most one OCR attempt.

`--resume` is not the persistent cache authority. Content/profile identity is checked first, and the wrapper intentionally uses isolated staging rather than trusting a filename-based resume marker.

## Commit and rollback

New extraction never writes directly into a valid cache or publication. Cache replacement uses an incoming sibling plus backup/restore. Publication prepares and validates a new human-facing source directory, atomically swaps it, and restores the old directory if the swap fails. No hidden metadata directory is created in the vault. Catalog JSON is updated last through temp-file, flush, and replace.

For an explicitly requested publication, catalog current revision advances only after successful publication. Without a configured vault, ingestion remains successful as cache/catalog-only and records a warning.

## Windows encoding isolation

MinerU is started with a subprocess argument list and `shell=False`. stdout/stderr are pipes, not the terminal. The child receives `PYTHONUTF8=1` and `PYTHONIOENCODING=utf-8`; the wrapper decodes bytes and parses JSON internally. Large MinerU JSON is neither echoed nor logged. If a final `UnicodeEncodeError`/GBK status-render failure occurs after Markdown, chunks, and assets independently validate, the extraction is accepted with a provenance warning.

## Batch and directories

Directory traversal uses only file types advertised by MinerU 3.3.1 and skips caches, staging, `.study-docs`, hidden/generated output directories, and temporary files. Documents are committed independently and processed serially in v1 to bound API concurrency and isolate failures.
