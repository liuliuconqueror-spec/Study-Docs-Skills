---
name: study-document-ingest
description: Persistently import university course PDF, Office, and image sources into a validated content-addressed MinerU cache and catalog, with optional rollback-safe Obsidian publication. Use for requests to ingest, import, add, or prepare course material for repeated study. Do not use for one-off conversion or for answering from already-ingested sources.
---

# Study Document Ingest

Create reusable study sources without reparsing unchanged files. MinerU is the sole parser and chunker; this skill only orchestrates identity, validation, caching, cataloging, and optional publication.

## Run

Use Python 3.10+ and the deterministic script. For the default user-level Codex installation:

```powershell
python "$HOME\.agents\skills\study-document-ingest\scripts\ingest.py" `
  <INPUT> [<INPUT> ...]
```

If the skills were installed elsewhere, run the same script from that location. Set
`STUDY_DOCS_PYTHON`, `STUDY_DOCS_MINERU_SCRIPT`, or `STUDY_DOCS_MINERU_CHUNKER`
only when automatic discovery does not match the installation.

Quality mode is the default. PDFs go directly to MinerU cloud + API auto + VLM. On Windows, quality-mode PPTX is first normalized locally to PDF using an isolated Microsoft PowerPoint native export (preferred) or LibreOffice headless export; absence of both is a pre-MinerU failure. The validated derivative is retained only in the cache. MinerU's installed local chunking module then derives profiled chunks from cached Markdown. MinerU's native split path is used when the active page limit requires it.

## Cloud-consent boundary

Explicitly providing files and requesting ingestion constitutes consent to send uncached documents to the configured MinerU cloud parser. For normal `quality` ingestion, including an unspecified mode that resolves to the default `quality` profile, do not request a second confirmation solely because an explicitly uploaded, selected, or provided source is uncached—provided the user explicitly asked to ingest or import it into a course and the operation only parses, caches, and optionally publishes course source material. A compatible cache hit sends nothing to cloud and needs no confirmation.

This consent does not authorize deletion of user files, overwriting unrecognized or user-modified publication directories, resolving ambiguous ownership, moving or modifying personal notes, destructive migration, use of files the user did not explicitly supply, or other unexpected side effects. Private mode must never upload to cloud or fall back to cloud.

- Use `--course` and `--title` when supplied. Otherwise accept `_Inbox` when course inference is uncertain.
- Use `--no-publish` when only the durable cache/catalog is wanted.
- Use `--vault <existing-vault>` only for an explicitly selected vault. Never guess among multiple vaults.
- Use `--dry-run` before a large directory import when classification or destinations need review.
- Human-facing publication is limited to `<vault>\课程资料\<Course>\<Document>\source.md` plus `assets\`. The reserved `<vault>\笔记\<Course>` tree is never a publication target.
- Keep catalog, manifests, provenance, retrieval chunks, and runtime metadata under `~/.codex/study-docs` (or `STUDY_DOCS_STATE_DIR`); do not publish them into the vault.
- Use `--migration-plan --json` for a read-only assessment of legacy `<Course>\Sources\<Document>` folders. It reports evidence and refuses ambiguity; it never moves or deletes anything.
- Use `--refresh` only when re-extraction is intentional; ordinary repeat runs must use the compatible cache.
- Use `--mode private` only for supported local PDF extraction. If the local dependency is absent, fail without cloud fallback.
- Never pass or print a token argument. The subprocess inherits `MINERU_TOKEN` privately.

Report per-document status, cache path, publication result, warnings, and whether MinerU ran. A failed item in a batch does not invalidate successful items.

## References

- Read [references/architecture.md](references/architecture.md) when diagnosing identity, revision, commit, or publication behavior.
- Read [references/manifest-schema.md](references/manifest-schema.md) when consuming provenance or validating cache artifacts.
- Read [references/config-schema.md](references/config-schema.md) when configuring a vault, defaults, or source retention.

For a generic one-time request such as “convert this PDF to Markdown,” use the existing `mineru` skill instead. For questions about cached course content, use `study-source-reader`.
