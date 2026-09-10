---
name: study-source-reader
description: Locate and retrieve from previously ingested university course sources using the local study-docs catalog, MinerU chunks, Markdown, and preserved visuals, with original-file verification for exact notation. Use for questions about what the user's ingested lecture, textbook, or course source says. Never parse or re-ingest already-cached material.
---

# Study Source Reader

Answer from durable study sources without rerunning MinerU.

## Retrieval workflow

1. Search the catalog and MinerU-generated chunks with the deterministic helper:

```powershell
python "$HOME\.agents\skills\study-source-reader\scripts\search.py" `
  '<QUERY>' [--course '<COURSE>'] [--title '<TITLE>'] --json
```

If the skills were installed elsewhere, run the same script from that location. The
reader uses `STUDY_DOCS_STATE_DIR` when the state directory is not `~/.codex/study-docs`.

2. If the result is `AMBIGUOUS`, use course/title/current context to disambiguate. Do not silently select a plausible-looking document.
3. Read only the returned chunks and nearby region of `source.md`; keep genuine heading/page/location metadata and never invent page numbers.
4. For conceptual questions, answer from chunks/Markdown unless the retrieved text itself is uncertain.
5. For exact formulas, equation numbers, Greek symbols, signs, subscripts/superscripts, exact table values, circuit diagrams, waveforms, spectra, labels, or visual layout, inspect nearby extracted assets and then a validated retained normalized PDF when present. For PPT-specific facts—animation/build order, hidden slides, transitions, speaker notes, editable structure, or placement not preserved by PDF—inspect the current same-hash original PPTX instead.
6. If the original is missing, state that exact visual verification is unavailable. Do not reconstruct exact notation as fact.

This skill must not call MinerU. If no compatible persistent source exists and ingestion is wanted, route the new source to `study-document-ingest`. For one-off generic conversion, use `mineru`.

The reader uses the external catalog/cache and is independent of the human-facing Obsidian publication copy. Publications under `<vault>\课程资料\<Course>\<Document>` do not change lookup; `<vault>\笔记\<Course>` remains personal-note space and is not searched as an ingested-source cache.

Read [references/retrieval-policy.md](references/retrieval-policy.md) for ranking, ambiguity, and precision-verification details.
