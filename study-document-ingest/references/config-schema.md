# Configuration

Runtime state lives at `~/.codex/study-docs` by default, never in a skill directory.
Set `STUDY_DOCS_STATE_DIR` to use another location.

```json
{
  "schema_version": 1,
  "vault_path": null,
  "default_mode": "quality",
  "default_course_fallback": "_Inbox",
  "source_retention": "reference",
  "publish_enabled": true,
  "publication_root": "课程资料"
}
```

- `vault_path`: existing, explicitly selected Obsidian vault or `null`. A command-line `--vault` applies only to that run.
- `default_mode`: `quality`, `fast`, or `private`.
- `default_course_fallback`: safe destination used when course inference is uncertain.
- `source_retention`: `reference` records absolute original path and hash; `copy` also archives one local copy by content hash under the state directory.
- `publish_enabled`: allows publication when a vault is selected. Cache/catalog operation does not require a vault.
- `publication_root`: must be `课程资料`, the sole visible root for machine-ingested sources. The personal-notes root `笔记` and obsolete alternatives are rejected. Existing schema-v1 config files without this field receive the required default at load time.

For the configured study vault, the clean visible boundary is:

```text
<vault>\
  笔记\<Course>\                  # never written by this pipeline
  课程资料\<Course>\<Document>\
    source.md
    assets\
```

Catalog, cache manifests, provenance, retrieval chunks, normalized PDFs, logs, and other machine state remain under the configured state directory.

Never place `MINERU_TOKEN` or any other secret in this file.
