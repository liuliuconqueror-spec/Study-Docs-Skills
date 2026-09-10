from __future__ import annotations

import importlib.util
import hashlib
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "ingest.py"
SPEC = importlib.util.spec_from_file_location("study_document_ingest", SCRIPT)
assert SPEC and SPEC.loader
ingest = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = ingest
SPEC.loader.exec_module(ingest)

READER_SCRIPT = ROOT.parent / "study-source-reader" / "scripts" / "search.py"
READER_SPEC = importlib.util.spec_from_file_location("study_source_reader_integration", READER_SCRIPT)
assert READER_SPEC and READER_SPEC.loader
reader = importlib.util.module_from_spec(READER_SPEC)
sys.modules[READER_SPEC.name] = reader
READER_SPEC.loader.exec_module(reader)


FAKE_MINERU = r'''#!/usr/bin/env python3
import json
import os
import sys
from pathlib import Path

args = sys.argv[1:]
if "--version" in args:
    print("mineru 3.3.1")
    raise SystemExit(0)
if "--doctor" in args:
    print(json.dumps({
        "version": "3.3.1",
        "python": {"ok": True, "detail": sys.version.split()[0]},
        "network": {"ok": True, "detail": "mock"},
        "token": {"ok": True, "detail": "accepted"},
        "optional_extras": {"pypdf (--split)": True},
        "healthy": True,
    }))
    raise SystemExit(0)

counter = Path(os.environ["FAKE_COUNTER"])
count = int(counter.read_text() or "0") if counter.exists() else 0
counter.write_text(str(count + 1), encoding="utf-8")
record = os.environ.get("FAKE_ENV_RECORD")
if record:
    Path(record).write_text(json.dumps({
        "PYTHONUTF8": os.environ.get("PYTHONUTF8"),
        "PYTHONIOENCODING": os.environ.get("PYTHONIOENCODING"),
    }), encoding="utf-8")
arg_record = os.environ.get("FAKE_ARGS_RECORD")
if arg_record:
    Path(arg_record).write_text(json.dumps(args), encoding="utf-8")

mode = os.environ.get("FAKE_MODE", "ok")
if mode == "fail":
    sys.stderr.write("mock extraction failed")
    raise SystemExit(1)

source = Path(args[0])
output = Path(args[args.index("--output") + 1])
target = output / source.stem
target.mkdir(parents=True, exist_ok=True)
markdown_path = target / (source.stem + ".md")
chunks_path = target / (source.stem + ".chunks.json")
default_text = (
    "# Signals and Systems\n\n"
    "Convolution combines an input with an impulse response. "
    "The Fourier transform describes frequency-domain structure. "
    "This deterministic mock paragraph is intentionally long enough for validation."
)
text = os.environ.get("FAKE_CONTENT", default_text)
if mode == "empty":
    text = ""
if mode == "image":
    text += "\n\n![circuit](images/circuit.png)"
    image = target / "images" / "circuit.png"
    image.parent.mkdir()
    image.write_bytes(b"mock-png")
if mode == "missing_image":
    text += "\n\n![missing](images/missing.png)"
if mode == "invalid_utf8":
    markdown_path.write_bytes(b"\xff\xfe\x00")
else:
    markdown_path.write_text(text, encoding="utf-8")

chunks = [{
    "id": source.stem + "-0",
    "index": 0,
    "heading": "Signals and Systems",
    "text": text,
    "chars": len(text),
    "source": str(source),
}]
if mode != "missing_chunks":
    if mode == "invalid_chunks":
        chunks_path.write_text("{not-json", encoding="utf-8")
    else:
        chunks_path.write_text(json.dumps(chunks, ensure_ascii=False), encoding="utf-8")

status = {
    "total": 1,
    "done": 1,
    "skipped": 0,
    "failed": 0,
    "results": [{
        "name": source.stem,
        "source": str(source),
        "api": "local" if args[args.index("--engine") + 1] == "local" else "standard",
        "modality": source.suffix.lstrip("."),
        "state": "done",
        "output_dir": str(target),
        "markdown_path": str(markdown_path),
        "task_id": "mock-task",
        "elapsed": 0.01,
        "error": None,
        "sinks": [],
    }],
}
if mode == "console_fail":
    sys.stderr.write("UnicodeEncodeError: 'gbk' codec can't encode character")
    raise SystemExit(1)
print(json.dumps(status, ensure_ascii=False))
'''


class IngestionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.state = self.root / "state"
        self.fake = self.root / "fake_mineru.py"
        self.fake.write_text(FAKE_MINERU, encoding="utf-8")
        self.counter = self.root / "counter.txt"
        self.conversion_counter = self.root / "conversion-counter.txt"
        self.conversion_records: list[dict[str, str]] = []
        self.env = mock.patch.dict(
            os.environ,
            {
                "FAKE_COUNTER": str(self.counter),
                "FAKE_MODE": "ok",
                "FAKE_CONTENT": (
                    "# Test material\n\n"
                    "This is a sufficiently detailed mock extraction for repeatable university study retrieval. "
                    "It includes enough text to avoid sparse-output fallback and supports deterministic tests."
                ),
            },
            clear=False,
        )
        self.env.start()
        self.addCleanup(self.env.stop)
        self.powerpoint_backend = ingest.NormalizerBackend(
            "powerpoint", "16.0.test", "POWERPNT.EXE"
        )

        def detect_normalizers():
            return [self.powerpoint_backend]

        def convert_normalized(source, destination, backend):
            from pypdf import PdfWriter

            count = (
                int(self.conversion_counter.read_text(encoding="utf-8"))
                if self.conversion_counter.exists()
                else 0
            )
            self.conversion_counter.write_text(str(count + 1), encoding="utf-8")
            destination.parent.mkdir(parents=True, exist_ok=False)
            writer = PdfWriter()
            writer.add_blank_page(width=612, height=792)
            with destination.open("wb") as handle:
                writer.write(handle)
            self.conversion_records.append(
                {
                    "source": str(source),
                    "destination": str(destination),
                    "backend": backend.name,
                }
            )
            return {"ok": True, "slide_count": 1}

        self.detect_normalizers = detect_normalizers
        self.convert_normalized = convert_normalized
        self.pipeline = ingest.IngestionPipeline(
            state_dir=self.state,
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=self.detect_normalizers,
            normalizer_converter=self.convert_normalized,
        )
        self.course_dir = self.root / "Signals"
        self.course_dir.mkdir()
        self.source = self.course_dir / "lecture.pptx"
        self.source.write_bytes(b"version-one")

    def calls(self) -> int:
        return int(self.counter.read_text()) if self.counter.exists() else 0

    def conversions(self) -> int:
        return (
            int(self.conversion_counter.read_text(encoding="utf-8"))
            if self.conversion_counter.exists()
            else 0
        )

    def test_portable_defaults_use_current_python_and_discovered_mineru(self):
        with mock.patch.object(ingest, "default_mineru_script", return_value=self.fake):
            pipeline = ingest.IngestionPipeline(
                state_dir=self.root / "portable-defaults",
                initialize=False,
            )
        self.assertEqual(pipeline.python_executable, Path(sys.executable).resolve())
        self.assertEqual(pipeline.mineru_script, self.fake.resolve())
        self.assertEqual(pipeline.chunker_script, self.fake.with_name("chunking.py").resolve())

    def run_ingest(self, source: Path | None = None, **kwargs):
        return self.pipeline.ingest_file(
            source or self.source,
            no_publish=kwargs.pop("no_publish", True),
            **kwargs,
        )

    # TEST 1
    def test_same_path_same_bytes_profile_is_cache_hit(self):
        first = self.run_ingest()
        second = self.run_ingest()
        self.assertEqual(first["status"], "INGESTED")
        self.assertEqual(second["status"], "CACHE_HIT")
        self.assertEqual(self.calls(), 1)
        self.assertEqual(self.conversions(), 1)

    # TEST 2
    def test_same_path_modified_bytes_creates_revision(self):
        first = self.run_ingest()
        self.source.write_bytes(b"version-two")
        second = self.run_ingest()
        self.assertNotEqual(first["cache_key"], second["cache_key"])
        self.assertNotEqual(first["revision_id"], second["revision_id"])
        self.assertEqual(first["document_id"], second["document_id"])
        self.assertEqual(self.calls(), 2)

    # TEST 3
    def test_different_filename_identical_bytes_reuses_extraction(self):
        first = self.run_ingest()
        cache_manifest_path = Path(first["cache_path"]) / "manifest.json"
        immutable_manifest = cache_manifest_path.read_bytes()
        renamed = self.course_dir / "renamed-copy.pptx"
        renamed.write_bytes(self.source.read_bytes())
        self.source.unlink()
        second = self.run_ingest(renamed)
        self.assertEqual(first["cache_key"], second["cache_key"])
        self.assertEqual(first["document_id"], second["document_id"])
        self.assertEqual(second["status"], "CACHE_HIT")
        self.assertEqual(cache_manifest_path.read_bytes(), immutable_manifest)
        catalog = ingest.read_json(self.state / "catalog.json")
        document = catalog["documents"][first["document_id"]]
        self.assertEqual(document["current_source_binding"]["path"], str(renamed.resolve()))
        self.assertEqual(document["title"], "lecture")
        lookup = reader.search_catalog(
            catalog_path=self.state / "catalog.json",
            query="Write the exact formula from this source",
        )
        self.assertTrue(lookup["original_available"])
        self.assertEqual(lookup["original_source"], str(renamed.resolve()))
        self.assertEqual(self.calls(), 1)
        self.assertEqual(self.conversions(), 1)

    # TEST 4
    def test_profile_change_is_cache_miss(self):
        first = self.run_ingest(mode="quality")
        second = self.run_ingest(mode="fast")
        self.assertNotEqual(first["cache_key"], second["cache_key"])
        self.assertEqual(self.calls(), 2)

    def test_chunk_size_change_reuses_extraction_and_generates_profiled_chunks(self):
        os.environ["FAKE_MODE"] = "image"
        os.environ["FAKE_CONTENT"] = "# Long material\n\n" + ("retrieval paragraph " * 120)
        args_record = self.root / "args.json"
        os.environ["FAKE_ARGS_RECORD"] = str(args_record)
        first = self.run_ingest(chunk_size=600)
        cache = Path(first["cache_path"])
        source_hash = ingest.sha256_file(cache / "source.md")
        asset_hash = ingest.sha256_file(cache / "assets" / "images" / "circuit.png")

        second = self.run_ingest(chunk_size=300)

        self.assertEqual(self.calls(), 1)
        self.assertEqual(self.conversions(), 1)
        self.assertEqual(second["status"], "CACHE_HIT")
        self.assertEqual(first["cache_key"], second["cache_key"])
        self.assertEqual(
            first["extraction_profile_fingerprint"],
            second["extraction_profile_fingerprint"],
        )
        self.assertNotEqual(
            first["retrieval_profile_fingerprint"],
            second["retrieval_profile_fingerprint"],
        )
        self.assertEqual(ingest.sha256_file(cache / "source.md"), source_hash)
        self.assertEqual(
            ingest.sha256_file(cache / "assets" / "images" / "circuit.png"),
            asset_hash,
        )
        self.assertTrue(Path(first["chunks_path"]).is_file())
        self.assertTrue(Path(second["chunks_path"]).is_file())
        metadata = ingest.read_json(Path(second["chunks_path"]).parent / "retrieval.json")
        self.assertEqual(metadata["retrieval_profile"]["chunk_size"], 300)
        self.assertEqual(
            metadata["retrieval_profile_fingerprint"],
            second["retrieval_profile_fingerprint"],
        )
        chunks = ingest.read_json(Path(second["chunks_path"]))
        self.assertTrue(all(chunk["chars"] <= 300 for chunk in chunks))
        mineru_args = json.loads(args_record.read_text(encoding="utf-8"))
        self.assertNotIn("--chunk", mineru_args)
        self.assertNotIn("--chunk-size", mineru_args)
        self.assertEqual(Path(mineru_args[0]).suffix.casefold(), ".pdf")

    # TEST 5
    def test_refresh_invokes_mineru(self):
        self.run_ingest()
        refreshed = self.run_ingest(refresh=True)
        self.assertEqual(refreshed["status"], "INGESTED")
        self.assertTrue(refreshed["refresh"])
        self.assertEqual(self.calls(), 2)

    # TEST 6
    def test_failed_new_extraction_preserves_old_cache_and_publication(self):
        vault = self.root / "vault"
        (vault / ".obsidian").mkdir(parents=True)
        first = self.pipeline.ingest_file(self.source, vault=vault)
        published = Path(first["published_path"]) / "source.md"
        old_text = published.read_text(encoding="utf-8")
        old_catalog = ingest.read_json(self.state / "catalog.json")
        self.source.write_bytes(b"changed-source")
        os.environ["FAKE_MODE"] = "fail"
        with self.assertRaises(ingest.PipelineError):
            self.pipeline.ingest_file(self.source, vault=vault)
        self.assertEqual(published.read_text(encoding="utf-8"), old_text)
        current = ingest.read_json(self.state / "catalog.json")
        self.assertEqual(current, old_catalog)
        self.assertTrue(Path(first["cache_path"]).is_dir())

    # TEST 7
    def test_empty_markdown_fails(self):
        root = self.root / "empty"
        root.mkdir()
        (root / "source.md").write_text("", encoding="utf-8")
        (root / "chunks.json").write_text("[]", encoding="utf-8")
        with self.assertRaises(ingest.ValidationError):
            ingest.validate_artifacts(root)

    # TEST 8
    def test_invalid_utf8_markdown_fails(self):
        root = self.root / "bad-utf8"
        root.mkdir()
        (root / "source.md").write_bytes(b"\xff\xfe")
        (root / "chunks.json").write_text("[]", encoding="utf-8")
        with self.assertRaises(ingest.ValidationError):
            ingest.validate_artifacts(root)

    # TEST 9
    def test_missing_chunks_fails(self):
        root = self.root / "missing-chunks"
        root.mkdir()
        (root / "source.md").write_text("x" * 100, encoding="utf-8")
        with self.assertRaises(ingest.ValidationError):
            ingest.validate_artifacts(root)

    # TEST 10
    def test_invalid_chunks_json_fails(self):
        root = self.root / "invalid-chunks"
        root.mkdir()
        (root / "source.md").write_text("x" * 100, encoding="utf-8")
        (root / "chunks.json").write_text("{bad", encoding="utf-8")
        with self.assertRaises(ingest.ValidationError):
            ingest.validate_artifacts(root)

    # TEST 11
    def test_valid_local_image_reference_succeeds(self):
        os.environ["FAKE_MODE"] = "image"
        result = self.run_ingest()
        manifest = ingest.read_json(Path(result["cache_path"]) / "manifest.json")
        self.assertEqual(manifest["validation"]["image_references"], 1)
        self.assertEqual(manifest["validation"]["resolved_images"], 1)
        self.assertTrue((Path(result["cache_path"]) / "assets" / "images" / "circuit.png").is_file())

    # TEST 12
    def test_single_missing_image_is_warning(self):
        os.environ["FAKE_MODE"] = "missing_image"
        result = self.run_ingest()
        self.assertTrue(any("missing_image_references" in warning for warning in result["warnings"]))

    # TEST 13
    def test_unsafe_course_and_title_are_sanitized(self):
        result = self.run_ingest(course="../CON", title="..\\AUX/evil")
        for value in (result["course"], result["title"]):
            self.assertNotIn("..", value)
            self.assertNotIn("/", value)
            self.assertNotIn("\\", value)

    # TEST 14
    def test_windows_reserved_names_are_safe(self):
        self.assertEqual(ingest.safe_component("CON"), "_CON")
        self.assertEqual(ingest.safe_component("lpt1"), "_lpt1")
        self.assertFalse(ingest.safe_component("name. ").endswith((".", " ")))

    # TEST 15
    def test_manifest_required_fields_and_hashes(self):
        original_hash = ingest.sha256_file(self.source)
        result = self.run_ingest()
        cache = Path(result["cache_path"])
        manifest = ingest.read_json(cache / "manifest.json")
        ingest.validate_manifest(manifest, cache)
        self.assertEqual(manifest["source_sha256"], original_hash)
        self.assertEqual(manifest["original_source"]["source_sha256"], original_hash)
        self.assertEqual(manifest["original_source"]["role"], "canonical_original")
        self.assertEqual(manifest["normalized_source"]["role"], "extraction_derivative")
        normalized_pdf = cache / manifest["normalized_source"]["path"]
        self.assertTrue(normalized_pdf.is_file())
        self.assertEqual(
            manifest["normalized_pdf_sha256"], ingest.sha256_file(normalized_pdf)
        )
        self.assertEqual(manifest["normalized_source"]["backend"], "powerpoint")
        self.assertEqual(manifest["profile"], manifest["extraction_profile"])
        self.assertNotIn("chunk_size", manifest["extraction_profile"])
        self.assertEqual(manifest["extraction_profile"]["input_modality"], "pptx")
        self.assertEqual(manifest["extraction_profile"]["normalization"], "pptx_to_pdf")
        self.assertNotIn("chunks.json", manifest["artifacts"])
        self.assertIn("normalized/source.pdf", manifest["artifacts"])
        self.assertEqual(
            manifest["cache_key"],
            ingest.extraction_cache_key(
                manifest["source_sha256"], manifest["extraction_profile_fingerprint"]
            ),
        )

    # TEST 16
    def test_atomic_write_failure_does_not_corrupt_old_catalog(self):
        path = self.root / "atomic.json"
        path.write_text('{"old": true}\n', encoding="utf-8")
        with mock.patch.object(ingest.os, "replace", side_effect=OSError("simulated")):
            with self.assertRaises(OSError):
                ingest.atomic_write_json(path, {"new": True})
        self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"old": True})

    # TEST 17
    def test_private_mode_cannot_construct_cloud_command(self):
        private = ingest.IngestionPipeline(
            state_dir=self.root / "private-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            local_engine_available_override=True,
        )
        pdf = self.root / "private.pdf"
        pdf.write_bytes(b"not-a-real-pdf")
        profile = ingest.profile_for_mode("private")
        command = private.build_mineru_command(pdf, self.root / "out", profile)
        self.assertEqual(command[command.index("--engine") + 1], "local")
        self.assertNotIn("--api", command)
        self.assertNotIn("cloud", command)
        with self.assertRaises(ingest.PrivateModeError):
            private.build_mineru_command(self.source, self.root / "out2", profile)

    # TEST 18
    def test_generic_parent_falls_back_to_inbox(self):
        generic = self.root / "Downloads" / "lecture.pptx"
        generic.parent.mkdir()
        generic.write_bytes(b"x")
        self.assertEqual(ingest.infer_course(generic), "_Inbox")

    # TEST 19
    def test_logical_update_preserves_previous_revision_cache(self):
        first = self.run_ingest()
        first_cache = Path(first["cache_path"])
        self.source.write_bytes(b"revision-two")
        second = self.run_ingest()
        catalog = ingest.read_json(self.state / "catalog.json")
        document = catalog["documents"][first["document_id"]]
        self.assertEqual(document["current_revision_id"], second["revision_id"])
        self.assertEqual(len(document["revisions"]), 2)
        self.assertTrue(first_cache.is_dir())

    # TEST 20
    def test_publication_swap_failure_restores_previous_version(self):
        vault = self.root / "vault-rollback"
        (vault / ".obsidian").mkdir(parents=True)
        os.environ["FAKE_CONTENT"] = "# Old\n\n" + ("old content " * 20)
        first = self.pipeline.ingest_file(self.source, vault=vault)
        published = Path(first["published_path"]) / "source.md"
        old = published.read_text(encoding="utf-8")
        old_catalog = ingest.read_json(self.state / "catalog.json")
        self.source.write_bytes(b"revision-for-publication-failure")
        os.environ["FAKE_CONTENT"] = "# New\n\n" + ("new content " * 20)

        def fail(stage: str):
            if stage == "after_source_swap":
                raise OSError("simulated metadata swap failure")

        with self.assertRaises(ingest.PublicationError):
            self.pipeline.ingest_file(
                self.source,
                vault=vault,
                publication_fault_injector=fail,
            )
        self.assertEqual(published.read_text(encoding="utf-8"), old)
        self.assertEqual(ingest.read_json(self.state / "catalog.json"), old_catalog)

    # TEST 21
    def test_mineru_child_receives_utf8_environment(self):
        record = self.root / "env.json"
        os.environ["FAKE_ENV_RECORD"] = str(record)
        self.run_ingest()
        environment = json.loads(record.read_text(encoding="utf-8"))
        self.assertEqual(environment["PYTHONUTF8"], "1")
        self.assertEqual(environment["PYTHONIOENCODING"], "utf-8")

    # TEST 22
    def test_console_status_failure_after_valid_artifacts_is_accepted(self):
        os.environ["FAKE_MODE"] = "console_fail"
        result = self.run_ingest()
        self.assertEqual(result["status"], "INGESTED")
        self.assertTrue(any("console-status" in warning for warning in result["warnings"]))

    # TEST 23
    def test_same_title_different_documents_do_not_overwrite(self):
        vault = self.root / "collision-vault"
        (vault / ".obsidian").mkdir(parents=True)
        first_file = self.course_dir / "one.pptx"
        second_file = self.course_dir / "two.pptx"
        first_file.write_bytes(b"one")
        second_file.write_bytes(b"two")
        first = self.pipeline.ingest_file(first_file, course="Course", title="Same", vault=vault)
        second = self.pipeline.ingest_file(second_file, course="Course", title="Same", vault=vault)
        self.assertNotEqual(first["document_id"], second["document_id"])
        self.assertNotEqual(first["published_path"], second["published_path"])
        self.assertTrue((Path(first["published_path"]) / "source.md").is_file())
        self.assertTrue((Path(second["published_path"]) / "source.md").is_file())

    # TEST 24
    def test_filename_change_does_not_change_content_cache(self):
        first = self.run_ingest()
        moved_dir = self.root / "Signals" / "Archive"
        moved_dir.mkdir()
        moved = moved_dir / "different-name.pptx"
        moved.write_bytes(self.source.read_bytes())
        second = self.run_ingest(moved, course="Signals")
        self.assertEqual(first["cache_key"], second["cache_key"])
        self.assertEqual(self.calls(), 1)

    def test_offline_integration_cache_revision_publication_and_rollback(self):
        vault = self.root / "offline-vault"
        (vault / ".obsidian").mkdir(parents=True)
        os.environ["FAKE_MODE"] = "image"
        first = self.pipeline.ingest_file(self.source, vault=vault)
        second = self.pipeline.ingest_file(self.source, vault=vault)
        self.assertEqual(second["status"], "CACHE_HIT")
        self.assertEqual(self.calls(), 1)
        self.source.write_bytes(b"offline-new-revision")
        updated = self.pipeline.ingest_file(self.source, vault=vault)
        self.assertNotEqual(first["revision_id"], updated["revision_id"])
        before_failure = Path(updated["published_path"], "source.md").read_text(encoding="utf-8")
        self.source.write_bytes(b"offline-failed-revision")
        os.environ["FAKE_MODE"] = "fail"
        with self.assertRaises(ingest.PipelineError):
            self.pipeline.ingest_file(self.source, vault=vault)
        self.assertEqual(
            Path(updated["published_path"], "source.md").read_text(encoding="utf-8"),
            before_failure,
        )

    def test_dry_run_is_read_only_and_does_not_call_mineru(self):
        dry_state = self.root / "never-created-state"
        pipeline = ingest.IngestionPipeline(
            state_dir=dry_state,
            python_executable=sys.executable,
            mineru_script=self.fake,
            initialize=False,
            normalizer_detector=self.detect_normalizers,
            normalizer_converter=self.convert_normalized,
        )
        result = pipeline.ingest_file(self.source, dry_run=True)
        self.assertEqual(result["status"], "DRY_RUN")
        self.assertTrue(result["would_run_mineru"])
        self.assertFalse(dry_state.exists())
        self.assertEqual(self.calls(), 0)

    def test_directory_expansion_skips_generated_and_temporary_files(self):
        root = self.root / "batch"
        (root / "Course").mkdir(parents=True)
        (root / "Course" / "one.pdf").write_bytes(b"one")
        (root / "Course" / "~$draft.pptx").write_bytes(b"temp")
        (root / "cache").mkdir()
        (root / "cache" / "generated.pdf").write_bytes(b"generated")
        selected = ingest.expand_inputs([str(root)], recursive=True)
        self.assertEqual([path.name for path in selected], ["one.pdf"])

    def test_source_retention_copy_is_content_addressed(self):
        result = self.run_ingest(source_retention="copy")
        manifest = ingest.read_json(Path(result["cache_path"]) / "manifest.json")
        retained = Path(manifest["source_retained_path"])
        self.assertTrue(retained.is_file())
        self.assertEqual(ingest.sha256_file(retained), ingest.sha256_file(self.source))

    def test_valid_v1_cache_migrates_without_mineru_reextraction(self):
        legacy_source = self.course_dir / "legacy.pdf"
        from pypdf import PdfWriter

        writer = PdfWriter()
        writer.add_blank_page(width=612, height=792)
        with legacy_source.open("wb") as handle:
            writer.write(handle)
        first = self.run_ingest(legacy_source, chunk_size=600)
        new_cache = Path(first["cache_path"])
        combined_profile = ingest.profile_for_mode("quality", chunk_size=600).to_dict()
        combined_profile["pipeline_schema_version"] = ingest.LEGACY_PIPELINE_SCHEMA_VERSION
        legacy_fingerprint = ingest.legacy_profile_fingerprint(combined_profile)
        legacy_key = ingest.legacy_extraction_cache_key(
            ingest.sha256_file(legacy_source), legacy_fingerprint
        )
        legacy_cache = self.state / "cache" / legacy_key
        legacy_cache.mkdir()
        shutil.copy2(new_cache / "source.md", legacy_cache / "source.md")
        shutil.copytree(new_cache / "assets", legacy_cache / "assets")
        shutil.copy2(first["chunks_path"], legacy_cache / "chunks.json")
        shutil.copy2(new_cache / "extraction.json", legacy_cache / "extraction.json")
        manifest = ingest.read_json(new_cache / "manifest.json")
        manifest["schema_version"] = ingest.LEGACY_PIPELINE_SCHEMA_VERSION
        manifest["cache_key"] = legacy_key
        manifest["revision_id"] = ingest.revision_id_for(legacy_key)
        manifest["profile"] = combined_profile
        manifest["profile_fingerprint"] = legacy_fingerprint
        manifest["chunk_enabled"] = True
        manifest["chunk_size"] = 600
        manifest["chunks_path"] = "chunks.json"
        manifest.pop("extraction_profile", None)
        manifest.pop("extraction_profile_fingerprint", None)
        manifest["artifacts"] = {}
        for path in sorted(legacy_cache.rglob("*")):
            if path.is_file() and path.name not in {"manifest.json", "extraction.json"}:
                relative = path.relative_to(legacy_cache).as_posix()
                manifest["artifacts"][relative] = {
                    "sha256": ingest.sha256_file(path),
                    "size_bytes": path.stat().st_size,
                }
        ingest.atomic_write_json(legacy_cache / "manifest.json", manifest)
        shutil.rmtree(new_cache)

        migrated = self.run_ingest(legacy_source, chunk_size=300)

        self.assertEqual(migrated["status"], "CACHE_HIT")
        self.assertEqual(self.calls(), 1)
        self.assertTrue(Path(migrated["cache_path"], "source.md").is_file())
        self.assertTrue(Path(migrated["chunks_path"]).is_file())
        self.assertTrue(
            any("legacy_v1_cache_migrated_without_mineru" in warning for warning in migrated["warnings"])
        )

    def test_private_mode_missing_local_dependency_fails_before_subprocess(self):
        private = ingest.IngestionPipeline(
            state_dir=self.root / "private-missing-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            local_engine_available_override=False,
        )
        pdf = self.root / "private-missing.pdf"
        pdf.write_bytes(b"private")
        with self.assertRaises(ingest.PrivateModeError):
            private.ingest_file(pdf, mode="private", no_publish=True)
        self.assertEqual(self.calls(), 0)

    def test_oversized_pdf_uses_mineru_native_split_flag(self):
        pdf = self.course_dir / "large.pdf"
        pdf.write_bytes(b"mock-large")
        args_record = self.root / "args.json"
        os.environ["FAKE_ARGS_RECORD"] = str(args_record)
        with mock.patch.object(ingest, "pdf_page_count", return_value=461), mock.patch.object(
            ingest, "load_mineru_limits", return_value={"agent": 20, "standard": 200}
        ):
            self.pipeline.ingest_file(pdf, mode="quality", no_publish=True)
        args = json.loads(args_record.read_text(encoding="utf-8"))
        self.assertIn("--split", args)

    def test_logs_redact_known_token_value(self):
        secret = "unit-test-secret-token"
        os.environ["MINERU_TOKEN"] = secret
        self.pipeline.logger.log("redaction_test", error=f"failure included {secret}")
        logs = "\n".join(
            path.read_text(encoding="utf-8") for path in (self.state / "logs").glob("*.jsonl")
        )
        self.assertNotIn(secret, logs)
        self.assertIn("[REDACTED]", logs)

    def test_quality_pptx_prefers_powerpoint_over_libreoffice(self):
        libreoffice = ingest.NormalizerBackend(
            "libreoffice", "26.2.test", "soffice.com"
        )
        pipeline = ingest.IngestionPipeline(
            state_dir=self.root / "priority-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=lambda: [libreoffice, self.powerpoint_backend],
            normalizer_converter=self.convert_normalized,
        )

        result = pipeline.ingest_file(self.source, mode="quality", no_publish=True)

        self.assertEqual(result["normalizer_backend"], "powerpoint")
        self.assertEqual(self.conversion_records[-1]["backend"], "powerpoint")

    def test_quality_pptx_uses_libreoffice_when_powerpoint_is_unavailable(self):
        libreoffice = ingest.NormalizerBackend(
            "libreoffice", "26.2.test", "soffice.com"
        )
        pipeline = ingest.IngestionPipeline(
            state_dir=self.root / "libreoffice-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=lambda: [libreoffice],
            normalizer_converter=self.convert_normalized,
        )

        result = pipeline.ingest_file(self.source, mode="quality", no_publish=True)

        self.assertEqual(result["normalizer_backend"], "libreoffice")
        self.assertEqual(self.conversion_records[-1]["backend"], "libreoffice")

    def test_no_pptx_normalizer_fails_before_mineru_or_catalog_advance(self):
        pipeline = ingest.IngestionPipeline(
            state_dir=self.root / "no-normalizer-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=lambda: [],
        )

        with self.assertRaisesRegex(ingest.NormalizationError, "neither Microsoft PowerPoint"):
            pipeline.ingest_file(self.source, mode="quality", no_publish=True)

        self.assertEqual(self.calls(), 0)
        self.assertEqual(pipeline.catalog()["documents"], {})

    def test_original_pptx_is_unchanged_by_normalization(self):
        before = self.source.read_bytes()
        before_hash = hashlib.sha256(before).hexdigest()

        self.run_ingest()

        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(ingest.sha256_file(self.source), before_hash)

    def test_empty_or_invalid_normalized_pdf_fails_before_mineru(self):
        for label, payload in (("empty", b""), ("invalid", b"not-a-pdf")):
            with self.subTest(label=label):
                def invalid_converter(source, destination, backend, data=payload):
                    destination.parent.mkdir(parents=True, exist_ok=False)
                    destination.write_bytes(data)
                    return {"ok": True, "slide_count": 1}

                pipeline = ingest.IngestionPipeline(
                    state_dir=self.root / f"{label}-pdf-state",
                    python_executable=sys.executable,
                    mineru_script=self.fake,
                    normalizer_detector=self.detect_normalizers,
                    normalizer_converter=invalid_converter,
                )
                with self.assertRaises(ingest.NormalizationError):
                    pipeline.ingest_file(self.source, mode="quality", no_publish=True)
        self.assertEqual(self.calls(), 0)

    def test_implausible_slide_page_mismatch_fails_before_mineru(self):
        pipeline = ingest.IngestionPipeline(
            state_dir=self.root / "mismatch-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=self.detect_normalizers,
            normalizer_converter=self.convert_normalized,
        )

        with mock.patch.object(ingest, "pptx_slide_count", return_value=12):
            with self.assertRaisesRegex(ingest.NormalizationError, "12 slides versus 1 pages"):
                pipeline.ingest_file(self.source, mode="quality", no_publish=True)

        self.assertEqual(self.calls(), 0)

    def test_catastrophic_image_loss_remains_fatal_after_pptx_normalization(self):
        references = "\n".join(
            f"![missing-{index}](images/missing-{index}.png)" for index in range(12)
        )
        os.environ["FAKE_CONTENT"] = (
            "# Visual lecture\n\n"
            + ("Validated text remains present. " * 10)
            + "\n\n"
            + references
        )

        with self.assertRaisesRegex(ingest.ValidationError, "catastrophic image-reference loss"):
            self.run_ingest()

        self.assertEqual(self.calls(), 1)
        self.assertEqual(self.pipeline.catalog()["documents"], {})

    def test_normalized_pdf_is_cache_only_not_human_facing_source(self):
        vault = self.root / "derivative-vault"
        (vault / ".obsidian").mkdir(parents=True)

        result = self.pipeline.ingest_file(self.source, vault=vault)

        published = Path(result["published_path"])
        self.assertFalse((published / "normalized").exists())
        self.assertTrue((Path(result["cache_path"]) / "normalized" / "source.pdf").is_file())

    def test_powerpoint_helper_is_copy_only_and_never_kills_processes(self):
        helper = ROOT / "scripts" / "export_pptx_pdf.ps1"
        source = helper.read_text(encoding="utf-8")

        self.assertIn("SaveCopyAs", source)
        self.assertIn("$true, $false, $false", source)
        self.assertNotIn("Stop-Process", source)
        self.assertNotIn(".Kill(", source)

    def test_normalizer_backend_and_version_are_extraction_identity(self):
        base = ingest.profile_for_mode("quality", chunk_size=600)
        powerpoint = self.pipeline._quality_pptx_profile(
            base, ingest.NormalizerBackend("powerpoint", "16.0")
        )
        libreoffice = self.pipeline._quality_pptx_profile(
            base, ingest.NormalizerBackend("libreoffice", "26.2")
        )
        powerpoint_other_version = self.pipeline._quality_pptx_profile(
            base, ingest.NormalizerBackend("powerpoint", "16.1")
        )

        self.assertNotEqual(
            ingest.extraction_profile_fingerprint(powerpoint),
            ingest.extraction_profile_fingerprint(libreoffice),
        )
        self.assertNotEqual(
            ingest.extraction_profile_fingerprint(powerpoint),
            ingest.extraction_profile_fingerprint(powerpoint_other_version),
        )
        smaller_chunks = ingest.dataclasses.replace(powerpoint, chunk_size=300)
        self.assertEqual(
            ingest.extraction_profile_fingerprint(powerpoint),
            ingest.extraction_profile_fingerprint(smaller_chunks),
        )

    def test_fast_mode_retains_direct_pptx_mineru_support(self):
        args_record = self.root / "fast-args.json"
        os.environ["FAKE_ARGS_RECORD"] = str(args_record)
        pipeline = ingest.IngestionPipeline(
            state_dir=self.root / "fast-direct-state",
            python_executable=sys.executable,
            mineru_script=self.fake,
            normalizer_detector=lambda: self.fail("normalizer detection should not run"),
            normalizer_converter=lambda *_: self.fail("normalizer conversion should not run"),
        )

        result = pipeline.ingest_file(self.source, mode="fast", no_publish=True)

        args = json.loads(args_record.read_text(encoding="utf-8"))
        self.assertEqual(Path(args[0]).resolve(), self.source.resolve())
        self.assertEqual(result["normalization_cache"], "not_applicable")

    def test_publication_is_human_only_under_course_library_and_never_notes(self):
        vault = self.root / "clean-vault"
        (vault / ".obsidian").mkdir(parents=True)
        notes = vault / ingest.RESERVED_NOTES_ROOT
        notes.mkdir()
        personal_note = notes / "personal.md"
        personal_note.write_text("my own note", encoding="utf-8")
        personal_before = personal_note.read_bytes()
        os.environ["FAKE_MODE"] = "image"

        result = self.pipeline.ingest_file(self.source, vault=vault)

        published = Path(result["published_path"])
        self.assertEqual(
            published,
            vault
            / ingest.DEFAULT_PUBLICATION_ROOT
            / "Signals"
            / "lecture",
        )
        self.assertEqual(
            sorted(child.name for child in published.iterdir()),
            ["assets", "source.md"],
        )
        self.assertFalse(any(vault.rglob("chunks.json")))
        self.assertFalse(any(vault.rglob("manifest.json")))
        self.assertFalse(any(path.name == ".study-docs" for path in vault.rglob("*")))
        self.assertEqual(personal_note.read_bytes(), personal_before)
        self.assertTrue(Path(result["chunks_path"]).is_relative_to(self.state))

    def test_reserved_personal_notes_root_is_rejected_by_config_validation(self):
        config = ingest.default_config()
        config["publication_root"] = ingest.RESERVED_NOTES_ROOT

        with self.assertRaisesRegex(ingest.ValidationError, "reserved personal-notes"):
            ingest.validate_config(config)

    def test_obsolete_publication_root_is_rejected(self):
        config = ingest.default_config()
        config["publication_root"] = "课程资料库"

        with self.assertRaisesRegex(ingest.ValidationError, "must be 课程资料"):
            ingest.validate_config(config)

    def test_user_authored_destination_collision_is_not_overwritten(self):
        vault = self.root / "user-collision-vault"
        (vault / ".obsidian").mkdir(parents=True)
        user_directory = (
            vault / ingest.DEFAULT_PUBLICATION_ROOT / "Course" / "Same"
        )
        user_directory.mkdir(parents=True)
        user_note = user_directory / "source.md"
        user_note.write_text("user-authored source-shaped note", encoding="utf-8")
        before = user_note.read_bytes()

        result = self.pipeline.ingest_file(
            self.source,
            course="Course",
            title="Same",
            vault=vault,
        )

        self.assertNotEqual(Path(result["published_path"]), user_directory)
        self.assertEqual(user_note.read_bytes(), before)
        self.assertTrue(Path(result["published_path"], "source.md").is_file())

    def test_user_modified_known_publication_is_not_overwritten(self):
        vault = self.root / "modified-publication-vault"
        (vault / ".obsidian").mkdir(parents=True)
        first = self.pipeline.ingest_file(self.source, vault=vault)
        published = Path(first["published_path"])
        user_note = published / "my-annotation.md"
        user_note.write_text("keep this", encoding="utf-8")
        before = user_note.read_bytes()

        with self.assertRaisesRegex(ingest.PublicationError, "user-modified"):
            self.pipeline.ingest_file(self.source, vault=vault)

        self.assertEqual(user_note.read_bytes(), before)
        self.assertEqual(self.calls(), 1)

    def test_reader_resolves_cache_after_course_library_publication(self):
        vault = self.root / "reader-vault"
        (vault / ".obsidian").mkdir(parents=True)
        result = self.pipeline.ingest_file(self.source, vault=vault)

        found = reader.search_catalog(
            catalog_path=self.state / "catalog.json",
            query="Explain repeatable university study retrieval",
        )

        self.assertEqual(found["status"], "FOUND")
        self.assertEqual(
            Path(found["source_markdown"]),
            Path(result["cache_path"]) / "source.md",
        )
        self.assertTrue(found["original_available"])
        self.assertFalse(found["mineru_invoked"])

    def test_migration_plan_refuses_ambiguous_legacy_folder_without_changes(self):
        vault = self.root / "ambiguous-vault"
        (vault / ".obsidian").mkdir(parents=True)
        legacy = vault / "Course" / "Sources" / "Unowned"
        legacy.mkdir(parents=True)
        source = legacy / "source.md"
        source.write_text("possibly user authored", encoding="utf-8")
        before = source.read_bytes()

        plan = ingest.plan_legacy_publication_migration(
            vault=vault,
            catalog=ingest.default_catalog(),
        )

        self.assertEqual(plan["status"], "PLAN_ONLY")
        self.assertEqual(plan["safe_count"], 0)
        self.assertEqual(plan["refused_count"], 1)
        self.assertEqual(plan["entries"][0]["status"], "REFUSED_AMBIGUOUS")
        self.assertIn("catalog_ownership_not_unique", plan["entries"][0]["reasons"])
        self.assertFalse(plan["operation_performed"])
        self.assertEqual(source.read_bytes(), before)
        self.assertFalse((vault / ingest.DEFAULT_PUBLICATION_ROOT).exists())

    def test_migration_plan_accepts_only_fully_proven_legacy_publication(self):
        result = self.run_ingest()
        cache = Path(result["cache_path"])
        vault = self.root / "proven-vault"
        (vault / ".obsidian").mkdir(parents=True)
        legacy = vault / "Signals" / "Sources" / "lecture"
        legacy.mkdir(parents=True)
        shutil.copy2(cache / "source.md", legacy / "source.md")
        shutil.copytree(cache / "assets", legacy / "assets")
        metadata = vault / "Signals" / ".study-docs" / result["document_id"]
        metadata.mkdir(parents=True)
        ingest.atomic_write_json(
            metadata / "manifest.json",
            {
                "document_id": result["document_id"],
                "revision_id": result["revision_id"],
                "cache_key": result["cache_key"],
                "published_path": str(legacy),
            },
        )
        catalog = ingest.read_json(self.state / "catalog.json")
        document = catalog["documents"][result["document_id"]]
        document["current_published_revision_id"] = result["revision_id"]
        document["publication"] = {
            "source_path": str(legacy),
            "metadata_path": str(metadata),
            "directory_name": "lecture",
        }

        plan = ingest.plan_legacy_publication_migration(
            vault=vault,
            catalog=catalog,
        )

        self.assertEqual(plan["safe_count"], 1)
        self.assertEqual(plan["refused_count"], 0)
        entry = plan["entries"][0]
        self.assertEqual(entry["status"], "SAFE_TO_MIGRATE")
        self.assertEqual(
            Path(entry["destination_path"]),
            vault / ingest.DEFAULT_PUBLICATION_ROOT / "Signals" / "lecture",
        )
        self.assertFalse(plan["operation_performed"])
        self.assertTrue((legacy / "source.md").is_file())
        self.assertFalse(Path(entry["destination_path"]).exists())


if __name__ == "__main__":
    unittest.main()
