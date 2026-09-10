from __future__ import annotations

import importlib.util
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "search.py"
SPEC = importlib.util.spec_from_file_location("study_source_reader", SCRIPT)
assert SPEC and SPEC.loader
reader = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = reader
SPEC.loader.exec_module(reader)


class ReaderTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.original = self.root / "textbook.pdf"
        self.original.write_bytes(b"source")
        self.source_sha = hashlib.sha256(b"source").hexdigest()
        self.cache = self.root / "cache" / "key-one"
        (self.cache / "assets").mkdir(parents=True)
        (self.cache / "source.md").write_text(
            "# Convolution\n\nConvolution combines x(t) and h(t).\n\n# Fourier Transform\n\nFrequency representation.",
            encoding="utf-8",
        )
        chunks = [
            {
                "id": "doc-0",
                "index": 0,
                "heading": "Convolution",
                "text": "Convolution combines x(t) and h(t).",
                "chars": 38,
                "source": str(self.original),
            },
            {
                "id": "doc-1",
                "index": 1,
                "heading": "Fourier Transform",
                "text": "Frequency representation.",
                "chars": 25,
                "source": str(self.original),
            },
        ]
        (self.cache / "chunks.json").write_text(
            json.dumps(chunks), encoding="utf-8"
        )
        (self.cache / "manifest.json").write_text(
            json.dumps(
                {
                    "source_original_path": str(self.original),
                    "source_retained_path": None,
                }
            ),
            encoding="utf-8",
        )
        self.catalog_path = self.root / "catalog.json"
        self.write_catalog(one_document=True)

    def document(self, document_id: str, title: str, cache: Path, original: Path):
        binding = {
            "path": str(original),
            "path_key": str(original).casefold(),
            "filename": original.name,
            "source_sha256": self.source_sha,
            "last_seen_at": "2026-01-01T00:00:00Z",
        }
        return {
            "document_id": document_id,
            "course": "Signals",
            "title": title,
            "source_aliases": [binding],
            "current_source_binding": dict(binding),
            "current_revision_id": "rev-one",
            "revisions": [
                {
                    "revision_id": "rev-one",
                    "cache_path": str(cache),
                    "source_sha256": self.source_sha,
                }
            ],
        }

    def write_catalog(self, *, one_document: bool):
        documents = {
            "doc-one": self.document("doc-one", "Core Lecture", self.cache, self.original)
        }
        if not one_document:
            second_cache = self.root / "cache" / "key-two"
            second_cache.mkdir(parents=True)
            for name in ("source.md", "chunks.json", "manifest.json"):
                (second_cache / name).write_bytes((self.cache / name).read_bytes())
            second_original = self.root / "other.pdf"
            second_original.write_bytes(b"other")
            documents["doc-two"] = self.document(
                "doc-two", "Alternate Lecture", second_cache, second_original
            )
        self.catalog_path.write_text(
            json.dumps({"schema_version": 1, "documents": documents, "cache_entries": {}}),
            encoding="utf-8",
        )

    def configure_pptx_derivative(self) -> tuple[Path, Path]:
        pptx = self.root / "lecture.pptx"
        pptx.write_bytes(self.original.read_bytes())
        self.original.unlink()
        normalized = self.cache / "normalized" / "source.pdf"
        normalized.parent.mkdir()
        normalized.write_bytes(b"validated-normalized-pdf")
        manifest = json.loads((self.cache / "manifest.json").read_text(encoding="utf-8"))
        manifest.update(
            {
                "source_filename": pptx.name,
                "source_original_path": str(pptx),
                "original_source": {
                    "role": "canonical_original",
                    "filename": pptx.name,
                    "path_at_extraction": str(pptx),
                    "source_sha256": self.source_sha,
                },
                "normalized_pdf_sha256": hashlib.sha256(normalized.read_bytes()).hexdigest(),
                "normalized_source": {
                    "role": "extraction_derivative",
                    "path": "normalized/source.pdf",
                    "sha256": hashlib.sha256(normalized.read_bytes()).hexdigest(),
                    "backend": "powerpoint",
                },
            }
        )
        (self.cache / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
        catalog = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        document = catalog["documents"]["doc-one"]
        binding = {
            "path": str(pptx),
            "path_key": str(pptx).casefold(),
            "filename": pptx.name,
            "source_sha256": self.source_sha,
            "last_seen_at": "2026-01-04T00:00:00Z",
        }
        document["current_source_binding"] = binding
        document["source_aliases"] = [binding]
        self.catalog_path.write_text(json.dumps(catalog), encoding="utf-8")
        return pptx, normalized

    # TEST A
    def test_concept_query_uses_chunks_without_original_verification(self):
        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="Explain convolution"
        )
        self.assertEqual(result["status"], "FOUND")
        self.assertFalse(result["precision_sensitive"])
        self.assertTrue(result["source_markdown"].endswith("source.md"))
        self.assertEqual(result["chunks"][0]["heading"], "Convolution")
        self.assertEqual(result["verification"], "CACHE_SUFFICIENT_UNLESS_RETRIEVED_TEXT_IS_AMBIGUOUS")

    # TEST B
    def test_exact_formula_query_requires_original_verification(self):
        result = reader.search_catalog(
            catalog_path=self.catalog_path,
            query="Write the exact formula and subscript used for convolution",
        )
        self.assertTrue(result["precision_sensitive"])
        self.assertEqual(result["verification"], "INSPECT_ORIGINAL")

    # TEST C
    def test_existing_material_never_invokes_mineru(self):
        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="Fourier transform"
        )
        self.assertFalse(result["mineru_invoked"])
        source = SCRIPT.read_text(encoding="utf-8")
        self.assertNotIn("subprocess", source)
        self.assertNotIn("mineru.py", source.casefold())

    # TEST D
    def test_missing_original_reports_unavailable_visual_verification(self):
        self.original.unlink()
        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="What is the exact equation?"
        )
        self.assertFalse(result["original_available"])
        self.assertEqual(
            result["verification"],
            "ORIGINAL_UNAVAILABLE_DO_NOT_INVENT_EXACT_NOTATION",
        )

    # TEST E
    def test_multiple_matches_remain_ambiguous_without_context(self):
        self.write_catalog(one_document=False)
        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="Explain this concept"
        )
        self.assertEqual(result["status"], "AMBIGUOUS")
        self.assertEqual(len(result["candidates"]), 2)

    def test_missing_old_binding_falls_back_to_new_identical_source_deterministically(self):
        moved = self.root / "renamed-textbook.pdf"
        moved.write_bytes(self.original.read_bytes())
        self.original.unlink()
        catalog = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        document = catalog["documents"]["doc-one"]
        document["current_source_binding"] = {
            "path": str(self.original),
            "path_key": str(self.original).casefold(),
            "filename": self.original.name,
            "source_sha256": self.source_sha,
            "last_seen_at": "2026-01-03T00:00:00Z",
        }
        document["source_aliases"].append(
            {
                "path": str(moved),
                "path_key": str(moved).casefold(),
                "filename": moved.name,
                "source_sha256": self.source_sha,
                "last_seen_at": "2026-01-02T00:00:00Z",
            }
        )
        self.catalog_path.write_text(json.dumps(catalog), encoding="utf-8")

        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="Write the exact formula"
        )

        self.assertTrue(result["original_available"])
        self.assertEqual(result["original_source"], str(moved.resolve()))
        self.assertEqual(result["verification"], "INSPECT_ORIGINAL")

    def test_revision_selects_profiled_chunks_with_legacy_fallback_unchanged(self):
        retrieval = self.cache / "retrieval" / "profile-one"
        retrieval.mkdir(parents=True)
        profiled_chunks = retrieval / "chunks.json"
        profiled_chunks.write_text(
            json.dumps(
                [
                    {
                        "id": "profiled-0",
                        "index": 0,
                        "heading": "Profiled",
                        "text": "Profiled retrieval artifact.",
                        "chars": 28,
                    }
                ]
            ),
            encoding="utf-8",
        )
        catalog = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        revision = catalog["documents"]["doc-one"]["revisions"][0]
        revision["chunks_path"] = str(profiled_chunks)
        revision["retrieval_profile_fingerprint"] = "profile-one"
        self.catalog_path.write_text(json.dumps(catalog), encoding="utf-8")

        result = reader.search_catalog(
            catalog_path=self.catalog_path, query="Profiled retrieval"
        )

        self.assertEqual(result["chunks_path"], str(profiled_chunks.resolve()))
        self.assertEqual(result["chunks"][0]["heading"], "Profiled")
        self.assertEqual(
            result["document"]["retrieval_profile_fingerprint"], "profile-one"
        )

    def test_retained_normalized_pdf_is_used_for_general_visual_verification(self):
        _, normalized = self.configure_pptx_derivative()

        result = reader.search_catalog(
            catalog_path=self.catalog_path,
            query="Show the exact waveform annotation from the slide",
        )

        self.assertEqual(result["verification"], "INSPECT_NORMALIZED_PDF")
        self.assertEqual(result["visual_source"], str(normalized.resolve()))
        self.assertEqual(result["visual_source_kind"], "normalized_pdf")
        self.assertFalse(result["mineru_invoked"])

    def test_ppt_specific_information_prefers_original_pptx(self):
        pptx, normalized = self.configure_pptx_derivative()

        result = reader.search_catalog(
            catalog_path=self.catalog_path,
            query="What is the animation order and speaker notes?",
        )

        self.assertEqual(result["verification"], "INSPECT_ORIGINAL_PPTX")
        self.assertEqual(result["visual_source"], str(pptx.resolve()))
        self.assertNotEqual(result["visual_source"], str(normalized.resolve()))
        self.assertEqual(result["visual_source_kind"], "original_pptx")
        self.assertFalse(result["mineru_invoked"])

    def test_missing_original_pptx_does_not_treat_pdf_as_ppt_specific_evidence(self):
        pptx, _ = self.configure_pptx_derivative()
        pptx.unlink()

        result = reader.search_catalog(
            catalog_path=self.catalog_path,
            query="Was this a hidden slide with a transition?",
        )

        self.assertEqual(
            result["verification"],
            "ORIGINAL_PPTX_UNAVAILABLE_DO_NOT_INVENT_PPT_SPECIFIC_DETAILS",
        )
        self.assertIsNone(result["visual_source"])

    def test_course_library_publication_path_does_not_change_cache_lookup(self):
        catalog = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        published = (
            self.root
            / "vault"
            / "课程资料"
            / "Signals"
            / "Core Lecture"
        )
        published.mkdir(parents=True)
        (published / "source.md").write_text(
            "human-facing publication copy", encoding="utf-8"
        )
        catalog["documents"]["doc-one"]["publication"] = {
            "source_path": str(published),
            "metadata_path": None,
            "directory_name": "Core Lecture",
        }
        self.catalog_path.write_text(json.dumps(catalog), encoding="utf-8")

        result = reader.search_catalog(
            catalog_path=self.catalog_path,
            query="Explain convolution",
        )

        self.assertEqual(result["status"], "FOUND")
        self.assertEqual(Path(result["source_markdown"]), self.cache / "source.md")
        self.assertNotEqual(Path(result["source_markdown"]), published / "source.md")
        self.assertFalse(result["mineru_invoked"])


if __name__ == "__main__":
    unittest.main()
