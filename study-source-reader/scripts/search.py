#!/usr/bin/env python3
"""Deterministic lexical retrieval for the study-docs catalog.

This helper never imports, invokes, or reimplements MinerU.  It locates an
already-ingested revision and ranks the MinerU-generated chunks.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable, Optional


DEFAULT_STATE_DIR = Path.home() / ".codex" / "study-docs"
_ENGLISH = re.compile(r"[a-z0-9_]{2,}", re.I)
_CJK = re.compile(r"[\u3400-\u9fff]+")
_PRECISION = re.compile(
    r"(?:"
    r"精确|准确|原样|逐字|公式|方程|等式|推导|希腊|符号|变量|上标|下标|角标|正负号|运算符|"
    r"式\s*\(?\d|电路图|波形|频谱|图中|图上|标注|标签|表格|单元格|版式|排版|页码|"
    r"exact|verbatim|formula|equation|derive|derivation|greek|symbol|variable|superscript|subscript|"
    r"operator|sign|circuit\s+diagram|waveform|spectrum|figure|annotation|label|table\s+cell|layout"
    r")",
    re.I,
)
_PPT_SPECIFIC = re.compile(
    r"(?:animation\s+order|build\s+order|hidden\s+slides?|transitions?|speaker\s+notes?|"
    r"editable\s+objects?|object\s+structure|exact\s+(?:slide\s+)?object\s+placement|z[- ]?order|"
    r"动画顺序|播放顺序|隐藏幻灯片|幻灯片切换|切换效果|演讲者备注|讲者备注|备注页|"
    r"可编辑对象|对象结构|对象精确位置|图层顺序)",
    re.I,
)
_STOPWORDS = {
    "what",
    "does",
    "about",
    "from",
    "this",
    "that",
    "explain",
    "find",
    "lecture",
    "textbook",
    "course",
    "using",
    "我的",
    "课本",
    "教材",
    "课件",
    "课程",
    "讲义",
    "什么",
    "怎么",
    "解释",
    "一下",
    "关于",
}


class ReaderError(RuntimeError):
    pass


def normalize(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "").casefold()


def query_terms(query: str) -> list[str]:
    normalized = normalize(query)
    terms: set[str] = set()
    for word in _ENGLISH.findall(normalized):
        if word not in _STOPWORDS:
            terms.add(word)
    for sequence in _CJK.findall(normalized):
        if sequence in _STOPWORDS:
            continue
        if 2 <= len(sequence) <= 8:
            terms.add(sequence)
        for width in (2, 3, 4):
            if len(sequence) >= width:
                for index in range(len(sequence) - width + 1):
                    token = sequence[index : index + width]
                    if token not in _STOPWORDS:
                        terms.add(token)
    return sorted(terms, key=lambda item: (-len(item), item))


def is_precision_sensitive(query: str) -> bool:
    return bool(_PRECISION.search(normalize(query)))


def is_ppt_specific(query: str) -> bool:
    return bool(_PPT_SPECIFIC.search(normalize(query)))


def load_catalog(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as handle:
            catalog = json.load(handle)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReaderError(f"catalog unavailable or invalid: {exc}") from exc
    if not isinstance(catalog, dict) or not isinstance(catalog.get("documents"), dict):
        raise ReaderError("catalog has an unsupported structure")
    return catalog


def _current_revision(document: dict[str, Any]) -> Optional[dict[str, Any]]:
    current = document.get("current_revision_id")
    revisions = document.get("revisions", [])
    if not current or not isinstance(revisions, list):
        return None
    for revision in revisions:
        if isinstance(revision, dict) and revision.get("revision_id") == current:
            return revision
    return None


def _document_search_text(document: dict[str, Any]) -> str:
    values = [str(document.get("course", "")), str(document.get("title", ""))]
    aliases = document.get("source_aliases", [])
    if isinstance(aliases, list):
        for alias in aliases:
            if isinstance(alias, dict):
                values.extend([str(alias.get("filename", "")), str(alias.get("path", ""))])
    return normalize(" ".join(values))


def _document_score(document: dict[str, Any], terms: list[str], query: str) -> int:
    haystack = _document_search_text(document)
    score = 0
    phrase = normalize(query).strip()
    if phrase and phrase in haystack:
        score += 40
    for term in terms:
        if term in haystack:
            score += 3 + min(5, len(term))
    return score


def resolve_document(
    catalog: dict[str, Any],
    *,
    query: str,
    course: Optional[str] = None,
    title: Optional[str] = None,
) -> tuple[Optional[tuple[str, dict[str, Any]]], list[dict[str, Any]]]:
    candidates: list[tuple[str, dict[str, Any]]] = []
    course_folded = normalize(course or "")
    title_folded = normalize(title or "")
    for document_id, document in catalog["documents"].items():
        if not isinstance(document, dict) or _current_revision(document) is None:
            continue
        if course_folded and course_folded not in normalize(str(document.get("course", ""))):
            continue
        if title_folded and title_folded not in normalize(str(document.get("title", ""))):
            continue
        candidates.append((document_id, document))
    if len(candidates) == 1:
        return candidates[0], []
    if not candidates:
        return None, []

    terms = query_terms(query)
    ranked = sorted(
        (
            (_document_score(document, terms, query), document_id, document)
            for document_id, document in candidates
        ),
        key=lambda row: (-row[0], str(row[2].get("course", "")), str(row[2].get("title", "")), row[1]),
    )
    if ranked[0][0] > 0 and (len(ranked) == 1 or ranked[0][0] >= ranked[1][0] + 5):
        return (ranked[0][1], ranked[0][2]), []
    summaries = [
        {
            "document_id": document_id,
            "course": document.get("course"),
            "title": document.get("title"),
            "score": score,
        }
        for score, document_id, document in ranked
    ]
    return None, summaries


def _load_chunks(path: Path) -> list[dict[str, Any]]:
    try:
        with path.open("r", encoding="utf-8") as handle:
            chunks = json.load(handle)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReaderError(f"chunks unavailable or invalid: {exc}") from exc
    if not isinstance(chunks, list):
        raise ReaderError("chunks.json is not an array")
    valid = [chunk for chunk in chunks if isinstance(chunk, dict) and isinstance(chunk.get("text"), str)]
    if not valid:
        raise ReaderError("chunks.json contains no readable chunks")
    return valid


def _chunk_score(chunk: dict[str, Any], query: str, terms: list[str]) -> int:
    text = normalize(str(chunk.get("text", "")))
    heading = normalize(str(chunk.get("heading", "")))
    phrase = normalize(query).strip()
    score = 0
    if phrase and phrase in heading:
        score += 80
    if phrase and phrase in text:
        score += 50
    for term in terms:
        if term in heading:
            score += 12 + min(8, len(term))
        occurrences = text.count(term)
        if occurrences:
            score += min(15, occurrences * 3) + min(5, len(term))
    return score


def _location_metadata(chunk: dict[str, Any]) -> dict[str, Any]:
    keys = ("page", "page_number", "pages", "source_span", "bbox", "images")
    return {key: chunk[key] for key in keys if key in chunk}


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _source_is_current(path: Path, source_sha: Optional[str]) -> bool:
    if not path.is_file():
        return False
    return not source_sha or _sha256_file(path) == source_sha


def _find_original(document: dict[str, Any], revision: dict[str, Any], manifest: dict[str, Any]) -> tuple[Optional[str], bool]:
    source_sha = revision.get("source_sha256")
    preferred_unavailable: Optional[str] = None
    candidates: list[dict[str, Any]] = []
    current = document.get("current_source_binding")
    if isinstance(current, dict) and current.get("source_sha256") == source_sha:
        candidates.append(current)
        if current.get("path"):
            preferred_unavailable = str(current["path"])
    aliases = document.get("source_aliases", [])
    if isinstance(aliases, list):
        matches = [
            alias
            for alias in aliases
            if isinstance(alias, dict) and alias.get("source_sha256") == source_sha
        ]
        matches.sort(key=lambda alias: normalize(str(alias.get("path", ""))))
        matches.sort(key=lambda alias: str(alias.get("last_seen_at", "")), reverse=True)
        current_key = normalize(str(current.get("path", ""))) if isinstance(current, dict) else ""
        candidates.extend(
            alias for alias in matches if normalize(str(alias.get("path", ""))) != current_key
        )
    for candidate in candidates:
        raw = candidate.get("path")
        if not raw:
            continue
        path = Path(raw).expanduser().resolve(strict=False)
        if preferred_unavailable is None:
            preferred_unavailable = str(path)
        if _source_is_current(path, str(source_sha) if source_sha else None):
            return str(path), True
    retained = manifest.get("source_retained_path")
    if retained:
        retained_path = Path(retained).expanduser().resolve(strict=False)
        if _source_is_current(retained_path, str(source_sha) if source_sha else None):
            return str(retained_path), True
    original = manifest.get("source_original_path")
    if original:
        original_path = Path(original).expanduser().resolve(strict=False)
        if _source_is_current(original_path, str(source_sha) if source_sha else None):
            return str(original_path), True
        if preferred_unavailable is None:
            preferred_unavailable = str(original_path)
    return preferred_unavailable, False


def _find_normalized_pdf(cache_dir: Path, manifest: dict[str, Any]) -> tuple[Optional[str], bool]:
    normalized = manifest.get("normalized_source")
    if not isinstance(normalized, dict) or not normalized.get("path"):
        return None, False
    candidate = (cache_dir / str(normalized["path"])).resolve(strict=False)
    try:
        candidate.relative_to(cache_dir)
    except ValueError:
        return str(candidate), False
    if candidate.suffix.casefold() != ".pdf" or not candidate.is_file():
        return str(candidate), False
    expected_hash = normalized.get("sha256") or manifest.get("normalized_pdf_sha256")
    if expected_hash and _sha256_file(candidate) != expected_hash:
        return str(candidate), False
    return str(candidate), True


def _revision_chunks_path(cache_dir: Path, revision: dict[str, Any]) -> Path:
    raw = revision.get("chunks_path")
    if not raw:
        return cache_dir / "chunks.json"
    candidate = Path(str(raw)).expanduser()
    if not candidate.is_absolute():
        candidate = cache_dir / candidate
    candidate = candidate.resolve(strict=False)
    try:
        candidate.relative_to(cache_dir)
    except ValueError as exc:
        raise ReaderError("catalog chunks path escapes the extraction cache") from exc
    return candidate


def search_catalog(
    *,
    catalog_path: Path,
    query: str,
    course: Optional[str] = None,
    title: Optional[str] = None,
    limit: int = 5,
) -> dict[str, Any]:
    catalog = load_catalog(catalog_path)
    selected, ambiguous = resolve_document(
        catalog, query=query, course=course, title=title
    )
    ppt_specific_query = is_ppt_specific(query)
    precision = is_precision_sensitive(query) or ppt_specific_query
    if selected is None:
        if ambiguous:
            return {
                "status": "AMBIGUOUS",
                "query": query,
                "precision_sensitive": precision,
                "candidates": ambiguous,
                "message": "Specify course or title; no document was chosen silently.",
            }
        return {
            "status": "NOT_FOUND",
            "query": query,
            "precision_sensitive": precision,
            "message": "No matching ingested document is available.",
        }
    document_id, document = selected
    revision = _current_revision(document)
    assert revision is not None
    cache_dir = Path(str(revision.get("cache_path", ""))).expanduser().resolve(strict=False)
    source_md = cache_dir / "source.md"
    chunks_path = _revision_chunks_path(cache_dir, revision)
    manifest_path = cache_dir / "manifest.json"
    if not source_md.is_file() or not chunks_path.is_file() or not manifest_path.is_file():
        raise ReaderError("current catalog revision is missing required cache artifacts")
    try:
        with manifest_path.open("r", encoding="utf-8") as handle:
            manifest = json.load(handle)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReaderError(f"manifest unavailable or invalid: {exc}") from exc
    chunks = _load_chunks(chunks_path)
    terms = query_terms(query)
    ranked = sorted(
        (
            (_chunk_score(chunk, query, terms), int(chunk.get("index", position)), chunk)
            for position, chunk in enumerate(chunks)
        ),
        key=lambda row: (-row[0], row[1]),
    )
    top = []
    for score, index, chunk in ranked[: max(1, limit)]:
        top.append(
            {
                "score": score,
                "id": chunk.get("id"),
                "index": index,
                "heading": chunk.get("heading", ""),
                "text": chunk.get("text", ""),
                "chars": chunk.get("chars"),
                "location": _location_metadata(chunk),
            }
        )
    original_path, original_available = _find_original(document, revision, manifest)
    normalized_path, normalized_available = _find_normalized_pdf(cache_dir, manifest)
    original_metadata = manifest.get("original_source")
    original_filename = (
        original_metadata.get("filename")
        if isinstance(original_metadata, dict)
        else manifest.get("source_filename")
    )
    pptx_revision = str(original_filename or original_path or "").casefold().endswith(".pptx")
    ppt_specific = pptx_revision and ppt_specific_query
    if precision:
        if ppt_specific:
            verification = (
                "INSPECT_ORIGINAL_PPTX"
                if original_available
                else "ORIGINAL_PPTX_UNAVAILABLE_DO_NOT_INVENT_PPT_SPECIFIC_DETAILS"
            )
            visual_source = original_path if original_available else None
            visual_source_kind = "original_pptx" if original_available else None
        elif normalized_available:
            verification = "INSPECT_NORMALIZED_PDF"
            visual_source = normalized_path
            visual_source_kind = "normalized_pdf"
        else:
            verification = (
                "INSPECT_ORIGINAL"
                if original_available
                else "ORIGINAL_UNAVAILABLE_DO_NOT_INVENT_EXACT_NOTATION"
            )
            visual_source = original_path if original_available else None
            visual_source_kind = "original" if original_available else None
    else:
        verification = "CACHE_SUFFICIENT_UNLESS_RETRIEVED_TEXT_IS_AMBIGUOUS"
        visual_source = None
        visual_source_kind = None
    return {
        "status": "FOUND",
        "query": query,
        "precision_sensitive": precision,
        "verification": verification,
        "document": {
            "document_id": document_id,
            "course": document.get("course"),
            "title": document.get("title"),
            "revision_id": revision.get("revision_id"),
            "retrieval_profile_fingerprint": revision.get("retrieval_profile_fingerprint"),
        },
        "cache_path": str(cache_dir),
        "source_markdown": str(source_md),
        "chunks_path": str(chunks_path),
        "assets_path": str(cache_dir / "assets"),
        "original_source": original_path,
        "original_available": original_available,
        "normalized_source": normalized_path,
        "normalized_available": normalized_available,
        "ppt_specific": ppt_specific,
        "visual_source": visual_source,
        "visual_source_kind": visual_source_kind,
        "chunks": top,
        "mineru_invoked": False,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="study-source-reader",
        description="Search previously ingested course sources without rerunning MinerU.",
    )
    parser.add_argument("query", help="Concept or exact item to retrieve")
    parser.add_argument("--course")
    parser.add_argument("--title")
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument(
        "--catalog",
        default=str(Path(os.environ.get("STUDY_DOCS_STATE_DIR", DEFAULT_STATE_DIR)) / "catalog.json"),
    )
    parser.add_argument("--json", action="store_true")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = search_catalog(
            catalog_path=Path(args.catalog),
            query=args.query,
            course=args.course,
            title=args.title,
            limit=args.limit,
        )
    except Exception as exc:
        result = {"status": "FAILED", "error": str(exc), "mineru_invoked": False}
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif result["status"] == "FOUND":
        document = result["document"]
        print(f"FOUND | {document['course']} | {document['title']} | {result['verification']}")
        for chunk in result["chunks"]:
            print(f"[{chunk['score']}] {chunk['heading']} (chunk {chunk['index']})")
            print(chunk["text"])
    elif result["status"] == "AMBIGUOUS":
        print("AMBIGUOUS | specify --course or --title")
        for candidate in result["candidates"]:
            print(f"  {candidate['document_id']} | {candidate['course']} | {candidate['title']}")
    else:
        print(f"{result['status']} | {result.get('message') or result.get('error')}")
    return 0 if result["status"] == "FOUND" else 2 if result["status"] == "AMBIGUOUS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
