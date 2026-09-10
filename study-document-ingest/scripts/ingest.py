#!/usr/bin/env python3
"""Persistent, content-addressed study-document ingestion over MinerU.

MinerU remains the only parser and chunker.  This module adds deterministic
identity, validation, cache/catalog management, and rollback-safe publication.
It deliberately uses only the Python standard library (plus optional pypdf for
PDF page-count preflight, already used by MinerU's native --split workflow).
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import hashlib
import json
import os
import re
import runpy
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import urllib.parse
import uuid
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Callable, Iterable, Optional


PIPELINE_SCHEMA_VERSION = "2"
LEGACY_PIPELINE_SCHEMA_VERSION = "1"
EXTRACTION_PROFILE_SCHEMA_VERSION = "1"
RETRIEVAL_PROFILE_SCHEMA_VERSION = "1"
PPTX_NORMALIZATION_SCHEMA_VERSION = "1"
CATALOG_SCHEMA_VERSION = 1
CONFIG_SCHEMA_VERSION = 1
DEFAULT_CHUNK_SIZE = 2000
DEFAULT_FALLBACK_COURSE = "_Inbox"
DEFAULT_PUBLICATION_ROOT = "课程资料"
RESERVED_NOTES_ROOT = "笔记"

DEFAULT_STATE_DIR = Path.home() / ".codex" / "study-docs"


def default_mineru_script() -> Path:
    """Locate the separately installed MinerU skill without a user-specific path."""
    sibling = Path(__file__).resolve().parents[2] / "mineru" / "scripts" / "mineru.py"
    candidates = [
        sibling,
        Path.home() / ".agents" / "skills" / "mineru" / "scripts" / "mineru.py",
    ]
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        candidates.append(Path(codex_home).expanduser() / "skills" / "mineru" / "scripts" / "mineru.py")
    candidates.append(Path.home() / ".codex" / "skills" / "mineru" / "scripts" / "mineru.py")

    seen: set[str] = set()
    unique: list[Path] = []
    for candidate in candidates:
        key = os.path.normcase(str(candidate.resolve(strict=False))).casefold()
        if key not in seen:
            seen.add(key)
            unique.append(candidate)
    return next((candidate for candidate in unique if candidate.is_file()), unique[0])

SUPPORTED_SUFFIXES = {
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".jp2",
    ".webp",
    ".gif",
    ".bmp",
    ".doc",
    ".docx",
    ".ppt",
    ".pptx",
    ".xls",
    ".xlsx",
    ".html",
}
IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".jp2",
    ".webp",
    ".gif",
    ".bmp",
    ".svg",
    ".tif",
    ".tiff",
}

GENERIC_PARENT_NAMES = {
    "downloads",
    "download",
    "desktop",
    "documents",
    "document",
    "temp",
    "tmp",
    "work",
    "output",
    "outputs",
    "cache",
    "staging",
    "attachments",
    "sources",
    "notes",
    "slides",
    "lectures",
    "lecture",
    "materials",
    "handouts",
    "pdf",
    "pdfs",
    "codex",
    "files-pasted-by-the-user-you",
    "下载",
    "桌面",
    "文档",
    "临时",
    "资料",
    "课件",
}
TERMINAL_GENERIC_PARENTS = {
    "downloads",
    "download",
    "desktop",
    "documents",
    "document",
    "temp",
    "tmp",
    "work",
    "output",
    "outputs",
    "attachments",
    "files-pasted-by-the-user-you",
    "下载",
    "桌面",
    "文档",
    "临时",
}
SKIP_DIRECTORY_NAMES = {
    ".study-docs",
    "study-docs",
    "cache",
    "staging",
    "output",
    "outputs",
    "__pycache__",
    ".git",
    ".obsidian",
}
WINDOWS_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

_INVALID_COMPONENT = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")
_IMAGE_REF = re.compile(
    r"!\[(?P<alt>[^\]]*)\]\((?P<target><[^>]+>|[^)\s]+)(?P<title>\s+(?:\"[^\"]*\"|'[^']*'))?\)"
)
_ENCODING_FAILURE = re.compile(
    r"UnicodeEncodeError|charmap|cp936|gbk|codec can.t encode", re.IGNORECASE
)
_PAGE_LIMIT_FAILURE = re.compile(
    r"-60006|too many pages|pages exceed|page limit", re.IGNORECASE
)


class PipelineError(RuntimeError):
    """Base error for deterministic pipeline failures."""


class ValidationError(PipelineError):
    """Raised when new or cached artifacts fail validation."""

    def __init__(self, message: str, *, code: str = "validation_failed"):
        super().__init__(message)
        self.code = code


class PublicationError(PipelineError):
    """Raised when publication cannot be completed and has been rolled back."""


class PrivateModeError(PipelineError):
    """Raised before execution when private mode cannot remain fully local."""


class NormalizationError(PipelineError):
    """Raised before MinerU when a quality-mode PPTX cannot be normalized safely."""


@dataclasses.dataclass(frozen=True)
class ExtractionProfile:
    pipeline_schema_version: str
    mode: str
    engine: str
    api: str
    model: str
    ocr_policy: str
    language_policy: str
    chunk_enabled: bool
    chunk_size: int
    split_policy: str
    input_modality: str = "native"
    normalization: str = "none"
    normalizer_backend: Optional[str] = None
    normalizer_version: Optional[str] = None
    normalization_schema_version: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class ProcessOutcome:
    command: list[str]
    returncode: int
    stdout: str
    stderr: str
    status: Optional[dict[str, Any]]
    duration: float


@dataclasses.dataclass
class ArtifactValidation:
    markdown_chars: int
    chunk_count: int
    image_references: int
    resolved_images: int
    missing_images: list[str]
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass(frozen=True)
class NormalizerBackend:
    name: str
    version: Optional[str]
    executable: Optional[str] = None


@dataclasses.dataclass
class NormalizationResult:
    path: Path
    backend: str
    version: Optional[str]
    slide_count: Optional[int]
    page_count: int
    sha256: str
    duration: float
    warnings: list[str]


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _profile_mapping(profile: ExtractionProfile | dict[str, Any]) -> dict[str, Any]:
    return profile.to_dict() if isinstance(profile, ExtractionProfile) else dict(profile)


def extraction_profile(profile: ExtractionProfile | dict[str, Any]) -> dict[str, Any]:
    """Return only settings that can change MinerU parsing output."""
    payload = _profile_mapping(profile)
    result = {
        "schema_version": EXTRACTION_PROFILE_SCHEMA_VERSION,
        "engine": payload.get("engine"),
        "api": payload.get("api"),
        "model": payload.get("model"),
        "ocr_policy": payload.get("ocr_policy"),
        "language_policy": payload.get("language_policy"),
        "split_policy": payload.get("split_policy"),
    }
    if payload.get("normalization", "none") != "none":
        result.update(
            {
                "input_modality": payload.get("input_modality"),
                "normalization": payload.get("normalization"),
                "normalizer_backend": payload.get("normalizer_backend"),
                "normalizer_version": payload.get("normalizer_version"),
                "normalization_schema_version": payload.get(
                    "normalization_schema_version"
                ),
            }
        )
    return result


def extraction_profile_fingerprint(profile: ExtractionProfile | dict[str, Any]) -> str:
    return sha256_bytes(canonical_json(extraction_profile(profile)).encode("utf-8"))


def retrieval_profile(
    profile: ExtractionProfile | dict[str, Any], *, chunker_fingerprint: str
) -> dict[str, Any]:
    """Return settings for deterministic derivation from cached source.md."""
    payload = _profile_mapping(profile)
    return {
        "schema_version": RETRIEVAL_PROFILE_SCHEMA_VERSION,
        "chunk_enabled": bool(payload.get("chunk_enabled", True)),
        "chunk_size": int(payload.get("chunk_size", DEFAULT_CHUNK_SIZE)),
        "chunker": "mineru-heading-aware-markdown",
        "chunker_fingerprint": chunker_fingerprint,
    }


def retrieval_profile_fingerprint(profile: dict[str, Any]) -> str:
    return sha256_bytes(canonical_json(profile).encode("utf-8"))


def profile_fingerprint(profile: ExtractionProfile | dict[str, Any]) -> str:
    """Backward-compatible alias for the extraction-only fingerprint."""
    return extraction_profile_fingerprint(profile)


def extraction_cache_key(source_sha256: str, fingerprint: str) -> str:
    structured = {
        "extraction_profile_fingerprint": fingerprint,
        "source_sha256": source_sha256,
    }
    return sha256_bytes(canonical_json(structured).encode("utf-8"))


def legacy_profile_fingerprint(profile: dict[str, Any]) -> str:
    return sha256_bytes(canonical_json(profile).encode("utf-8"))


def legacy_extraction_cache_key(source_sha256: str, fingerprint: str) -> str:
    structured = {
        "profile_fingerprint": fingerprint,
        "source_sha256": source_sha256,
    }
    return sha256_bytes(canonical_json(structured).encode("utf-8"))


def profile_for_mode(
    mode: str,
    *,
    language: str = "ch",
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> ExtractionProfile:
    if chunk_size < 200:
        raise PipelineError("chunk size must be at least 200 characters")
    common = {
        "pipeline_schema_version": PIPELINE_SCHEMA_VERSION,
        "mode": mode,
        "language_policy": language,
        "chunk_enabled": True,
        "chunk_size": chunk_size,
    }
    if mode == "quality":
        return ExtractionProfile(
            **common,
            engine="cloud",
            api="auto",
            model="vlm",
            ocr_policy="normal-first-controlled-sparse-fallback",
            split_policy="mineru-native-on-active-backend-limit",
        )
    if mode == "fast":
        return ExtractionProfile(
            **common,
            engine="auto",
            api="auto",
            model="pipeline",
            ocr_policy="disabled",
            split_policy="mineru-native-on-active-backend-limit",
        )
    if mode == "private":
        return ExtractionProfile(
            **common,
            engine="local",
            api="disabled",
            model="local-text",
            ocr_policy="disabled",
            split_policy="disabled",
        )
    raise PipelineError(f"unsupported mode: {mode}")


def default_config() -> dict[str, Any]:
    return {
        "schema_version": CONFIG_SCHEMA_VERSION,
        "vault_path": None,
        "default_mode": "quality",
        "default_course_fallback": DEFAULT_FALLBACK_COURSE,
        "source_retention": "reference",
        "publish_enabled": True,
        "publication_root": DEFAULT_PUBLICATION_ROOT,
    }


def default_catalog() -> dict[str, Any]:
    return {
        "schema_version": CATALOG_SCHEMA_VERSION,
        "updated_at": utc_now(),
        "documents": {},
        "cache_entries": {},
    }


def safe_component(value: str, *, fallback: str = "_Untitled", max_length: int = 96) -> str:
    original = unicodedata.normalize("NFKC", str(value or "")).strip()
    cleaned = _INVALID_COMPONENT.sub("_", original)
    cleaned = cleaned.replace("..", "_")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if not cleaned or cleaned in {".", ".."}:
        cleaned = fallback
    if cleaned.upper() in WINDOWS_RESERVED:
        cleaned = f"_{cleaned}"
    if len(cleaned) > max_length:
        suffix = sha256_bytes(original.encode("utf-8"))[:10]
        cleaned = f"{cleaned[: max_length - 12].rstrip(' .') }--{suffix}"
    return cleaned


def normalized_path_key(path: Path | str) -> str:
    resolved = Path(path).expanduser().resolve(strict=False)
    return os.path.normcase(str(resolved)).casefold()


def is_within(path: Path, root: Path) -> bool:
    try:
        return os.path.commonpath(
            [normalized_path_key(path), normalized_path_key(root)]
        ) == normalized_path_key(root)
    except ValueError:
        return False


def ensure_within(path: Path, root: Path, label: str) -> Path:
    resolved = path.resolve(strict=False)
    if not is_within(resolved, root.resolve(strict=False)):
        raise ValidationError(f"{label} escapes controlled directory: {path}", code="path_escape")
    return resolved


def infer_course(source: Path, fallback: str = DEFAULT_FALLBACK_COURSE) -> str:
    date_like = re.compile(r"^\d{4}(?:[-_.]\d{1,2}){1,2}$")
    uuid_like = re.compile(r"^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$", re.I)
    parents = list(source.resolve(strict=False).parents)
    if parents and parents[0].name.strip().casefold() in TERMINAL_GENERIC_PARENTS:
        return safe_component(fallback, fallback=DEFAULT_FALLBACK_COURSE)
    for parent in parents[:3]:
        name = parent.name.strip()
        folded = name.casefold()
        if not name:
            continue
        if folded.startswith(".") or folded in GENERIC_PARENT_NAMES:
            continue
        if date_like.fullmatch(name) or uuid_like.fullmatch(name) or (
            folded.startswith("tmp") and len(name) > 3
        ):
            continue
        # Generated Codex task directories are poor course evidence.
        if folded.startswith("files-") or len(name) > 80:
            continue
        return safe_component(name, fallback=fallback)
    return safe_component(fallback, fallback=DEFAULT_FALLBACK_COURSE)


def source_title(source: Path) -> str:
    return safe_component(source.stem, fallback="_Untitled")


def atomic_write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.parent / f".{path.name}.{uuid.uuid4().hex}.tmp"
    try:
        with tmp.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, sort_keys=True, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.parent / f".{path.name}.{uuid.uuid4().hex}.tmp"
    try:
        with tmp.open("w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def read_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def validate_config(config: dict[str, Any]) -> None:
    if not isinstance(config, dict) or config.get("schema_version") != CONFIG_SCHEMA_VERSION:
        raise ValidationError("unsupported or invalid config schema", code="invalid_config")
    if config.get("default_mode") not in {"quality", "fast", "private"}:
        raise ValidationError("invalid default_mode in config", code="invalid_config")
    if config.get("source_retention") not in {"reference", "copy"}:
        raise ValidationError("source_retention must be reference or copy", code="invalid_config")
    vault = config.get("vault_path")
    if vault is not None and not isinstance(vault, str):
        raise ValidationError("vault_path must be a string or null", code="invalid_config")
    configured_publication_root(config)


def configured_publication_root(config: dict[str, Any]) -> str:
    value = config.get("publication_root", DEFAULT_PUBLICATION_ROOT)
    if not isinstance(value, str) or not value.strip():
        raise ValidationError("publication_root must be a directory name", code="invalid_config")
    value = value.strip()
    if (
        Path(value).is_absolute()
        or "/" in value
        or "\\" in value
        or value in {".", ".."}
        or safe_component(value) != value
    ):
        raise ValidationError(
            "publication_root must be one safe top-level directory name",
            code="invalid_config",
        )
    if value.casefold() == RESERVED_NOTES_ROOT.casefold():
        raise ValidationError(
            f"publication_root cannot use the reserved personal-notes directory {RESERVED_NOTES_ROOT}",
            code="invalid_config",
        )
    if value.casefold() != DEFAULT_PUBLICATION_ROOT.casefold():
        raise ValidationError(
            f"publication_root must be {DEFAULT_PUBLICATION_ROOT}",
            code="invalid_config",
        )
    return value


def validate_catalog_shape(catalog: dict[str, Any]) -> None:
    if not isinstance(catalog, dict) or catalog.get("schema_version") != CATALOG_SCHEMA_VERSION:
        raise ValidationError("unsupported or invalid catalog schema", code="invalid_catalog")
    if not isinstance(catalog.get("documents"), dict):
        raise ValidationError("catalog documents must be an object", code="invalid_catalog")
    if not isinstance(catalog.get("cache_entries"), dict):
        raise ValidationError("catalog cache_entries must be an object", code="invalid_catalog")


def initialize_state(state_dir: Path) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    for child in ("cache", "staging", "logs"):
        (state_dir / child).mkdir(parents=True, exist_ok=True)
    config_path = state_dir / "config.json"
    catalog_path = state_dir / "catalog.json"
    if not config_path.exists():
        atomic_write_json(config_path, default_config())
    if not catalog_path.exists():
        atomic_write_json(catalog_path, default_catalog())


def load_config(state_dir: Path) -> dict[str, Any]:
    path = state_dir / "config.json"
    config = default_config()
    if path.exists():
        loaded = read_json(path)
        if not isinstance(loaded, dict):
            raise ValidationError("config must be an object", code="invalid_config")
        config.update(loaded)
    validate_config(config)
    return config


def load_catalog(state_dir: Path) -> dict[str, Any]:
    path = state_dir / "catalog.json"
    catalog = read_json(path) if path.exists() else default_catalog()
    validate_catalog_shape(catalog)
    return catalog


def _redact_text(text: str) -> str:
    token = os.environ.get("MINERU_TOKEN")
    if token:
        text = text.replace(token, "[REDACTED]")
    return text


def _redact_value(value: Any, key: str = "") -> Any:
    if any(word in key.casefold() for word in ("token", "secret", "password")):
        return "[REDACTED]"
    if isinstance(value, str):
        return _redact_text(value)
    if isinstance(value, list):
        return [_redact_value(item) for item in value]
    if isinstance(value, dict):
        return {str(k): _redact_value(v, str(k)) for k, v in value.items()}
    return value


class PipelineLogger:
    def __init__(self, log_dir: Path):
        self.log_dir = log_dir

    def log(self, event: str, **fields: Any) -> None:
        self.log_dir.mkdir(parents=True, exist_ok=True)
        path = self.log_dir / f"ingest-{dt.date.today().isoformat()}.jsonl"
        if path.exists() and path.stat().st_size > 2 * 1024 * 1024:
            rotated = path.with_suffix(path.suffix + ".1")
            try:
                rotated.unlink(missing_ok=True)
                path.replace(rotated)
            except OSError:
                pass
        record = _redact_value({"timestamp": utc_now(), "event": event, **fields})
        with path.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")


def parse_json_status(stdout: str) -> Optional[dict[str, Any]]:
    text = stdout.lstrip("\ufeff\r\n\t ")
    try:
        value = json.loads(text)
        return value if isinstance(value, dict) else None
    except json.JSONDecodeError:
        pass
    decoder = json.JSONDecoder()
    for match in re.finditer(r"\{", text):
        try:
            value, _ = decoder.raw_decode(text[match.start() :])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict) and ("results" in value or "healthy" in value):
            return value
    return None


def subprocess_utf8_environment() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def run_captured(command: list[str], *, timeout: int = 1200) -> ProcessOutcome:
    started = time.monotonic()
    try:
        completed = subprocess.run(
            command,
            shell=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=subprocess_utf8_environment(),
            timeout=timeout,
            check=False,
        )
        stdout = completed.stdout.decode("utf-8", errors="replace")
        stderr = completed.stderr.decode("utf-8", errors="replace")
        return ProcessOutcome(
            command=command,
            returncode=completed.returncode,
            stdout=stdout,
            stderr=stderr,
            status=parse_json_status(stdout),
            duration=time.monotonic() - started,
        )
    except subprocess.TimeoutExpired as exc:
        stdout_bytes = exc.stdout or b""
        stderr_bytes = exc.stderr or b""
        if isinstance(stdout_bytes, str):
            stdout = stdout_bytes
        else:
            stdout = stdout_bytes.decode("utf-8", errors="replace")
        if isinstance(stderr_bytes, str):
            stderr = stderr_bytes
        else:
            stderr = stderr_bytes.decode("utf-8", errors="replace")
        return ProcessOutcome(
            command=command,
            returncode=124,
            stdout=stdout,
            stderr=stderr + "\nsubprocess timeout",
            status=parse_json_status(stdout),
            duration=time.monotonic() - started,
        )
    except OSError as exc:
        return ProcessOutcome(
            command=command,
            returncode=127,
            stdout="",
            stderr=f"{type(exc).__name__}: {exc}",
            status=None,
            duration=time.monotonic() - started,
        )


def mineru_fingerprint(mineru_script: Path) -> Optional[str]:
    root = mineru_script.parent.parent
    relative_files = (
        Path("SKILL.md"),
        Path("scripts/mineru.py"),
        Path("scripts/chunking.py"),
        Path("scripts/splitter.py"),
        Path("scripts/local_engine.py"),
    )
    digest = hashlib.sha256()
    found = 0
    for relative in relative_files:
        path = root / relative
        if not path.is_file():
            continue
        found += 1
        digest.update(relative.as_posix().encode("utf-8"))
        digest.update(b"\0")
        digest.update(sha256_file(path).encode("ascii"))
        digest.update(b"\0")
    return digest.hexdigest() if found else None


def mineru_version_from_source(mineru_script: Path) -> Optional[str]:
    try:
        text = mineru_script.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    match = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', text, re.M)
    return match.group(1) if match else None


def interpreter_has_module(python_executable: Path, module: str) -> bool:
    command = [
        str(python_executable),
        "-c",
        "import importlib.util,sys;sys.exit(0 if importlib.util.find_spec(sys.argv[1]) else 1)",
        module,
    ]
    outcome = run_captured(command, timeout=30)
    return outcome.returncode == 0


def load_mineru_limits(mineru_script: Path) -> dict[str, int]:
    fallback = {"agent": 20, "standard": 200}
    if not mineru_script.is_file():
        return fallback
    try:
        text = mineru_script.read_text(encoding="utf-8", errors="replace")
        agent = re.search(r"^AGENT_MAX_PAGES\s*=\s*(\d+)", text, re.M)
        standard = re.search(r"^STANDARD_MAX_PAGES\s*=\s*(\d+)", text, re.M)
        if not agent or not standard:
            return fallback
        return {
            "agent": int(agent.group(1)),
            "standard": int(standard.group(1)),
        }
    except Exception:
        return fallback


def pdf_page_count(path: Path) -> Optional[int]:
    try:
        from pypdf import PdfReader  # type: ignore

        return len(PdfReader(str(path)).pages)
    except Exception:
        return None


def readable_pdf_page_count(path: Path) -> int:
    """Open a normalized PDF structurally and require at least one readable page."""
    try:
        from pypdf import PdfReader  # type: ignore

        count = len(PdfReader(str(path)).pages)
    except Exception as exc:
        raise NormalizationError(f"normalized PDF is structurally unreadable: {exc}") from exc
    if count < 1:
        raise NormalizationError("normalized PDF contains no pages")
    return count


def pptx_slide_count(path: Path) -> Optional[int]:
    """Read the declared slide list directly from the PPTX package, if available."""
    try:
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read("ppt/presentation.xml"))
        namespace = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        slide_list = root.find(f"{namespace}sldIdLst")
        return len(list(slide_list)) if slide_list is not None else 0
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile):
        return None


def _json_stdout(stdout: str) -> Optional[dict[str, Any]]:
    try:
        value = json.loads(stdout.strip().lstrip("\ufeff"))
    except json.JSONDecodeError:
        return None
    return value if isinstance(value, dict) else None


def detect_normalizer_backends(powerpoint_script: Path) -> list[NormalizerBackend]:
    """Detect local converters without starting PowerPoint or changing Office state."""
    backends: list[NormalizerBackend] = []
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if os.name == "nt" and powershell and powerpoint_script.is_file():
        outcome = run_captured(
            [
                powershell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(powerpoint_script),
                "-Probe",
            ],
            timeout=30,
        )
        report = _json_stdout(outcome.stdout)
        if outcome.returncode == 0 and report and report.get("available"):
            backends.append(
                NormalizerBackend(
                    "powerpoint",
                    str(report.get("version")) if report.get("version") else None,
                    str(report.get("executable")) if report.get("executable") else None,
                )
            )

    candidates: list[Path] = []
    discovered = shutil.which("soffice.com") or shutil.which("soffice.exe") or shutil.which("soffice")
    if discovered:
        candidates.append(Path(discovered))
    for variable in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(variable)
        if root:
            candidates.extend(
                [
                    Path(root) / "LibreOffice" / "program" / "soffice.com",
                    Path(root) / "LibreOffice" / "program" / "soffice.exe",
                ]
            )
    seen: set[str] = set()
    for candidate in candidates:
        resolved = candidate.expanduser().resolve(strict=False)
        key = normalized_path_key(resolved)
        if key in seen or not resolved.is_file():
            continue
        seen.add(key)
        outcome = run_captured([str(resolved), "--version"], timeout=30)
        if outcome.returncode != 0:
            continue
        line = next((item.strip() for item in outcome.stdout.splitlines() if item.strip()), "")
        match = re.search(r"LibreOffice\s+([^\s]+)", line, re.I)
        backends.append(
            NormalizerBackend(
                "libreoffice",
                match.group(1) if match else (line or None),
                str(resolved),
            )
        )
        break
    return backends


def convert_pptx_to_pdf(
    source: Path,
    destination: Path,
    backend: NormalizerBackend,
    *,
    powerpoint_script: Path,
) -> dict[str, Any]:
    """Convert one PPTX with a selected local backend into a controlled path."""
    destination.parent.mkdir(parents=True, exist_ok=False)
    if backend.name == "powerpoint":
        powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        if not powershell or not powerpoint_script.is_file():
            raise NormalizationError("PowerPoint export helper is unavailable")
        outcome = run_captured(
            [
                powershell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(powerpoint_script),
                "-Source",
                str(source),
                "-Destination",
                str(destination),
            ],
            timeout=300,
        )
        report = _json_stdout(outcome.stdout)
        if outcome.returncode != 0 or not report or not report.get("ok"):
            detail = _redact_text(outcome.stderr or outcome.stdout).strip()[-1200:]
            raise NormalizationError(f"PowerPoint PDF export failed: {detail or 'unknown error'}")
        return report
    if backend.name == "libreoffice":
        executable = Path(str(backend.executable or "")).resolve(strict=False)
        if not executable.is_file():
            raise NormalizationError("LibreOffice executable is unavailable")
        profile_dir = destination.parent / "libreoffice-profile"
        profile_dir.mkdir()
        outcome = run_captured(
            [
                str(executable),
                "--headless",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(destination.parent),
                str(source),
            ],
            timeout=300,
        )
        if outcome.returncode != 0:
            detail = _redact_text(outcome.stderr or outcome.stdout).strip()[-1200:]
            raise NormalizationError(f"LibreOffice PDF export failed: {detail or 'unknown error'}")
        produced = destination.parent / f"{source.stem}.pdf"
        if produced.resolve(strict=False) != destination.resolve(strict=False) and produced.is_file():
            produced.replace(destination)
        return {"ok": True, "backend": "libreoffice"}
    raise NormalizationError(f"unsupported PPTX normalizer backend: {backend.name}")


def _status_markdown_candidates(status: Optional[dict[str, Any]]) -> Iterable[str]:
    if not isinstance(status, dict):
        return []
    candidates: list[str] = []
    results = status.get("results")
    if isinstance(results, list):
        for item in results:
            if isinstance(item, dict) and isinstance(item.get("markdown_path"), str):
                candidates.append(item["markdown_path"])
    return candidates


def discover_markdown(output_root: Path, status: Optional[dict[str, Any]]) -> Path:
    root = output_root.resolve(strict=False)
    for raw in _status_markdown_candidates(status):
        candidate = Path(raw)
        if not candidate.is_absolute():
            candidate = output_root / candidate
        try:
            candidate = ensure_within(candidate, root, "MinerU markdown path")
        except ValidationError:
            continue
        if candidate.is_file() and candidate.suffix.casefold() == ".md":
            return candidate
    candidates = [
        path
        for path in output_root.rglob("*.md")
        if path.is_file() and is_within(path, root)
    ]
    if not candidates:
        raise ValidationError("MinerU produced no Markdown artifact", code="missing_markdown")
    candidates.sort(key=lambda path: (path.stat().st_size, str(path)), reverse=True)
    return candidates[0]


def _decode_image_target(raw: str) -> tuple[str, bool]:
    bracketed = raw.startswith("<") and raw.endswith(">")
    target = raw[1:-1] if bracketed else raw
    return urllib.parse.unquote(target), bracketed


def _is_remote_target(target: str) -> bool:
    lowered = target.casefold()
    return lowered.startswith(("http://", "https://", "data:", "#"))


def normalize_mineru_artifacts(
    markdown_path: Path,
    normalized_root: Path,
) -> list[str]:
    """Copy MinerU's Markdown/images into the immutable extraction cache."""
    try:
        markdown = markdown_path.read_text(encoding="utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        raise ValidationError(f"Markdown is not valid UTF-8: {exc}", code="invalid_utf8") from exc

    normalized_root.mkdir(parents=True, exist_ok=False)
    assets_root = normalized_root / "assets"
    assets_root.mkdir()
    markdown_dir = markdown_path.parent.resolve(strict=False)
    warnings: list[str] = []

    for candidate in markdown_dir.rglob("*"):
        if not candidate.is_file() or candidate.suffix.casefold() not in IMAGE_SUFFIXES:
            continue
        resolved = candidate.resolve(strict=False)
        if not is_within(resolved, markdown_dir):
            warnings.append(f"skipped asset outside MinerU output: {candidate}")
            continue
        relative = resolved.relative_to(markdown_dir)
        target = assets_root / relative
        ensure_within(target, assets_root, "asset destination")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(resolved, target)

    def rewrite(match: re.Match[str]) -> str:
        raw_target = match.group("target")
        decoded, bracketed = _decode_image_target(raw_target)
        if _is_remote_target(decoded):
            return match.group(0)
        candidate = (markdown_dir / decoded).resolve(strict=False)
        if not is_within(candidate, markdown_dir):
            warnings.append(f"image reference escapes MinerU output: {decoded}")
            return match.group(0)
        try:
            relative = candidate.relative_to(markdown_dir)
        except ValueError:
            return match.group(0)
        new_target = Path("assets") / relative
        encoded = urllib.parse.quote(new_target.as_posix(), safe="/._-")
        if bracketed:
            encoded = f"<{encoded}>"
        title = match.group("title") or ""
        return f"![{match.group('alt')}]({encoded}{title})"

    normalized_markdown = _IMAGE_REF.sub(rewrite, markdown)
    atomic_write_text(normalized_root / "source.md", normalized_markdown)
    return warnings


def validate_extraction_artifacts(root: Path) -> ArtifactValidation:
    root = root.resolve(strict=False)
    markdown_path = ensure_within(root / "source.md", root, "source.md")
    if not markdown_path.is_file():
        raise ValidationError("source.md is missing", code="missing_markdown")
    try:
        markdown = markdown_path.read_text(encoding="utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        raise ValidationError(f"source.md is not valid UTF-8: {exc}", code="invalid_utf8") from exc
    if not markdown.strip():
        raise ValidationError("source.md is empty", code="sparse_markdown")
    non_whitespace = len(re.sub(r"\s+", "", markdown))
    if non_whitespace < 40:
        raise ValidationError(
            f"source.md is implausibly sparse ({non_whitespace} non-whitespace characters)",
            code="sparse_markdown",
        )
    if "\x00" in markdown:
        raise ValidationError("source.md contains NUL bytes", code="corrupt_markdown")

    local_targets: list[str] = []
    for match in _IMAGE_REF.finditer(markdown):
        target, _ = _decode_image_target(match.group("target"))
        if not _is_remote_target(target):
            local_targets.append(target)
    missing: list[str] = []
    resolved = 0
    for target in local_targets:
        candidate = (root / target).resolve(strict=False)
        if not is_within(candidate, root):
            missing.append(f"unsafe:{target}")
        elif candidate.is_file():
            resolved += 1
        else:
            missing.append(target)
    warnings: list[str] = []
    if non_whitespace < 120:
        warnings.append(f"sparse_markdown:{non_whitespace}")
    if missing:
        warnings.append(f"missing_image_references:{len(missing)}/{len(local_targets)}")
        if len(missing) >= 3 and len(missing) * 2 > len(local_targets):
            raise ValidationError(
                f"catastrophic image-reference loss: {len(missing)}/{len(local_targets)} missing",
                code="missing_images",
            )
    return ArtifactValidation(
        markdown_chars=len(markdown),
        chunk_count=0,
        image_references=len(local_targets),
        resolved_images=resolved,
        missing_images=missing,
        warnings=warnings,
    )


def validate_chunks(chunks_path: Path) -> int:
    if not chunks_path.is_file():
        raise ValidationError("required chunks.json is missing", code="missing_chunks")
    try:
        chunks = read_json(chunks_path)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValidationError(f"chunks.json is invalid: {exc}", code="invalid_chunks") from exc
    if not isinstance(chunks, list) or not chunks:
        raise ValidationError("chunks.json must be a non-empty array", code="invalid_chunks")
    for position, chunk in enumerate(chunks):
        if not isinstance(chunk, dict):
            raise ValidationError(f"chunk {position} is not an object", code="invalid_chunks")
        if not isinstance(chunk.get("index"), int) or not isinstance(chunk.get("text"), str):
            raise ValidationError(
                f"chunk {position} lacks MinerU index/text fields", code="invalid_chunks"
            )
        heading = chunk.get("heading", "")
        if not isinstance(heading, str):
            raise ValidationError(f"chunk {position} heading is invalid", code="invalid_chunks")
    return len(chunks)


def validate_artifacts(
    root: Path, *, chunks_path_override: Optional[Path] = None
) -> ArtifactValidation:
    """Validate an extraction plus one selected retrieval artifact."""
    root = root.resolve(strict=False)
    chunks_path = (
        chunks_path_override.resolve(strict=False)
        if chunks_path_override is not None
        else ensure_within(root / "chunks.json", root, "chunks.json")
    )
    validation = validate_extraction_artifacts(root)
    validation.chunk_count = validate_chunks(chunks_path)
    return validation


def artifact_hashes(cache_root: Path) -> dict[str, dict[str, Any]]:
    """Hash immutable extraction artifacts only; retrievals are separate."""
    result: dict[str, dict[str, Any]] = {}
    candidates = [cache_root / "source.md"]
    for artifact_root in (cache_root / "assets", cache_root / "normalized"):
        if artifact_root.is_dir():
            candidates.extend(path for path in artifact_root.rglob("*") if path.is_file())
    for path in sorted(candidates):
        if not path.is_file():
            continue
        relative = path.relative_to(cache_root).as_posix()
        result[relative] = {"sha256": sha256_file(path), "size_bytes": path.stat().st_size}
    return result


def validate_manifest(manifest: dict[str, Any], cache_root: Path) -> None:
    common_required = {
        "schema_version",
        "document_id",
        "revision_id",
        "cache_key",
        "source_filename",
        "source_original_path",
        "source_size_bytes",
        "source_mtime_ns",
        "source_sha256",
        "course",
        "document_title",
        "mineru_version",
        "mineru_skill_fingerprint",
        "engine_requested",
        "api_requested",
        "model_requested",
        "engine_used",
        "api_used",
        "model_used",
        "ocr_used",
        "split_used",
        "started_at",
        "completed_at",
        "status",
        "markdown_path",
        "assets_path",
        "published_path",
        "warnings",
        "artifacts",
    }
    schema = str(manifest.get("schema_version"))
    if schema == LEGACY_PIPELINE_SCHEMA_VERSION:
        required = common_required | {
            "profile",
            "profile_fingerprint",
            "chunk_enabled",
            "chunk_size",
            "chunks_path",
        }
    elif schema == PIPELINE_SCHEMA_VERSION:
        required = common_required | {
            "profile",
            "profile_fingerprint",
            "extraction_profile",
            "extraction_profile_fingerprint",
        }
    else:
        raise ValidationError("manifest schema mismatch", code="invalid_manifest")
    missing = sorted(required - set(manifest))
    if missing:
        raise ValidationError(f"manifest missing fields: {', '.join(missing)}", code="invalid_manifest")
    if schema == LEGACY_PIPELINE_SCHEMA_VERSION:
        expected_key = legacy_extraction_cache_key(
            manifest["source_sha256"], manifest["profile_fingerprint"]
        )
    else:
        fingerprint = manifest["extraction_profile_fingerprint"]
        if manifest["profile_fingerprint"] != fingerprint:
            raise ValidationError("manifest profile aliases disagree", code="invalid_manifest")
        if manifest["profile"] != manifest["extraction_profile"]:
            raise ValidationError("manifest extraction profile aliases disagree", code="invalid_manifest")
        if extraction_profile_fingerprint(manifest["extraction_profile"]) != fingerprint:
            raise ValidationError("manifest extraction profile fingerprint is inconsistent", code="invalid_manifest")
        expected_key = extraction_cache_key(manifest["source_sha256"], fingerprint)
    if manifest["cache_key"] != expected_key:
        raise ValidationError("manifest cache_key is inconsistent", code="invalid_manifest")
    artifacts = manifest.get("artifacts")
    if not isinstance(artifacts, dict):
        raise ValidationError("manifest artifacts must be an object", code="invalid_manifest")
    for relative, metadata in artifacts.items():
        path = ensure_within(cache_root / relative, cache_root, "manifest artifact")
        if not path.is_file() or not isinstance(metadata, dict):
            raise ValidationError(f"manifest artifact is missing: {relative}", code="invalid_manifest")
        if metadata.get("sha256") != sha256_file(path):
            raise ValidationError(f"manifest artifact hash mismatch: {relative}", code="invalid_manifest")
        if schema == PIPELINE_SCHEMA_VERSION and (
            relative == "chunks.json" or relative.startswith("retrieval/")
        ):
            raise ValidationError(
                "retrieval artifact is incorrectly bound into extraction manifest",
                code="invalid_manifest",
            )
    extraction = manifest.get("extraction_profile")
    if schema == PIPELINE_SCHEMA_VERSION and isinstance(extraction, dict) and extraction.get(
        "normalization"
    ) == "pptx_to_pdf":
        original = manifest.get("original_source")
        normalized = manifest.get("normalized_source")
        if not isinstance(original, dict) or not isinstance(normalized, dict):
            raise ValidationError(
                "normalized PPTX manifest lacks original/normalized provenance",
                code="invalid_manifest",
            )
        if original.get("source_sha256") != manifest.get("source_sha256"):
            raise ValidationError(
                "original PPTX provenance hash is inconsistent", code="invalid_manifest"
            )
        relative = normalized.get("path")
        if not isinstance(relative, str):
            raise ValidationError("normalized PDF path is invalid", code="invalid_manifest")
        normalized_path = ensure_within(cache_root / relative, cache_root, "normalized PDF")
        if not normalized_path.is_file() or normalized_path.stat().st_size <= 0:
            raise ValidationError("normalized PDF is missing or empty", code="invalid_manifest")
        actual_hash = sha256_file(normalized_path)
        if normalized.get("sha256") != actual_hash or manifest.get(
            "normalized_pdf_sha256"
        ) != actual_hash:
            raise ValidationError("normalized PDF hash mismatch", code="invalid_manifest")
        try:
            actual_pages = readable_pdf_page_count(normalized_path)
        except NormalizationError as exc:
            raise ValidationError(str(exc), code="invalid_manifest") from exc
        if normalized.get("page_count") != actual_pages:
            raise ValidationError("normalized PDF page count mismatch", code="invalid_manifest")


def validate_cache(cache_dir: Path, source_sha: str, fingerprint: str) -> tuple[dict[str, Any], ArtifactValidation]:
    manifest_path = cache_dir / "manifest.json"
    if not manifest_path.is_file():
        raise ValidationError("cache manifest is missing", code="invalid_cache")
    try:
        manifest = read_json(manifest_path)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValidationError(f"cache manifest is invalid: {exc}", code="invalid_cache") from exc
    if manifest.get("source_sha256") != source_sha:
        raise ValidationError("cache source hash mismatch", code="invalid_cache")
    if str(manifest.get("schema_version")) != PIPELINE_SCHEMA_VERSION:
        raise ValidationError("cache is not the current extraction schema", code="legacy_cache")
    if manifest.get("extraction_profile_fingerprint") != fingerprint:
        raise ValidationError("cache profile mismatch", code="invalid_cache")
    validation = validate_extraction_artifacts(cache_dir)
    validate_manifest(manifest, cache_dir)
    return manifest, validation


def find_compatible_v1_cache(
    cache_root: Path,
    *,
    source_sha: str,
    desired_extraction_profile: dict[str, Any],
) -> Optional[Path]:
    """Find a validated v1 cache whose parsing settings match the v2 extraction profile."""
    if not cache_root.is_dir():
        return None
    for manifest_path in sorted(cache_root.glob("*/manifest.json"), key=lambda path: str(path)):
        cache_dir = manifest_path.parent
        try:
            manifest = read_json(manifest_path)
            if str(manifest.get("schema_version")) != LEGACY_PIPELINE_SCHEMA_VERSION:
                continue
            if manifest.get("source_sha256") != source_sha:
                continue
            legacy_profile = manifest.get("profile")
            if not isinstance(legacy_profile, dict):
                continue
            if extraction_profile(legacy_profile) != desired_extraction_profile:
                continue
            validate_artifacts(cache_dir)
            validate_manifest(manifest, cache_dir)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, ValidationError, ValueError):
            continue
        return cache_dir
    return None


def prepare_v1_migration(
    legacy_cache: Path,
    prepared_dir: Path,
    *,
    cache_key: str,
    extraction_profile_value: dict[str, Any],
    extraction_fingerprint: str,
) -> tuple[dict[str, Any], ArtifactValidation]:
    """Materialize a v2 extraction cache from validated v1 artifacts, without MinerU."""
    legacy_manifest = read_json(legacy_cache / "manifest.json")
    validate_artifacts(legacy_cache)
    validate_manifest(legacy_manifest, legacy_cache)
    prepared_dir.mkdir(parents=True, exist_ok=False)
    shutil.copy2(legacy_cache / "source.md", prepared_dir / "source.md")
    if (legacy_cache / "assets").is_dir():
        shutil.copytree(legacy_cache / "assets", prepared_dir / "assets")
    else:
        (prepared_dir / "assets").mkdir()
    if (legacy_cache / "extraction.json").is_file():
        shutil.copy2(legacy_cache / "extraction.json", prepared_dir / "extraction.json")
    validation = validate_extraction_artifacts(prepared_dir)
    warnings = list(legacy_manifest.get("warnings", []))
    warnings.append("legacy_v1_cache_migrated_without_mineru")
    manifest = {
        "schema_version": PIPELINE_SCHEMA_VERSION,
        "document_id": legacy_manifest["document_id"],
        "revision_id": revision_id_for(cache_key),
        "cache_key": cache_key,
        "source_filename": legacy_manifest["source_filename"],
        "source_original_path": legacy_manifest["source_original_path"],
        "source_size_bytes": legacy_manifest["source_size_bytes"],
        "source_mtime_ns": legacy_manifest["source_mtime_ns"],
        "source_sha256": legacy_manifest["source_sha256"],
        "source_retained_path": legacy_manifest.get("source_retained_path"),
        "course": legacy_manifest["course"],
        "document_title": legacy_manifest["document_title"],
        "profile": extraction_profile_value,
        "profile_fingerprint": extraction_fingerprint,
        "extraction_profile": extraction_profile_value,
        "extraction_profile_fingerprint": extraction_fingerprint,
        "mineru_version": legacy_manifest["mineru_version"],
        "mineru_skill_fingerprint": legacy_manifest["mineru_skill_fingerprint"],
        "engine_requested": legacy_manifest["engine_requested"],
        "api_requested": legacy_manifest["api_requested"],
        "model_requested": legacy_manifest["model_requested"],
        "engine_used": legacy_manifest["engine_used"],
        "api_used": legacy_manifest["api_used"],
        "model_used": legacy_manifest["model_used"],
        "ocr_used": legacy_manifest["ocr_used"],
        "split_used": legacy_manifest["split_used"],
        "started_at": legacy_manifest["started_at"],
        "completed_at": legacy_manifest["completed_at"],
        "status": "validated",
        "markdown_path": "source.md",
        "assets_path": "assets",
        "published_path": None,
        "warnings": list(dict.fromkeys(warnings)),
        "validation": validation.to_dict(),
        "artifacts": artifact_hashes(prepared_dir),
    }
    atomic_write_json(prepared_dir / "manifest.json", manifest)
    validate_manifest(manifest, prepared_dir)
    return manifest, validation


def _swap_directory(new_dir: Path, destination: Path) -> Optional[Path]:
    backup = destination.parent / f".{destination.name}.backup-{uuid.uuid4().hex}"
    had_destination = destination.exists()
    if had_destination:
        destination.rename(backup)
    try:
        new_dir.rename(destination)
    except Exception:
        if had_destination and backup.exists() and not destination.exists():
            backup.rename(destination)
        raise
    return backup if had_destination else None


def commit_cache(prepared_dir: Path, cache_dir: Path) -> None:
    cache_dir.parent.mkdir(parents=True, exist_ok=True)
    incoming = cache_dir.parent / f".{cache_dir.name}.incoming-{uuid.uuid4().hex}"
    shutil.copytree(prepared_dir, incoming)
    existed_before = cache_dir.exists()
    backup: Optional[Path] = None
    try:
        backup = _swap_directory(incoming, cache_dir)
        validate_extraction_artifacts(cache_dir)
        validate_manifest(read_json(cache_dir / "manifest.json"), cache_dir)
    except Exception:
        if backup is not None:
            if cache_dir.exists():
                shutil.rmtree(cache_dir, ignore_errors=True)
            if backup.exists():
                backup.rename(cache_dir)
        elif not existed_before and cache_dir.exists():
            shutil.rmtree(cache_dir, ignore_errors=True)
        if incoming.exists():
            shutil.rmtree(incoming, ignore_errors=True)
        raise
    else:
        if backup and backup.exists():
            shutil.rmtree(backup, ignore_errors=True)


def validate_retrieval_artifact(
    retrieval_dir: Path,
    *,
    cache_dir: Path,
    cache_key: str,
    source_sha: str,
    profile: dict[str, Any],
    fingerprint: str,
) -> ArtifactValidation:
    metadata_path = retrieval_dir / "retrieval.json"
    if not metadata_path.is_file():
        raise ValidationError("retrieval metadata is missing", code="invalid_retrieval")
    try:
        metadata = read_json(metadata_path)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValidationError(f"retrieval metadata is invalid: {exc}", code="invalid_retrieval") from exc
    required = {
        "schema_version",
        "extraction_cache_key",
        "source_sha256",
        "source_markdown_sha256",
        "retrieval_profile",
        "retrieval_profile_fingerprint",
        "chunks_path",
        "chunks_sha256",
        "chunk_count",
    }
    missing = sorted(required - set(metadata))
    if missing:
        raise ValidationError(
            f"retrieval metadata missing fields: {', '.join(missing)}",
            code="invalid_retrieval",
        )
    if str(metadata["schema_version"]) != RETRIEVAL_PROFILE_SCHEMA_VERSION:
        raise ValidationError("retrieval schema mismatch", code="invalid_retrieval")
    if metadata["extraction_cache_key"] != cache_key or metadata["source_sha256"] != source_sha:
        raise ValidationError("retrieval extraction association mismatch", code="invalid_retrieval")
    if metadata["retrieval_profile"] != profile:
        raise ValidationError("retrieval profile mismatch", code="invalid_retrieval")
    if metadata["retrieval_profile_fingerprint"] != fingerprint:
        raise ValidationError("retrieval profile fingerprint mismatch", code="invalid_retrieval")
    if retrieval_profile_fingerprint(metadata["retrieval_profile"]) != fingerprint:
        raise ValidationError("retrieval profile fingerprint is inconsistent", code="invalid_retrieval")
    source_md = cache_dir / "source.md"
    if metadata["source_markdown_sha256"] != sha256_file(source_md):
        raise ValidationError("retrieval source.md association mismatch", code="invalid_retrieval")
    chunks_path = ensure_within(retrieval_dir / str(metadata["chunks_path"]), retrieval_dir, "chunks")
    validation = validate_artifacts(cache_dir, chunks_path_override=chunks_path)
    if metadata["chunks_sha256"] != sha256_file(chunks_path):
        raise ValidationError("retrieval chunks hash mismatch", code="invalid_retrieval")
    if metadata["chunk_count"] != validation.chunk_count:
        raise ValidationError("retrieval chunk count mismatch", code="invalid_retrieval")
    return validation


def ensure_retrieval_artifact(
    cache_dir: Path,
    *,
    cache_key: str,
    source_sha: str,
    profile: dict[str, Any],
    fingerprint: str,
    chunker_script: Path,
) -> tuple[Path, ArtifactValidation, bool]:
    """Create or validate a retrieval derivation without invoking MinerU cloud extraction."""
    if not profile.get("chunk_enabled"):
        raise PipelineError("the installed study reader requires chunking to remain enabled")
    target = cache_dir / "retrieval" / fingerprint
    if target.is_dir():
        try:
            validation = validate_retrieval_artifact(
                target,
                cache_dir=cache_dir,
                cache_key=cache_key,
                source_sha=source_sha,
                profile=profile,
                fingerprint=fingerprint,
            )
            return target / "chunks.json", validation, True
        except ValidationError:
            pass

    if not chunker_script.is_file():
        raise PipelineError(f"MinerU chunking module is missing: {chunker_script}")
    namespace = runpy.run_path(str(chunker_script), run_name="_study_docs_mineru_chunking")
    chunk_markdown = namespace.get("chunk_markdown")
    if not callable(chunk_markdown):
        raise PipelineError("MinerU chunking module does not expose chunk_markdown")
    markdown = (cache_dir / "source.md").read_text(encoding="utf-8", errors="strict")
    chunks = chunk_markdown(
        markdown,
        max_chars=int(profile["chunk_size"]),
        source="source.md",
    )
    if not isinstance(chunks, list) or not chunks:
        raise ValidationError("MinerU chunker produced no chunks", code="invalid_chunks")

    target.parent.mkdir(parents=True, exist_ok=True)
    incoming = target.parent / f".{target.name}.incoming-{uuid.uuid4().hex}"
    incoming.mkdir()
    try:
        atomic_write_json(incoming / "chunks.json", chunks)
        metadata = {
            "schema_version": RETRIEVAL_PROFILE_SCHEMA_VERSION,
            "extraction_cache_key": cache_key,
            "source_sha256": source_sha,
            "source_markdown_sha256": sha256_file(cache_dir / "source.md"),
            "retrieval_profile": profile,
            "retrieval_profile_fingerprint": fingerprint,
            "chunks_path": "chunks.json",
            "chunks_sha256": sha256_file(incoming / "chunks.json"),
            "chunk_count": len(chunks),
            "generated_at": utc_now(),
        }
        atomic_write_json(incoming / "retrieval.json", metadata)
        validation = validate_retrieval_artifact(
            incoming,
            cache_dir=cache_dir,
            cache_key=cache_key,
            source_sha=source_sha,
            profile=profile,
            fingerprint=fingerprint,
        )
        backup = _swap_directory(incoming, target)
        try:
            validation = validate_retrieval_artifact(
                target,
                cache_dir=cache_dir,
                cache_key=cache_key,
                source_sha=source_sha,
                profile=profile,
                fingerprint=fingerprint,
            )
        except Exception:
            if target.exists():
                shutil.rmtree(target, ignore_errors=True)
            if backup and backup.exists():
                backup.rename(target)
            raise
        else:
            if backup and backup.exists():
                shutil.rmtree(backup, ignore_errors=True)
            return target / "chunks.json", validation, False
    finally:
        if incoming.exists():
            shutil.rmtree(incoming, ignore_errors=True)


def publication_artifact_hashes(root: Path) -> dict[str, dict[str, Any]]:
    """Hash only the human-facing source.md/assets publication tree."""
    result: dict[str, dict[str, Any]] = {}
    candidates = [root / "source.md"]
    assets = root / "assets"
    if assets.is_dir():
        candidates.extend(path for path in assets.rglob("*") if path.is_file())
    for path in sorted(candidates):
        if path.is_file():
            relative = path.relative_to(root).as_posix()
            result[relative] = {
                "sha256": sha256_file(path),
                "size_bytes": path.stat().st_size,
            }
    return result


def publication_tree_fingerprint(root: Path) -> str:
    return sha256_bytes(canonical_json(publication_artifact_hashes(root)).encode("utf-8"))


def _unexpected_publication_entries(root: Path) -> list[str]:
    if not root.is_dir():
        return []
    return sorted(
        child.name for child in root.iterdir() if child.name not in {"source.md", "assets"}
    )


def plan_legacy_publication_migration(
    *,
    vault: Path,
    catalog: dict[str, Any],
    publication_root: str = DEFAULT_PUBLICATION_ROOT,
) -> dict[str, Any]:
    """Return a read-only, evidence-backed plan; never copy, move, or delete files."""
    vault = vault.expanduser().resolve(strict=False)
    library_root = ensure_within(vault / publication_root, vault, "publication root")
    candidates: dict[str, Path] = {}
    if vault.is_dir():
        for course_dir in sorted(path for path in vault.iterdir() if path.is_dir()):
            legacy_sources = course_dir / "Sources"
            if legacy_sources.is_dir():
                for child in sorted(path for path in legacy_sources.iterdir() if path.is_dir()):
                    candidates[normalized_path_key(child)] = child.resolve(strict=False)
    for document in catalog.get("documents", {}).values():
        publication = document.get("publication") if isinstance(document, dict) else None
        raw = publication.get("source_path") if isinstance(publication, dict) else None
        if not raw:
            continue
        source_path = Path(str(raw)).expanduser().resolve(strict=False)
        if is_within(source_path, vault) and not is_within(source_path, library_root):
            candidates[normalized_path_key(source_path)] = source_path

    entries: list[dict[str, Any]] = []
    documents = catalog.get("documents", {})
    for candidate in sorted(candidates.values(), key=normalized_path_key):
        matches: list[tuple[str, dict[str, Any]]] = []
        for document_id, document in documents.items():
            publication = document.get("publication") if isinstance(document, dict) else None
            raw = publication.get("source_path") if isinstance(publication, dict) else None
            if raw and normalized_path_key(raw) == normalized_path_key(candidate):
                matches.append((document_id, document))
        entry: dict[str, Any] = {
            "legacy_source_path": str(candidate),
            "status": "REFUSED_AMBIGUOUS",
            "reasons": [],
            "operation_performed": False,
        }
        if len(matches) != 1:
            entry["reasons"].append("catalog_ownership_not_unique")
            entries.append(entry)
            continue
        document_id, document = matches[0]
        publication = document.get("publication", {})
        current_revision_id = document.get("current_published_revision_id")
        revisions = document.get("revisions", [])
        revision = next(
            (
                item
                for item in revisions
                if isinstance(item, dict) and item.get("revision_id") == current_revision_id
            ),
            None,
        )
        entry.update(
            {
                "document_id": document_id,
                "course": document.get("course"),
                "title": document.get("title"),
                "revision_id": current_revision_id,
            }
        )
        if not candidate.is_dir():
            entry["reasons"].append("legacy_source_missing")
        unexpected = _unexpected_publication_entries(candidate)
        if unexpected:
            entry["reasons"].append("unexpected_files_or_folders:" + ",".join(unexpected))
        if not isinstance(revision, dict):
            entry["reasons"].append("published_revision_missing_from_catalog")
            cache_dir = None
        else:
            cache_dir = Path(str(revision.get("cache_path", ""))).expanduser().resolve(strict=False)
            if not cache_dir.is_dir():
                entry["reasons"].append("cache_missing")
            else:
                try:
                    validate_extraction_artifacts(cache_dir)
                    validate_manifest(read_json(cache_dir / "manifest.json"), cache_dir)
                except (OSError, ValueError, ValidationError, json.JSONDecodeError) as exc:
                    entry["reasons"].append(f"cache_validation_failed:{type(exc).__name__}")
                if publication_artifact_hashes(candidate) != publication_artifact_hashes(cache_dir):
                    entry["reasons"].append("published_tree_differs_from_cache")
        metadata_raw = publication.get("metadata_path") if isinstance(publication, dict) else None
        if not metadata_raw:
            entry["reasons"].append("legacy_manifest_not_recorded")
        else:
            metadata_path = Path(str(metadata_raw)).expanduser().resolve(strict=False)
            manifest_path = metadata_path / "manifest.json"
            if not is_within(metadata_path, vault) or not manifest_path.is_file():
                entry["reasons"].append("legacy_manifest_missing_or_outside_vault")
            else:
                try:
                    legacy_manifest = read_json(manifest_path)
                except (OSError, UnicodeDecodeError, json.JSONDecodeError):
                    entry["reasons"].append("legacy_manifest_invalid")
                else:
                    expected_cache_key = revision.get("cache_key") if isinstance(revision, dict) else None
                    if (
                        legacy_manifest.get("document_id") != document_id
                        or legacy_manifest.get("revision_id") != current_revision_id
                        or legacy_manifest.get("cache_key") != expected_cache_key
                        or normalized_path_key(legacy_manifest.get("published_path", ""))
                        != normalized_path_key(candidate)
                    ):
                        entry["reasons"].append("legacy_manifest_identity_mismatch")
        directory_name = safe_component(
            publication.get("directory_name") or document.get("title") or candidate.name
        )
        destination = ensure_within(
            library_root
            / safe_component(str(document.get("course") or DEFAULT_FALLBACK_COURSE))
            / directory_name,
            library_root,
            "migration destination",
        )
        entry["destination_path"] = str(destination)
        if destination.exists():
            entry["reasons"].append("destination_already_exists")
        if not entry["reasons"]:
            entry.update(
                {
                    "status": "SAFE_TO_MIGRATE",
                    "evidence": [
                        "unique_catalog_publication",
                        "matching_document_revision_cache_manifest",
                        "source_and_asset_hashes_match_cache",
                        "no_unexpected_visible_entries",
                        "destination_absent",
                    ],
                    "recommended_action": (
                        "republish_from_validated_cache_then_remove_legacy_source_and "
                        "legacy_vault_metadata only after explicit approval"
                    ),
                }
            )
        entries.append(entry)
    return {
        "status": "PLAN_ONLY",
        "vault": str(vault),
        "publication_root": str(library_root),
        "reserved_notes_root": str(vault / RESERVED_NOTES_ROOT),
        "safe_count": sum(entry["status"] == "SAFE_TO_MIGRATE" for entry in entries),
        "refused_count": sum(entry["status"] == "REFUSED_AMBIGUOUS" for entry in entries),
        "entries": entries,
        "operation_performed": False,
    }


def publish_revision(
    cache_dir: Path,
    *,
    vault: Path,
    publication_root: str,
    course: str,
    document_directory: str,
    expected_existing_fingerprint: Optional[str] = None,
    fault_injector: Optional[Callable[[str], None]] = None,
) -> dict[str, Any]:
    vault = vault.resolve(strict=False)
    if not vault.is_dir():
        raise PublicationError(f"vault does not exist: {vault}")
    if publication_root.casefold() == RESERVED_NOTES_ROOT.casefold():
        raise PublicationError(f"refusing to publish into reserved notes root: {publication_root}")
    if publication_root.casefold() != DEFAULT_PUBLICATION_ROOT.casefold():
        raise PublicationError(
            f"refusing to publish outside the required source root: {DEFAULT_PUBLICATION_ROOT}"
        )
    library_root = ensure_within(vault / publication_root, vault, "publication root")
    source_parent = ensure_within(
        library_root / safe_component(course, fallback=DEFAULT_FALLBACK_COURSE),
        library_root,
        "publication course",
    )
    source_destination = ensure_within(
        source_parent / safe_component(document_directory),
        library_root,
        "publication source",
    )
    source_parent.mkdir(parents=True, exist_ok=True)
    if source_destination.exists():
        actual_existing_fingerprint = publication_tree_fingerprint(source_destination)
        if (
            not expected_existing_fingerprint
            or actual_existing_fingerprint != expected_existing_fingerprint
            or _unexpected_publication_entries(source_destination)
        ):
            raise PublicationError(
                "refusing to overwrite an unrecognized or user-modified publication directory: "
                f"{source_destination}"
            )

    nonce = uuid.uuid4().hex
    source_new = source_parent / f".{source_destination.name}.new-{nonce}"
    source_existed = source_destination.exists()
    source_backup: Optional[Path] = None
    try:
        source_new.mkdir()
        shutil.copy2(cache_dir / "source.md", source_new / "source.md")
        if (cache_dir / "assets").is_dir():
            shutil.copytree(cache_dir / "assets", source_new / "assets")
        validate_extraction_artifacts(source_new)
        if fault_injector:
            fault_injector("before_source_swap")
        source_backup = _swap_directory(source_new, source_destination)
        if fault_injector:
            fault_injector("after_source_swap")
    except Exception as exc:
        if source_backup is not None:
            if source_destination.exists():
                shutil.rmtree(source_destination, ignore_errors=True)
            if source_backup.exists():
                source_backup.rename(source_destination)
        elif not source_existed and source_destination.exists():
            shutil.rmtree(source_destination, ignore_errors=True)
        if source_new.exists():
            shutil.rmtree(source_new, ignore_errors=True)
        raise PublicationError(f"publication failed and was rolled back: {exc}") from exc
    else:
        if source_backup and source_backup.exists():
            shutil.rmtree(source_backup, ignore_errors=True)
        return {
            "source_path": str(source_destination),
            "metadata_path": None,
            "tree_fingerprint": publication_tree_fingerprint(source_destination),
        }


def _document_aliases(document: dict[str, Any]) -> Iterable[dict[str, Any]]:
    aliases = document.get("source_aliases", [])
    return aliases if isinstance(aliases, list) else []


def resolve_document_id(
    catalog: dict[str, Any],
    *,
    source_path: Path,
    source_sha: str,
    course: str,
    title: str,
) -> str:
    key = normalized_path_key(source_path)
    exact: list[str] = []
    identical: list[str] = []
    identical_context: list[str] = []
    for document_id, document in catalog["documents"].items():
        if any(alias.get("path_key") == key for alias in _document_aliases(document)):
            exact.append(document_id)
        has_identical_content = any(
            alias.get("source_sha256") == source_sha for alias in _document_aliases(document)
        )
        if has_identical_content:
            identical.append(document_id)
        if (
            document.get("course", "").casefold() == course.casefold()
            and document.get("title", "").casefold() == title.casefold()
            and has_identical_content
        ):
            identical_context.append(document_id)
    if len(exact) == 1:
        return exact[0]
    if len(identical_context) == 1:
        return identical_context[0]
    if len(identical) == 1:
        return identical[0]
    stable = sha256_bytes(key.encode("utf-8"))[:24]
    return f"doc-{stable}"


def revision_id_for(cache_key: str) -> str:
    return f"rev-{cache_key[:24]}"


def publication_directory_name(
    catalog: dict[str, Any],
    *,
    document_id: str,
    course: str,
    title: str,
    vault: Path,
    publication_root: str,
) -> str:
    existing = catalog["documents"].get(document_id, {})
    publication = existing.get("publication", {}) if isinstance(existing, dict) else {}
    if isinstance(publication, dict) and publication.get("directory_name"):
        return safe_component(publication["directory_name"])
    base = safe_component(title)
    proposed = (
        vault
        / publication_root
        / safe_component(course, fallback=DEFAULT_FALLBACK_COURSE)
        / base
    )
    proposed_key = normalized_path_key(proposed)
    collision = proposed.exists()
    for other_id, document in catalog["documents"].items():
        if other_id == document_id:
            continue
        other_publication = document.get("publication", {})
        if isinstance(other_publication, dict):
            source_path = other_publication.get("source_path")
            if source_path and normalized_path_key(source_path) == proposed_key:
                collision = True
                break
    return f"{base}--{document_id[-8:]}" if collision else base


def update_catalog(
    catalog: dict[str, Any],
    *,
    document_id: str,
    revision_id: str,
    source_path: Path,
    source_sha: str,
    course: str,
    title: str,
    cache_key: str,
    extraction_fingerprint: str,
    retrieval_profile_value: dict[str, Any],
    retrieval_fingerprint: str,
    chunks_path: Path,
    cache_dir: Path,
    publication: Optional[dict[str, Any]],
    publication_directory: Optional[str],
) -> dict[str, Any]:
    documents = catalog["documents"]
    document = documents.setdefault(
        document_id,
        {
            "document_id": document_id,
            "course": course,
            "title": title,
            "source_aliases": [],
            "current_source_binding": None,
            "revisions": [],
            "current_revision_id": None,
            "current_published_revision_id": None,
            "publication": None,
        },
    )
    document["course"] = course
    document["title"] = title
    alias_key = normalized_path_key(source_path)
    seen_at = utc_now()
    binding = {
        "path": str(source_path),
        "path_key": alias_key,
        "filename": source_path.name,
        "source_sha256": source_sha,
        "last_seen_at": seen_at,
    }
    aliases = [a for a in document.get("source_aliases", []) if a.get("path_key") != alias_key]
    aliases.append(binding)
    document["source_aliases"] = aliases
    document["current_source_binding"] = dict(binding)
    revision = {
        "revision_id": revision_id,
        "cache_key": cache_key,
        "source_sha256": source_sha,
        "profile_fingerprint": extraction_fingerprint,
        "extraction_profile_fingerprint": extraction_fingerprint,
        "retrieval_profile": retrieval_profile_value,
        "retrieval_profile_fingerprint": retrieval_fingerprint,
        "chunks_path": str(chunks_path),
        "cache_path": str(cache_dir),
        "validated_at": utc_now(),
        "published": publication is not None,
    }
    revisions = [r for r in document.get("revisions", []) if r.get("revision_id") != revision_id]
    revisions.append(revision)
    document["revisions"] = revisions
    document["current_revision_id"] = revision_id
    if publication is not None:
        document["current_published_revision_id"] = revision_id
        document["publication"] = {
            **publication,
            "directory_name": publication_directory,
        }
    cache_entry = catalog["cache_entries"].get(cache_key, {})
    retrieval_profiles = cache_entry.get("retrieval_profiles", {})
    if not isinstance(retrieval_profiles, dict):
        retrieval_profiles = {}
    retrieval_profiles[retrieval_fingerprint] = {
        "retrieval_profile": retrieval_profile_value,
        "chunks_path": str(chunks_path),
        "validated_at": utc_now(),
    }
    catalog["cache_entries"][cache_key] = {
        "cache_key": cache_key,
        "cache_path": str(cache_dir),
        "source_sha256": source_sha,
        "profile_fingerprint": extraction_fingerprint,
        "extraction_profile_fingerprint": extraction_fingerprint,
        "retrieval_profiles": retrieval_profiles,
        "validated_at": utc_now(),
    }
    catalog["updated_at"] = utc_now()
    return catalog


def archive_source(state_dir: Path, source: Path, source_sha: str) -> Path:
    destination_dir = state_dir / "sources" / source_sha
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / safe_component(source.name)
    if destination.exists():
        if sha256_file(destination) != source_sha:
            destination = destination_dir / f"{safe_component(source.stem)}--{source_sha[:8]}{source.suffix}"
    if not destination.exists():
        shutil.copy2(source, destination)
    return destination


def _safe_command_for_record(command: list[str]) -> list[str]:
    safe: list[str] = []
    redact_next = False
    for argument in command:
        if redact_next:
            safe.append("[REDACTED]")
            redact_next = False
            continue
        safe.append(_redact_text(argument))
        if argument == "--token":
            redact_next = True
    return safe


class IngestionPipeline:
    def __init__(
        self,
        *,
        state_dir: Path | str = DEFAULT_STATE_DIR,
        python_executable: Path | str | None = None,
        mineru_script: Path | str | None = None,
        chunker_script: Path | str | None = None,
        powerpoint_script: Path | str | None = None,
        initialize: bool = True,
        local_engine_available_override: Optional[bool] = None,
        normalizer_detector: Optional[Callable[[], list[NormalizerBackend]]] = None,
        normalizer_converter: Optional[
            Callable[[Path, Path, NormalizerBackend], dict[str, Any]]
        ] = None,
    ):
        self.state_dir = Path(state_dir).expanduser().resolve(strict=False)
        configured_python = os.environ.get("STUDY_DOCS_PYTHON")
        configured_mineru = os.environ.get("STUDY_DOCS_MINERU_SCRIPT")
        self.python_executable = Path(
            python_executable or configured_python or sys.executable
        ).expanduser().resolve(strict=False)
        self.mineru_script = Path(
            mineru_script or configured_mineru or default_mineru_script()
        ).expanduser().resolve(strict=False)
        configured_chunker = os.environ.get("STUDY_DOCS_MINERU_CHUNKER")
        sibling_chunker = self.mineru_script.with_name("chunking.py")
        discovered_chunker = default_mineru_script().with_name("chunking.py")
        self.chunker_script = Path(
            chunker_script
            or configured_chunker
            or (sibling_chunker if sibling_chunker.is_file() else discovered_chunker)
        ).expanduser().resolve(strict=False)
        self.powerpoint_script = Path(
            powerpoint_script or Path(__file__).with_name("export_pptx_pdf.ps1")
        ).expanduser().resolve(strict=False)
        self.local_engine_available_override = local_engine_available_override
        self.normalizer_detector = normalizer_detector
        self.normalizer_converter = normalizer_converter
        self._normalizer_cache: Optional[list[NormalizerBackend]] = None
        if initialize:
            initialize_state(self.state_dir)
        self.logger = PipelineLogger(self.state_dir / "logs") if initialize else None

    def config(self) -> dict[str, Any]:
        return load_config(self.state_dir)

    def catalog(self) -> dict[str, Any]:
        return load_catalog(self.state_dir)

    def local_engine_available(self) -> bool:
        if self.local_engine_available_override is not None:
            return self.local_engine_available_override
        return interpreter_has_module(self.python_executable, "pymupdf4llm")

    def available_normalizers(self) -> list[NormalizerBackend]:
        if self._normalizer_cache is None:
            detected = (
                self.normalizer_detector()
                if self.normalizer_detector is not None
                else detect_normalizer_backends(self.powerpoint_script)
            )
            by_name = {backend.name: backend for backend in detected}
            self._normalizer_cache = [
                by_name[name]
                for name in ("powerpoint", "libreoffice")
                if name in by_name
            ]
        return list(self._normalizer_cache)

    def _quality_pptx_profile(
        self, profile: ExtractionProfile, backend: NormalizerBackend
    ) -> ExtractionProfile:
        return dataclasses.replace(
            profile,
            input_modality="pptx",
            normalization="pptx_to_pdf",
            normalizer_backend=backend.name,
            normalizer_version=backend.version,
            normalization_schema_version=PPTX_NORMALIZATION_SCHEMA_VERSION,
        )

    def _normalize_pptx(
        self,
        *,
        source: Path,
        staging_dir: Path,
        backend: NormalizerBackend,
        expected_source_sha: str,
    ) -> NormalizationResult:
        destination = ensure_within(
            staging_dir / "pptx-normalization" / "source.pdf",
            staging_dir,
            "normalized PDF",
        )
        started = time.monotonic()
        if self.logger:
            self.logger.log(
                "normalization_start",
                source=str(source),
                backend=backend.name,
                backend_version=backend.version,
                destination=str(destination),
            )
        if self.normalizer_converter is not None:
            report = self.normalizer_converter(source, destination, backend)
        else:
            report = convert_pptx_to_pdf(
                source,
                destination,
                backend,
                powerpoint_script=self.powerpoint_script,
            )
        if not destination.is_file():
            raise NormalizationError("PPTX normalization produced no PDF")
        if destination.stat().st_size <= 0:
            raise NormalizationError("PPTX normalization produced an empty PDF")
        ensure_within(destination, staging_dir, "normalized PDF")
        page_count = readable_pdf_page_count(destination)
        slide_count = pptx_slide_count(source)
        reported_slide_count = report.get("slide_count") if isinstance(report, dict) else None
        if slide_count is None and isinstance(reported_slide_count, int) and reported_slide_count >= 0:
            slide_count = reported_slide_count
        warnings: list[str] = []
        if slide_count == 0:
            raise NormalizationError("source PPTX declares no slides")
        if slide_count is not None and slide_count != page_count:
            delta = abs(slide_count - page_count)
            ratio = min(slide_count, page_count) / max(slide_count, page_count)
            mismatch = f"pptx_pdf_count_mismatch:{slide_count}_slides/{page_count}_pages"
            if delta == 1 and ratio >= 0.9:
                warnings.append(mismatch)
            else:
                raise NormalizationError(
                    "implausible PPTX/PDF count mismatch: "
                    f"{slide_count} slides versus {page_count} pages"
                )
        if sha256_file(source) != expected_source_sha:
            raise NormalizationError("original PPTX changed during normalization")
        duration = time.monotonic() - started
        result = NormalizationResult(
            path=destination,
            backend=backend.name,
            version=backend.version,
            slide_count=slide_count,
            page_count=page_count,
            sha256=sha256_file(destination),
            duration=duration,
            warnings=warnings,
        )
        if self.logger:
            self.logger.log(
                "normalization_finish",
                source=str(source),
                backend=result.backend,
                backend_version=result.version,
                slide_count=result.slide_count,
                page_count=result.page_count,
                normalized_pdf_sha256=result.sha256,
                duration=round(result.duration, 3),
                warnings=result.warnings,
            )
        return result

    def build_mineru_command(
        self,
        source: Path,
        output_dir: Path,
        profile: ExtractionProfile,
        *,
        split: bool = False,
        ocr: bool = False,
    ) -> list[str]:
        if profile.mode == "private":
            if source.suffix.casefold() != ".pdf":
                raise PrivateModeError(
                    "private mode supports only local PDF extraction in MinerU 3.3.1; "
                    "Office/image inputs would risk cloud routing and are rejected"
                )
            if not self.local_engine_available():
                raise PrivateModeError(
                    "private mode requires the optional pymupdf4llm local engine; it is not installed. "
                    "No cloud fallback was attempted."
                )
        command = [
            str(self.python_executable),
            str(self.mineru_script),
            str(source),
            "--output",
            str(output_dir),
            "--engine",
            profile.engine,
        ]
        if profile.mode != "private":
            command.extend(["--api", profile.api, "--model", profile.model])
        command.extend(["--lang", profile.language_policy, "--workers", "1"])
        if split:
            command.append("--split")
        if ocr:
            command.append("--ocr")
        command.extend(["--json", "--quiet"])
        return command

    def _initial_split_decision(
        self, source: Path, profile: ExtractionProfile
    ) -> tuple[bool, Optional[int], Optional[str]]:
        if source.suffix.casefold() != ".pdf" or profile.mode == "private":
            return False, None, None
        pages = pdf_page_count(source)
        if pages is None:
            return False, None, "pdf_page_count_unavailable; MinerU will route normally"
        limits = load_mineru_limits(self.mineru_script)
        standard = profile.api == "standard" or (
            profile.api == "auto" and bool(os.environ.get("MINERU_TOKEN"))
        )
        cap = limits["standard" if standard else "agent"]
        return pages > cap, pages, None

    def _run_one_attempt(
        self,
        *,
        source: Path,
        attempt_dir: Path,
        profile: ExtractionProfile,
        split: bool,
        ocr: bool,
    ) -> tuple[ProcessOutcome, Path, ArtifactValidation, list[str]]:
        output_dir = attempt_dir / "mineru-output"
        output_dir.mkdir(parents=True, exist_ok=False)
        command = self.build_mineru_command(source, output_dir, profile, split=split, ocr=ocr)
        if self.logger:
            self.logger.log(
                "mineru_start",
                source=str(source),
                command=_safe_command_for_record(command),
                split=split,
                ocr=ocr,
            )
        outcome = run_captured(command)
        try:
            markdown_path = discover_markdown(output_dir, outcome.status)
        except ValidationError as exc:
            detail = _redact_text((outcome.stderr or outcome.stdout)[-1000:])
            raise PipelineError(f"{exc}; MinerU detail: {detail}") from exc
        normalized = attempt_dir / "normalized"
        normalization_warnings = normalize_mineru_artifacts(markdown_path, normalized)
        validation = validate_extraction_artifacts(normalized)
        if outcome.returncode != 0:
            if _ENCODING_FAILURE.search(outcome.stderr + "\n" + outcome.stdout):
                validation.warnings.append(
                    "MinerU console-status rendering failed after valid artifacts; artifacts accepted"
                )
            else:
                raise PipelineError(
                    "MinerU exited unsuccessfully despite artifacts: "
                    + _redact_text((outcome.stderr or outcome.stdout)[-1000:])
                )
        status_results = outcome.status.get("results") if outcome.status else None
        if isinstance(status_results, list):
            failures = [
                item
                for item in status_results
                if isinstance(item, dict) and item.get("state") == "failed"
            ]
            if failures:
                raise PipelineError(
                    "MinerU status reported failure: "
                    + _redact_text(str(failures[0].get("error") or "unknown failure"))
                )
        warnings = normalization_warnings + validation.warnings
        if self.logger:
            self.logger.log(
                "mineru_finish",
                source=str(source),
                returncode=outcome.returncode,
                duration=round(outcome.duration, 3),
                validation=validation.to_dict(),
                warnings=warnings,
                stderr_tail=_redact_text(outcome.stderr[-1000:]),
            )
        return outcome, normalized, validation, warnings

    def _extract(
        self,
        *,
        source: Path,
        staging_dir: Path,
        profile: ExtractionProfile,
    ) -> tuple[ProcessOutcome, Path, ArtifactValidation, dict[str, Any]]:
        split, page_count, split_warning = self._initial_split_decision(source, profile)
        attempts: list[dict[str, Any]] = []
        warnings = [split_warning] if split_warning else []
        ocr_used = False
        split_used = split
        attempt_index = 0

        def attempt(*, use_split: bool, use_ocr: bool):
            nonlocal attempt_index
            attempt_index += 1
            attempt_dir = staging_dir / f"attempt-{attempt_index}"
            attempt_dir.mkdir(parents=True, exist_ok=False)
            try:
                outcome, normalized, validation, attempt_warnings = self._run_one_attempt(
                    source=source,
                    attempt_dir=attempt_dir,
                    profile=profile,
                    split=use_split,
                    ocr=use_ocr,
                )
                attempts.append(
                    {
                        "attempt": attempt_index,
                        "split": use_split,
                        "ocr": use_ocr,
                        "returncode": outcome.returncode,
                        "duration": round(outcome.duration, 3),
                        "status_available": outcome.status is not None,
                        "warnings": attempt_warnings,
                    }
                )
                return outcome, normalized, validation, attempt_warnings
            except Exception as exc:
                attempts.append(
                    {
                        "attempt": attempt_index,
                        "split": use_split,
                        "ocr": use_ocr,
                        "error": _redact_text(str(exc))[:1000],
                    }
                )
                raise

        try:
            outcome, normalized, validation, attempt_warnings = attempt(
                use_split=split, use_ocr=False
            )
        except Exception as first_error:
            text = str(first_error)
            retry_split = (
                source.suffix.casefold() == ".pdf"
                and profile.mode != "private"
                and not split
                and _PAGE_LIMIT_FAILURE.search(text)
            )
            retry_ocr = (
                isinstance(first_error, ValidationError)
                and first_error.code == "sparse_markdown"
                and profile.mode == "quality"
                and source.suffix.casefold() in ({".pdf"} | IMAGE_SUFFIXES)
            )
            if retry_split:
                split_used = True
                outcome, normalized, validation, attempt_warnings = attempt(
                    use_split=True, use_ocr=False
                )
            elif retry_ocr:
                ocr_used = True
                outcome, normalized, validation, attempt_warnings = attempt(
                    use_split=split, use_ocr=True
                )
            else:
                raise

        # A technically valid but extremely sparse output gets one controlled OCR
        # fallback in quality mode.  The validated first result stays untouched until
        # the replacement also validates.
        sparse_warning = any(w.startswith("sparse_markdown:") for w in validation.warnings)
        if (
            sparse_warning
            and profile.mode == "quality"
            and source.suffix.casefold() in ({".pdf"} | IMAGE_SUFFIXES)
            and not ocr_used
        ):
            try:
                replacement = attempt(use_split=split_used, use_ocr=True)
            except Exception as exc:
                warnings.append(
                    "controlled OCR fallback failed; retained initial result: "
                    + _redact_text(str(exc))
                )
            else:
                outcome, normalized, validation, attempt_warnings = replacement
                ocr_used = True

        warnings.extend(attempt_warnings)
        metadata = {
            "attempts": attempts,
            "page_count": page_count,
            "split_used": split_used,
            "ocr_used": ocr_used,
            "warnings": list(dict.fromkeys(warnings)),
        }
        return outcome, normalized, validation, metadata

    def _backend_used(self, status: Optional[dict[str, Any]]) -> tuple[Optional[str], Optional[str]]:
        if not status:
            return None, None
        results = status.get("results")
        if not isinstance(results, list):
            return None, None
        for item in results:
            if isinstance(item, dict) and item.get("api"):
                api = str(item["api"])
                engine = "local" if api == "local" else "cloud" if api in {"agent", "standard"} else None
                return engine, api
        return None, None

    def _make_manifest(
        self,
        *,
        source: Path,
        source_sha: str,
        source_stat: os.stat_result,
        document_id: str,
        revision_id: str,
        cache_key: str,
        course: str,
        title: str,
        profile: ExtractionProfile,
        fingerprint: str,
        outcome: ProcessOutcome,
        normalized: Path,
        validation: ArtifactValidation,
        extraction_metadata: dict[str, Any],
        started_at: str,
        completed_at: str,
        retained_source: Optional[Path],
        normalization_result: Optional[NormalizationResult] = None,
    ) -> dict[str, Any]:
        engine_used, api_used = self._backend_used(outcome.status)
        warnings = list(
            dict.fromkeys(
                extraction_metadata.get("warnings", [])
                + (normalization_result.warnings if normalization_result else [])
                + validation.warnings
            )
        )
        extraction_profile_value = extraction_profile(profile)
        manifest = {
            "schema_version": PIPELINE_SCHEMA_VERSION,
            "document_id": document_id,
            "revision_id": revision_id,
            "cache_key": cache_key,
            "source_filename": source.name,
            "source_original_path": str(source),
            "source_size_bytes": source_stat.st_size,
            "source_mtime_ns": source_stat.st_mtime_ns,
            "source_sha256": source_sha,
            "source_retained_path": str(retained_source) if retained_source else None,
            "original_source": {
                "role": "canonical_original",
                "filename": source.name,
                "path_at_extraction": str(source),
                "source_sha256": source_sha,
                "current_binding_at_extraction": {
                    "path": str(source),
                    "observed_at": completed_at,
                },
                "current_binding_authority": "catalog",
            },
            "course": course,
            "document_title": title,
            "profile": extraction_profile_value,
            "profile_fingerprint": fingerprint,
            "extraction_profile": extraction_profile_value,
            "extraction_profile_fingerprint": fingerprint,
            "mineru_version": mineru_version_from_source(self.mineru_script),
            "mineru_skill_fingerprint": mineru_fingerprint(self.mineru_script),
            "engine_requested": profile.engine,
            "api_requested": profile.api,
            "model_requested": profile.model,
            "engine_used": engine_used,
            "api_used": api_used,
            "model_used": None,
            "ocr_used": extraction_metadata.get("ocr_used", False),
            "split_used": extraction_metadata.get("split_used", False),
            "started_at": started_at,
            "completed_at": completed_at,
            "status": "validated",
            "markdown_path": "source.md",
            "assets_path": "assets",
            "published_path": None,
            "warnings": warnings,
            "validation": validation.to_dict(),
            "artifacts": artifact_hashes(normalized),
        }
        if normalization_result is not None:
            manifest["normalized_pdf_sha256"] = normalization_result.sha256
            manifest["normalized_source"] = {
                "role": "extraction_derivative",
                "path": "normalized/source.pdf",
                "sha256": normalization_result.sha256,
                "backend": normalization_result.backend,
                "version": normalization_result.version,
                "slide_count": normalization_result.slide_count,
                "page_count": normalization_result.page_count,
                "duration_seconds": round(normalization_result.duration, 6),
            }
        return manifest

    def _record_extraction(
        self,
        normalized: Path,
        *,
        outcome: ProcessOutcome,
        metadata: dict[str, Any],
    ) -> None:
        status = outcome.status
        compact_status: Optional[dict[str, Any]] = None
        if isinstance(status, dict):
            compact_status = dict(status)
            results = compact_status.get("results")
            if isinstance(results, list):
                cleaned = []
                for item in results:
                    if isinstance(item, dict):
                        item = {k: v for k, v in item.items() if k != "chunks"}
                    cleaned.append(item)
                compact_status["results"] = cleaned
        record = {
            "schema_version": PIPELINE_SCHEMA_VERSION,
            "command": _safe_command_for_record(outcome.command),
            "returncode": outcome.returncode,
            "duration_seconds": round(outcome.duration, 3),
            "status": compact_status,
            "stderr_tail": _redact_text(outcome.stderr[-2000:]),
            "metadata": metadata,
        }
        atomic_write_json(normalized / "extraction.json", record)

    def ingest_file(
        self,
        source: Path | str,
        *,
        course: Optional[str] = None,
        title: Optional[str] = None,
        mode: Optional[str] = None,
        language: str = "ch",
        chunk_size: int = DEFAULT_CHUNK_SIZE,
        vault: Optional[Path | str] = None,
        refresh: bool = False,
        no_publish: bool = False,
        dry_run: bool = False,
        source_retention: Optional[str] = None,
        publication_fault_injector: Optional[Callable[[str], None]] = None,
    ) -> dict[str, Any]:
        total_started = time.monotonic()
        started_at = utc_now()
        source_path = Path(source).expanduser().resolve(strict=False)
        if not source_path.is_file():
            raise PipelineError(f"source is not a file: {source_path}")
        if source_path.suffix.casefold() not in SUPPORTED_SUFFIXES:
            raise PipelineError(f"unsupported source type: {source_path.suffix}")
        config = self.config()
        publication_root = configured_publication_root(config)
        chosen_mode = mode or config["default_mode"]
        profile = profile_for_mode(chosen_mode, language=language, chunk_size=chunk_size)
        normalization_backend: Optional[NormalizerBackend] = None
        if chosen_mode == "quality" and source_path.suffix.casefold() == ".pptx":
            available = self.available_normalizers()
            if not available:
                raise NormalizationError(
                    "quality-mode PPTX requires local PDF normalization, but neither "
                    "Microsoft PowerPoint automation nor LibreOffice headless export is available. "
                    "Install or repair one of those applications; direct PPTX quality extraction "
                    "was not attempted."
                )
            normalization_backend = available[0]
            profile = self._quality_pptx_profile(profile, normalization_backend)
        if not self.chunker_script.is_file():
            raise PipelineError(f"MinerU chunking module is missing: {self.chunker_script}")
        hash_started = time.monotonic()
        source_sha = sha256_file(source_path)
        hash_duration = time.monotonic() - hash_started
        extraction_profile_value = extraction_profile(profile)
        extraction_fingerprint = extraction_profile_fingerprint(profile)
        retrieval_profile_value = retrieval_profile(
            profile, chunker_fingerprint=sha256_file(self.chunker_script)
        )
        retrieval_fingerprint = retrieval_profile_fingerprint(retrieval_profile_value)
        cache_key = extraction_cache_key(source_sha, extraction_fingerprint)
        cache_dir = self.state_dir / "cache" / cache_key
        selected_course = safe_component(
            course or infer_course(source_path, config["default_course_fallback"]),
            fallback=config["default_course_fallback"],
        )
        selected_title = safe_component(title or source_title(source_path))
        catalog_started = time.monotonic()
        catalog = self.catalog()
        document_id = resolve_document_id(
            catalog,
            source_path=source_path,
            source_sha=source_sha,
            course=selected_course,
            title=selected_title,
        )
        existing_document = catalog["documents"].get(document_id)
        if isinstance(existing_document, dict):
            if course is None and existing_document.get("course"):
                selected_course = safe_component(
                    str(existing_document["course"]),
                    fallback=config["default_course_fallback"],
                )
            if title is None and existing_document.get("title"):
                selected_title = safe_component(str(existing_document["title"]))
        revision_id = revision_id_for(cache_key)
        cache_lookup_duration = time.monotonic() - catalog_started

        vault_value = str(vault) if vault is not None else config.get("vault_path")
        effective_vault = Path(vault_value).expanduser().resolve(strict=False) if vault_value else None
        if effective_vault is not None and not effective_vault.is_dir():
            raise PipelineError(f"configured vault does not exist: {effective_vault}")
        publish_requested = bool(
            not no_publish and config.get("publish_enabled", True) and effective_vault is not None
        )
        dry_publication = None
        if effective_vault is not None and not no_publish:
            directory_name = publication_directory_name(
                catalog,
                document_id=document_id,
                course=selected_course,
                title=selected_title,
                vault=effective_vault,
                publication_root=publication_root,
            )
            dry_publication = str(
                effective_vault / publication_root / selected_course / directory_name
            )

        cache_valid = False
        cache_manifest: Optional[dict[str, Any]] = None
        cache_validation: Optional[ArtifactValidation] = None
        cache_warning: Optional[str] = None
        if cache_dir.is_dir():
            try:
                cache_manifest, cache_validation = validate_cache(
                    cache_dir, source_sha, extraction_fingerprint
                )
                cache_valid = True
            except ValidationError as exc:
                cache_warning = f"existing cache ignored because validation failed: {exc}"
        legacy_cache = None
        if not cache_valid and not refresh:
            legacy_cache = find_compatible_v1_cache(
                self.state_dir / "cache",
                source_sha=source_sha,
                desired_extraction_profile=extraction_profile_value,
            )
        extraction_available = cache_valid or legacy_cache is not None
        retrieval_target = cache_dir / "retrieval" / retrieval_fingerprint
        retrieval_valid = False
        if cache_valid and retrieval_target.is_dir():
            try:
                validate_retrieval_artifact(
                    retrieval_target,
                    cache_dir=cache_dir,
                    cache_key=cache_key,
                    source_sha=source_sha,
                    profile=retrieval_profile_value,
                    fingerprint=retrieval_fingerprint,
                )
                retrieval_valid = True
            except ValidationError:
                retrieval_valid = False

        if dry_run:
            return {
                "status": "DRY_RUN",
                "source": str(source_path),
                "course": selected_course,
                "title": selected_title,
                "mode": chosen_mode,
                "profile_fingerprint": extraction_fingerprint,
                "extraction_profile_fingerprint": extraction_fingerprint,
                "retrieval_profile_fingerprint": retrieval_fingerprint,
                "cache_key": cache_key,
                "cache": "hit" if extraction_available and not refresh else "miss",
                "retrieval_cache": "hit" if retrieval_valid and not refresh else "miss",
                "refresh": refresh,
                "would_run_mineru": refresh or not extraction_available,
                "would_normalize_pptx": bool(
                    normalization_backend and (refresh or not extraction_available)
                ),
                "normalizer_backend": (
                    normalization_backend.name if normalization_backend else None
                ),
                "would_generate_chunks": refresh or not retrieval_valid,
                "legacy_v1_cache": str(legacy_cache) if legacy_cache else None,
                "publish_destination": dry_publication,
                "would_publish": publish_requested,
                "publication_root": publication_root,
                "warnings": [cache_warning] if cache_warning else [],
            }

        retention = source_retention or config.get("source_retention", "reference")
        if retention not in {"reference", "copy"}:
            raise PipelineError("source retention must be reference or copy")
        metrics = {
            "hash_duration": round(hash_duration, 6),
            "cache_lookup_duration": round(cache_lookup_duration, 6),
            "normalization_duration": 0.0,
            "mineru_duration": 0.0,
            "retrieval_duration": 0.0,
            "validation_duration": 0.0,
            "publish_duration": 0.0,
            "total_duration": 0.0,
        }
        warnings = [cache_warning] if cache_warning else []
        staging_dir: Optional[Path] = None
        cache_hit = extraction_available and not refresh
        normalization_result: Optional[NormalizationResult] = None

        try:
            if legacy_cache is not None and not cache_valid and not refresh:
                staging_dir = self.state_dir / "staging" / (
                    f"{dt.datetime.now().strftime('%Y%m%dT%H%M%S')}-{cache_key[:12]}-{uuid.uuid4().hex[:8]}"
                )
                staging_dir.mkdir(parents=True, exist_ok=False)
                normalized = staging_dir / "normalized"
                cache_manifest, cache_validation = prepare_v1_migration(
                    legacy_cache,
                    normalized,
                    cache_key=cache_key,
                    extraction_profile_value=extraction_profile_value,
                    extraction_fingerprint=extraction_fingerprint,
                )
                commit_cache(normalized, cache_dir)
                cache_manifest, cache_validation = validate_cache(
                    cache_dir, source_sha, extraction_fingerprint
                )
                warnings.extend(cache_manifest.get("warnings", []))
            elif not cache_hit:
                staging_dir = self.state_dir / "staging" / (
                    f"{dt.datetime.now().strftime('%Y%m%dT%H%M%S')}-{cache_key[:12]}-{uuid.uuid4().hex[:8]}"
                )
                staging_dir.mkdir(parents=True, exist_ok=False)
                mineru_source = source_path
                if normalization_backend is not None:
                    normalization_result = self._normalize_pptx(
                        source=source_path,
                        staging_dir=staging_dir,
                        backend=normalization_backend,
                        expected_source_sha=source_sha,
                    )
                    metrics["normalization_duration"] = round(
                        normalization_result.duration, 6
                    )
                    warnings.extend(normalization_result.warnings)
                    mineru_source = normalization_result.path
                extract_started = time.monotonic()
                outcome, normalized, validation, extraction_metadata = self._extract(
                    source=mineru_source,
                    staging_dir=staging_dir,
                    profile=profile,
                )
                metrics["mineru_duration"] = round(time.monotonic() - extract_started, 6)
                if sha256_file(source_path) != source_sha:
                    raise PipelineError("source changed during extraction; staged result was not committed")
                if normalization_result is not None:
                    derivative = normalized / "normalized" / "source.pdf"
                    derivative.parent.mkdir()
                    shutil.copy2(normalization_result.path, derivative)
                    if sha256_file(derivative) != normalization_result.sha256:
                        raise NormalizationError("normalized PDF changed while being retained")
                    extraction_metadata["normalization"] = {
                        "backend": normalization_result.backend,
                        "version": normalization_result.version,
                        "slide_count": normalization_result.slide_count,
                        "page_count": normalization_result.page_count,
                        "normalized_pdf_sha256": normalization_result.sha256,
                        "duration_seconds": round(normalization_result.duration, 6),
                        "warnings": normalization_result.warnings,
                    }
                retained_source = (
                    archive_source(self.state_dir, source_path, source_sha)
                    if retention == "copy"
                    else None
                )
                completed_at = utc_now()
                manifest = self._make_manifest(
                    source=source_path,
                    source_sha=source_sha,
                    source_stat=source_path.stat(),
                    document_id=document_id,
                    revision_id=revision_id,
                    cache_key=cache_key,
                    course=selected_course,
                    title=selected_title,
                    profile=profile,
                    fingerprint=extraction_fingerprint,
                    outcome=outcome,
                    normalized=normalized,
                    validation=validation,
                    extraction_metadata=extraction_metadata,
                    started_at=started_at,
                    completed_at=completed_at,
                    retained_source=retained_source,
                    normalization_result=normalization_result,
                )
                self._record_extraction(normalized, outcome=outcome, metadata=extraction_metadata)
                atomic_write_json(normalized / "manifest.json", manifest)
                validate_started = time.monotonic()
                validate_extraction_artifacts(normalized)
                validate_manifest(manifest, normalized)
                metrics["validation_duration"] = round(
                    time.monotonic() - validate_started, 6
                )
                commit_cache(normalized, cache_dir)
                cache_manifest, cache_validation = validate_cache(
                    cache_dir, source_sha, extraction_fingerprint
                )
                warnings.extend(manifest.get("warnings", []))
            else:
                assert cache_manifest is not None and cache_validation is not None
                warnings.extend(cache_manifest.get("warnings", []))

            retrieval_started = time.monotonic()
            chunks_path, retrieval_validation, retrieval_hit = ensure_retrieval_artifact(
                cache_dir,
                cache_key=cache_key,
                source_sha=source_sha,
                profile=retrieval_profile_value,
                fingerprint=retrieval_fingerprint,
                chunker_script=self.chunker_script,
            )
            metrics["retrieval_duration"] = round(time.monotonic() - retrieval_started, 6)

            publication: Optional[dict[str, Any]] = None
            publication_directory: Optional[str] = None
            if publish_requested:
                assert effective_vault is not None and cache_manifest is not None
                publication_directory = publication_directory_name(
                    catalog,
                    document_id=document_id,
                    course=selected_course,
                    title=selected_title,
                    vault=effective_vault,
                    publication_root=publication_root,
                )
                source_destination = (
                    effective_vault
                    / publication_root
                    / selected_course
                    / publication_directory
                )
                existing_publication = None
                existing_for_document = catalog.get("documents", {}).get(document_id)
                if isinstance(existing_for_document, dict) and isinstance(
                    existing_for_document.get("publication"), dict
                ):
                    existing_publication = existing_for_document["publication"]
                expected_existing_fingerprint = (
                    existing_publication.get("tree_fingerprint")
                    if isinstance(existing_publication, dict)
                    and normalized_path_key(existing_publication.get("source_path", ""))
                    == normalized_path_key(source_destination)
                    else None
                )
                publish_started = time.monotonic()
                publication = publish_revision(
                    cache_dir,
                    vault=effective_vault,
                    publication_root=publication_root,
                    course=selected_course,
                    document_directory=publication_directory,
                    expected_existing_fingerprint=expected_existing_fingerprint,
                    fault_injector=publication_fault_injector,
                )
                metrics["publish_duration"] = round(
                    time.monotonic() - publish_started, 6
                )
            elif not no_publish and effective_vault is None:
                warnings.append("vault_not_configured; cached without publication")

            # Re-read immediately before commit so sequential multi-file ingestion sees
            # every prior successful document.  Publication failures never reach here.
            catalog = self.catalog()
            update_catalog(
                catalog,
                document_id=document_id,
                revision_id=revision_id,
                source_path=source_path,
                source_sha=source_sha,
                course=selected_course,
                title=selected_title,
                cache_key=cache_key,
                extraction_fingerprint=extraction_fingerprint,
                retrieval_profile_value=retrieval_profile_value,
                retrieval_fingerprint=retrieval_fingerprint,
                chunks_path=chunks_path,
                cache_dir=cache_dir,
                publication=publication,
                publication_directory=publication_directory,
            )
            atomic_write_json(self.state_dir / "catalog.json", catalog)
            metrics["total_duration"] = round(time.monotonic() - total_started, 6)
            result = {
                "status": "CACHE_HIT" if cache_hit else "INGESTED",
                "source": str(source_path),
                "course": selected_course,
                "title": selected_title,
                "document_id": document_id,
                "revision_id": revision_id,
                "cache_key": cache_key,
                "cache_path": str(cache_dir),
                "profile_fingerprint": extraction_fingerprint,
                "extraction_profile_fingerprint": extraction_fingerprint,
                "retrieval_profile_fingerprint": retrieval_fingerprint,
                "chunks_path": str(chunks_path),
                "retrieval_cache": "hit" if retrieval_hit else "generated",
                "published": publication is not None,
                "published_path": publication.get("source_path") if publication else None,
                "publication_root": publication_root,
                "backend": cache_manifest.get("api_used") if cache_manifest else None,
                "normalizer_backend": (
                    cache_manifest.get("normalized_source", {}).get("backend")
                    if cache_manifest and isinstance(cache_manifest.get("normalized_source"), dict)
                    else None
                ),
                "normalizer_version": (
                    cache_manifest.get("normalized_source", {}).get("version")
                    if cache_manifest and isinstance(cache_manifest.get("normalized_source"), dict)
                    else None
                ),
                "normalized_pdf_sha256": (
                    cache_manifest.get("normalized_pdf_sha256") if cache_manifest else None
                ),
                "normalization_cache": (
                    "hit" if cache_hit and normalization_backend is not None
                    else "generated" if normalization_result is not None
                    else "not_applicable"
                ),
                "refresh": refresh,
                "warnings": list(dict.fromkeys(filter(None, warnings))),
                "metrics": metrics,
            }
            if self.logger:
                self.logger.log("ingest_success", **result)
            if staging_dir and staging_dir.exists():
                shutil.rmtree(staging_dir, ignore_errors=True)
            return result
        except Exception as exc:
            metrics["total_duration"] = round(time.monotonic() - total_started, 6)
            if staging_dir and staging_dir.exists():
                failure = {
                    "timestamp": utc_now(),
                    "source": str(source_path),
                    "document_id": document_id,
                    "cache_key": cache_key,
                    "error_type": type(exc).__name__,
                    "error": _redact_text(str(exc))[:2000],
                    "metrics": metrics,
                }
                try:
                    atomic_write_json(staging_dir / "failure.json", failure)
                except OSError:
                    pass
            if self.logger:
                self.logger.log(
                    "ingest_failure",
                    source=str(source_path),
                    document_id=document_id,
                    cache_key=cache_key,
                    error_type=type(exc).__name__,
                    error=str(exc),
                    metrics=metrics,
                    staging_path=str(staging_dir) if staging_dir else None,
                )
            raise

    def migration_plan(self, *, vault: Optional[Path | str] = None) -> dict[str, Any]:
        config = self.config()
        publication_root = configured_publication_root(config)
        vault_value = str(vault) if vault is not None else config.get("vault_path")
        if not vault_value:
            raise PipelineError("no Obsidian vault is configured for migration planning")
        effective_vault = Path(vault_value).expanduser().resolve(strict=False)
        if not effective_vault.is_dir():
            raise PipelineError(f"configured vault does not exist: {effective_vault}")
        return plan_legacy_publication_migration(
            vault=effective_vault,
            catalog=self.catalog(),
            publication_root=publication_root,
        )

    def doctor(self) -> dict[str, Any]:
        checks: dict[str, Any] = {}
        checks["state_directory"] = {
            "exists": self.state_dir.is_dir(),
            "path": str(self.state_dir),
        }
        checks["config"] = {"ok": False}
        try:
            config = self.config()
            checks["config"] = {"ok": True, "path": str(self.state_dir / "config.json")}
        except Exception as exc:
            config = default_config()
            checks["config"] = {"ok": False, "error": str(exc)}
        checks["catalog"] = {"ok": False}
        try:
            self.catalog()
            checks["catalog"] = {"ok": True, "path": str(self.state_dir / "catalog.json")}
        except Exception as exc:
            checks["catalog"] = {"ok": False, "error": str(exc)}
        checks["python"] = {
            "ok": self.python_executable.is_file(),
            "path": str(self.python_executable),
        }
        checks["mineru_skill"] = {
            "ok": self.mineru_script.is_file(),
            "path": str(self.mineru_script),
            "version": mineru_version_from_source(self.mineru_script),
            "fingerprint": mineru_fingerprint(self.mineru_script),
        }
        checks["mineru_chunker"] = {
            "ok": self.chunker_script.is_file(),
            "path": str(self.chunker_script),
            "fingerprint": sha256_file(self.chunker_script) if self.chunker_script.is_file() else None,
        }
        pypdf_ok = interpreter_has_module(self.python_executable, "pypdf")
        help_outcome = (
            run_captured(
                [str(self.python_executable), str(self.mineru_script), "--help"],
                timeout=30,
            )
            if self.python_executable.is_file() and self.mineru_script.is_file()
            else None
        )
        split_flag = bool(
            help_outcome
            and help_outcome.returncode == 0
            and "--split" in help_outcome.stdout
        )
        checks["pypdf"] = {
            "ok": pypdf_ok,
            "split_available": pypdf_ok and split_flag,
        }
        checks["mineru_cli"] = {
            "help_ok": bool(help_outcome and help_outcome.returncode == 0),
            "split_flag": split_flag,
        }
        checks["local_engine"] = {
            "pymupdf4llm": self.local_engine_available(),
            "required_for_private_mode": True,
        }
        normalizers = self.available_normalizers()
        checks["pptx_normalization"] = {
            "available": bool(normalizers),
            "selected": normalizers[0].name if normalizers else None,
            "backends": [dataclasses.asdict(backend) for backend in normalizers],
            "required_for_quality_pptx": True,
        }
        checks["token"] = {"detected": bool(os.environ.get("MINERU_TOKEN"))}

        mineru_doctor: dict[str, Any] = {"ok": False}
        if self.python_executable.is_file() and self.mineru_script.is_file():
            outcome = run_captured(
                [
                    str(self.python_executable),
                    str(self.mineru_script),
                    "--doctor",
                    "--json",
                ],
                timeout=90,
            )
            report = outcome.status or parse_json_status(outcome.stdout)
            mineru_doctor = {
                "ok": outcome.returncode == 0 and bool(report and report.get("healthy")),
                "returncode": outcome.returncode,
                "report": report,
                "error": _redact_text(outcome.stderr[-1000:]) if outcome.stderr else None,
            }
            if isinstance(report, dict):
                token_report = report.get("token", {})
                checks["token"]["accepted"] = bool(
                    isinstance(token_report, dict) and token_report.get("ok")
                )
                checks["standard_api"] = {
                    "available": checks["token"]["detected"] and checks["token"].get("accepted", False)
                }
        checks["mineru_doctor"] = mineru_doctor

        checks["state_writable"] = {"ok": _write_probe(self.state_dir)}
        vault_value = config.get("vault_path")
        if vault_value:
            vault = Path(vault_value).expanduser().resolve(strict=False)
            checks["vault"] = {
                "configured": True,
                "path": str(vault),
                "exists": vault.is_dir(),
                "writable": _write_probe(vault) if vault.is_dir() else False,
            }
        else:
            checks["vault"] = {"configured": False, "path": None}
        required_ok = [
            checks["state_directory"]["exists"],
            checks["config"]["ok"],
            checks["catalog"]["ok"],
            checks["python"]["ok"],
            checks["mineru_skill"]["ok"],
            checks["mineru_chunker"]["ok"],
            checks["pypdf"]["ok"],
            checks["state_writable"]["ok"],
            checks["mineru_doctor"]["ok"],
        ]
        return {"healthy": all(required_ok), "checks": checks}


def _write_probe(directory: Path) -> bool:
    if not directory.is_dir():
        return False
    try:
        with tempfile.NamedTemporaryFile(prefix=".study-docs-write-", dir=directory, delete=False) as handle:
            probe = Path(handle.name)
            handle.write(b"ok")
            handle.flush()
            os.fsync(handle.fileno())
        probe.unlink()
        return True
    except OSError:
        return False


def _is_obvious_temp_file(path: Path) -> bool:
    name = path.name.casefold()
    return name.startswith(("~$", ".~", ".")) or name.endswith((".tmp", ".part", ".partial"))


def expand_inputs(inputs: Iterable[str], *, recursive: bool) -> list[Path]:
    expanded: list[Path] = []
    seen: set[str] = set()
    for raw in inputs:
        path = Path(raw).expanduser().resolve(strict=False)
        candidates: Iterable[Path]
        if path.is_dir():
            if recursive:
                gathered: list[Path] = []
                for current, directories, files in os.walk(path, followlinks=False):
                    directories[:] = [
                        name
                        for name in directories
                        if name.casefold() not in SKIP_DIRECTORY_NAMES and not name.startswith(".")
                    ]
                    gathered.extend(Path(current) / name for name in files)
                candidates = sorted(gathered)
            else:
                candidates = sorted(child for child in path.iterdir() if child.is_file())
        else:
            candidates = [path]
        for candidate in candidates:
            if (
                candidate.is_file()
                and candidate.suffix.casefold() in SUPPORTED_SUFFIXES
                and not _is_obvious_temp_file(candidate)
            ):
                key = normalized_path_key(candidate)
                if key not in seen:
                    seen.add(key)
                    expanded.append(candidate)
    return expanded


def human_result(result: dict[str, Any]) -> str:
    status = result.get("status", "UNKNOWN")
    source = result.get("source", "")
    course = result.get("course", "")
    duration = result.get("metrics", {}).get("total_duration")
    published = " PUBLISHED" if result.get("published") else ""
    warning_count = len(result.get("warnings", []))
    suffix = f" warnings={warning_count}" if warning_count else ""
    return f"{status}{published} | {source} | course={course} | duration={duration}s{suffix}"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="study-document-ingest",
        description="Persistently ingest course documents through MinerU into a validated content-addressed cache.",
    )
    parser.add_argument("inputs", nargs="*", help="File(s) or directories to ingest")
    parser.add_argument("--course", help="Explicit course name")
    parser.add_argument("--title", help="Explicit title (single input only)")
    parser.add_argument("--mode", choices=["quality", "fast", "private"])
    parser.add_argument("--lang", default="ch", help="MinerU language code (default: ch)")
    parser.add_argument("--chunk-size", type=int, default=DEFAULT_CHUNK_SIZE)
    parser.add_argument("--vault", help="Existing Obsidian vault for this run; does not change config")
    parser.add_argument("--refresh", action="store_true", help="Force MinerU even when compatible cache exists")
    parser.add_argument("--no-publish", action="store_true", help="Cache/catalog only")
    parser.add_argument("--dry-run", action="store_true", help="Plan without MinerU or writes")
    parser.add_argument("--recursive", action="store_true", help="Recurse into input directories")
    parser.add_argument("--source-retention", choices=["reference", "copy"])
    parser.add_argument("--doctor", action="store_true", help="Check pipeline and MinerU environment")
    parser.add_argument(
        "--migration-plan",
        action="store_true",
        help="Inspect legacy vault publications and print a read-only migration plan",
    )
    parser.add_argument("--json", action="store_true", help="Emit concise machine-readable status")
    parser.add_argument(
        "--state-dir",
        default=os.environ.get("STUDY_DOCS_STATE_DIR", str(DEFAULT_STATE_DIR)),
        help=argparse.SUPPRESS,
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    pipeline = IngestionPipeline(
        state_dir=args.state_dir,
        initialize=not (args.dry_run or args.migration_plan),
    )
    if args.migration_plan:
        try:
            report = pipeline.migration_plan(vault=args.vault)
        except Exception as exc:
            report = {
                "status": "FAILED",
                "error": str(exc),
                "operation_performed": False,
            }
            code = 1
        else:
            code = 0
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        return code
    if args.doctor:
        report = pipeline.doctor()
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print("HEALTHY" if report["healthy"] else "UNHEALTHY")
            for name, detail in report["checks"].items():
                print(f"  {name}: {json.dumps(detail, ensure_ascii=False)}")
        return 0 if report["healthy"] else 1
    if not args.inputs:
        parser.error(
            "at least one input is required unless --doctor or --migration-plan is used"
        )
    sources = expand_inputs(args.inputs, recursive=args.recursive)
    if not sources:
        parser.error("no supported files found")
    if args.title and len(sources) != 1:
        parser.error("--title is only valid when exactly one source is selected")
    results: list[dict[str, Any]] = []
    failed = 0
    for source in sources:
        try:
            result = pipeline.ingest_file(
                source,
                course=args.course,
                title=args.title,
                mode=args.mode,
                language=args.lang,
                chunk_size=args.chunk_size,
                vault=args.vault,
                refresh=args.refresh,
                no_publish=args.no_publish,
                dry_run=args.dry_run,
                source_retention=args.source_retention,
            )
        except Exception as exc:
            failed += 1
            result = {
                "status": "FAILED",
                "source": str(source),
                "error_type": type(exc).__name__,
                "error": _redact_text(str(exc)),
            }
        results.append(result)
        if not args.json:
            if result["status"] == "FAILED":
                print(f"FAILED | {result['source']} | {result['error']}")
            else:
                print(human_result(result))
    if args.json:
        print(
            json.dumps(
                {
                    "total": len(results),
                    "succeeded": len(results) - failed,
                    "failed": failed,
                    "results": results,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
