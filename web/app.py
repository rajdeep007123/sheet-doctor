#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import io
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import pandas as pd
import requests
import streamlit as st


ROOT = Path(__file__).resolve().parent.parent
CSV_DIAGNOSE = ROOT / "skills" / "csv-doctor" / "scripts" / "diagnose.py"
CSV_HEAL = ROOT / "skills" / "csv-doctor" / "scripts" / "heal.py"
EXCEL_DIAGNOSE = ROOT / "skills" / "excel-doctor" / "scripts" / "diagnose.py"
EXCEL_HEAL = ROOT / "skills" / "excel-doctor" / "scripts" / "heal.py"
LOADER_PATH = ROOT / "skills" / "csv-doctor" / "scripts" / "loader.py"
HEAL_PATH = ROOT / "skills" / "csv-doctor" / "scripts" / "heal.py"

TEXTUAL_EXTS = {".csv", ".tsv", ".txt", ".json", ".jsonl"}
MODERN_EXCEL_EXTS = {".xlsx", ".xlsm"}
LEGACY_PREVIEW_EXTS = {".xls", ".ods"}
WORKBOOK_EXTS = MODERN_EXCEL_EXTS | LEGACY_PREVIEW_EXTS
SUPPORTED_EXTS = TEXTUAL_EXTS | WORKBOOK_EXTS
MAX_REMOTE_FILE_MB = 100
MAX_REMOTE_FILE_BYTES = MAX_REMOTE_FILE_MB * 1024 * 1024
SEMANTIC_ROLE_OPTIONS = ["auto", "ignore", "identifier", "name", "date", "amount", "measurement", "currency", "status", "department", "category", "notes"]


@st.cache_resource(show_spinner=False)
def load_loader_module():
    spec = importlib.util.spec_from_file_location("sheet_doctor_loader", LOADER_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


LOADER = load_loader_module()


@st.cache_resource(show_spinner=False)
def load_heal_module():
    spec = importlib.util.spec_from_file_location("sheet_doctor_heal_preview", HEAL_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules["sheet_doctor_heal_preview"] = module
    spec.loader.exec_module(module)
    return module


HEAL = load_heal_module()


def ensure_state() -> None:
    st.session_state.setdefault("processing", False)
    st.session_state.setdefault("job", [])
    st.session_state.setdefault("results", [])
    st.session_state.setdefault(
        "prompt_input",
        "Make this readable for humans, keep what can be saved, and clearly show what still needs manual review.",
    )
    st.session_state.setdefault("public_urls_input", "")


def infer_intent(prompt: str) -> str:
    text = prompt.lower().strip()
    heal_markers = [
        "fix",
        "clean",
        "heal",
        "repair",
        "readable",
        "understandable",
        "make sense",
        "human",
        "organize",
        "organise",
    ]
    diagnose_markers = [
        "diagnose",
        "check",
        "analyze",
        "analyse",
        "inspect",
        "what is wrong",
        "what's wrong",
        "why",
    ]
    heal_score = sum(marker in text for marker in heal_markers)
    diagnose_score = sum(marker in text for marker in diagnose_markers)
    return "Make Readable" if heal_score >= diagnose_score else "Diagnose Only"


def run_script(script: Path, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(script), *args],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )


def normalize_public_url(raw_url: str) -> str:
    parsed = urlparse(raw_url.strip())
    if not parsed.scheme:
        raise ValueError("URL must start with http:// or https://")

    host = parsed.netloc.lower()
    path = parsed.path
    query = parse_qs(parsed.query, keep_blank_values=True)

    if host == "github.com" and "/blob/" in path:
        owner_repo, blob_path = path.lstrip("/").split("/blob/", 1)
        owner, repo = owner_repo.split("/", 1)
        branch, file_path = blob_path.split("/", 1)
        return f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{file_path}"

    if "dropbox.com" in host:
        query["dl"] = ["1"]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    if "box.com" in host:
        query["download"] = ["1"]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    if host in {"drive.google.com", "docs.google.com"}:
        sheet_match = re.search(r"/spreadsheets/d/([^/]+)", path)
        if sheet_match:
            gid = query.get("gid", ["0"])[0]
            return (
                f"https://docs.google.com/spreadsheets/d/{sheet_match.group(1)}/export"
                f"?format=xlsx&gid={gid}"
            )
        match = re.search(r"/file/d/([^/]+)", path)
        if match:
            return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
        if "id" in query:
            return f"https://drive.google.com/uc?export=download&id={query['id'][0]}"

    if host.endswith("1drv.ms") or "onedrive.live.com" in host:
        query["download"] = ["1"]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    return raw_url.strip()


def parse_public_urls(raw_urls: str) -> list[str]:
    return [line.strip() for line in raw_urls.splitlines() if line.strip()]


def remote_filename(raw_url: str, response: requests.Response) -> str:
    content_disposition = response.headers.get("content-disposition", "")
    match = re.search(r'filename\\*=UTF-8\'\'([^;]+)|filename="([^"]+)"|filename=([^;]+)', content_disposition, re.I)
    if match:
        for group in match.groups():
            if group:
                return Path(group.strip().strip('"')).name
    redirected = response.url or raw_url
    return Path(urlparse(redirected).path).name or Path(urlparse(raw_url).path).name or "downloaded_file"


def infer_extension_from_response(
    raw_url: str,
    response: requests.Response,
    filename: str,
    content: bytes,
) -> str:
    ext = Path(filename).suffix.lower()
    if ext in SUPPORTED_EXTS:
        return ext

    content_type = response.headers.get("content-type", "").split(";")[0].strip().lower()
    content_type_map = {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsm",
        "application/vnd.ms-excel": ".xls",
        "application/vnd.oasis.opendocument.spreadsheet": ".ods",
        "text/csv": ".csv",
        "text/tab-separated-values": ".tsv",
        "application/json": ".json",
        "application/x-ndjson": ".jsonl",
        "application/jsonl": ".jsonl",
        "application/jsonlines": ".jsonl",
    }
    if content_type in content_type_map:
        return content_type_map[content_type]

    parsed = urlparse(raw_url)
    if "docs.google.com" in parsed.netloc.lower() and "/spreadsheets/" in parsed.path:
        return ".xlsx"

    if content.startswith(b"PK"):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                names = set(zf.namelist())
                if "mimetype" in names:
                    try:
                        mimetype = zf.read("mimetype").decode("utf-8", errors="ignore").strip()
                    except Exception:
                        mimetype = ""
                    if mimetype == "application/vnd.oasis.opendocument.spreadsheet":
                        return ".ods"
                if "xl/vbaProject.bin" in names:
                    return ".xlsm"
                if "xl/workbook.xml" in names:
                    return ".xlsx"
        except zipfile.BadZipFile:
            pass

    if content[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return ".xls"

    try:
        sample = content[:8192].decode("utf-8", errors="replace")
    except Exception:
        sample = ""

    lines = [line.strip() for line in sample.splitlines() if line.strip()]
    if not lines:
        return ext

    if lines[0].startswith("{") or lines[0].startswith("["):
        if len(lines) > 1 and all(line.startswith("{") for line in lines[:5]):
            return ".jsonl"
        return ".json"
    if "\t" in "\n".join(lines[:5]):
        return ".tsv"
    if any(delim in "\n".join(lines[:5]) for delim in [",", ";", "|"]):
        return ".csv"
    return ".txt"


def fetch_remote_source(raw_url: str, folder: Path) -> dict:
    url = normalize_public_url(raw_url)
    response = requests.get(url, timeout=60, allow_redirects=True, stream=True)
    try:
        response.raise_for_status()
        content_length = response.headers.get("Content-Length")
        if content_length:
            try:
                declared_size = int(content_length)
            except ValueError:
                declared_size = None
            if declared_size and declared_size > MAX_REMOTE_FILE_BYTES:
                raise ValueError(f"Remote file is larger than {MAX_REMOTE_FILE_MB} MB.")

        chunks: list[bytes] = []
        downloaded = 0
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if not chunk:
                continue
            downloaded += len(chunk)
            if downloaded > MAX_REMOTE_FILE_BYTES:
                raise ValueError(f"Remote file is larger than {MAX_REMOTE_FILE_MB} MB.")
            chunks.append(chunk)
        content = b"".join(chunks)
    finally:
        response.close()

    filename = remote_filename(raw_url, response)
    ext = infer_extension_from_response(raw_url, response, filename, content)
    if ext not in SUPPORTED_EXTS:
        raise ValueError(f"Unsupported remote file type: {ext or '[missing extension]'}")

    if not Path(filename).suffix:
        filename = f"{filename}{ext}"

    target = folder / filename
    target.write_bytes(content)
    return {
        "name": filename,
        "ext": ext,
        "bytes": content,
        "path": target,
    }


def workbook_sheet_info(path: Path) -> tuple[list[str], bool, Optional[str]]:
    try:
        if path.suffix.lower() == ".ods":
            with pd.ExcelFile(path, engine="odf") as xf:
                sheet_names = list(xf.sheet_names)
        else:
            with pd.ExcelFile(path) as xf:
                sheet_names = list(xf.sheet_names)
    except Exception as exc:
        return [], False, str(exc)

    try:
        same_columns = LOADER._sheets_same_columns(
            path,
            sheet_names,
            engine="odf" if path.suffix.lower() == ".ods" else None,
        )
    except Exception:
        same_columns = False

    return sheet_names, same_columns, None


def load_preview(path: Path, sheet_name: Optional[str] = None, consolidate: Optional[bool] = None) -> dict:
    return LOADER.load_file(path, sheet_name=sheet_name, consolidate_sheets=consolidate)


def workbook_semantic_info(
    path: Path,
    sheet_name: Optional[str] = None,
    consolidate: Optional[bool] = None,
    header_row_override: Optional[int] = None,
    role_overrides: Optional[dict[int, str]] = None,
) -> tuple[Optional[dict], Optional[str]]:
    try:
        return (
            HEAL.inspect_healing_plan(
                path,
                sheet_name=sheet_name,
                consolidate_sheets=consolidate,
                header_row_override=header_row_override,
                role_overrides=role_overrides,
            ),
            None,
        )
    except Exception as exc:
        return None, str(exc)


def create_readable_export(
    source_path: Path,
    output_path: Path,
    prompt: str,
    sheet_name: Optional[str] = None,
    consolidate: Optional[bool] = None,
) -> dict:
    loaded = load_preview(source_path, sheet_name=sheet_name, consolidate=consolidate)
    df = loaded["dataframe"].copy()

    clean_columns = []
    for idx, column in enumerate(df.columns, start=1):
        text = str(column).strip()
        clean_columns.append(text or f"column_{idx}")
    df.columns = clean_columns

    for column in df.columns:
        df[column] = df[column].map(
            lambda value: "" if pd.isna(value) else str(value).replace("\r\n", " ").replace("\n", " ").strip()
        )

    notes = pd.DataFrame(
        [
            {"field": "source_file", "value": source_path.name},
            {"field": "detected_format", "value": loaded["detected_format"]},
            {"field": "prompt", "value": prompt.strip() or "Make this readable for humans"},
            {"field": "sheet_name", "value": loaded.get("sheet_name") or ""},
            {"field": "sheet_names", "value": ", ".join(loaded.get("sheet_names") or [])},
            {"field": "delimiter", "value": loaded.get("delimiter") or ""},
            {"field": "warnings", "value": " | ".join(loaded.get("warnings") or [])},
        ]
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Readable Data")
        notes.to_excel(writer, index=False, sheet_name="Load Notes")

    return loaded


def preview_payload(loaded: dict, workbook_semantics: Optional[dict] = None) -> dict:
    df = loaded["dataframe"]
    head = df.head(50).fillna("").astype(str)
    return {
        "rows": int(df.shape[0]),
        "columns": int(df.shape[1]),
        "detected_format": loaded.get("detected_format"),
        "sheet_name": loaded.get("sheet_name"),
        "sheet_names": loaded.get("sheet_names"),
        "warnings": loaded.get("warnings") or [],
        "head_records": head.to_dict(orient="records"),
        "workbook_semantics": workbook_semantics,
    }


def workbook_mode_details(ext: str, tabular_rescue: bool = False) -> Optional[dict]:
    if ext in MODERN_EXCEL_EXTS:
        if tabular_rescue:
            return {
                "mode": "tabular-rescue",
                "why": "Use this when the workbook is really a messy table and you want a flattened 3-sheet rescue output.",
                "tradeoff": "Workbook structure, formulas, and sheet-native layout are flattened into rows/columns.",
            }
        return {
            "mode": "workbook-native",
            "why": "Use this when preserving workbook sheets and workbook structure matters more than flattening the data.",
            "tradeoff": "The output stays a workbook and only safe sheet-level cleanup is applied; it does not become a Clean Data / Quarantine table rescue.",
        }
    if ext in LEGACY_PREVIEW_EXTS:
        return {
            "mode": "tabular-rescue-fallback",
            "why": "Legacy workbook support is routed through csv-doctor or a readable export fallback because excel-doctor does not support workbook-native .xls/.ods repair.",
            "tradeoff": "Workbook-native structure is not preserved.",
        }
    return None


def heal_support_message(ext: str, tabular_rescue: bool = False) -> tuple[bool, str]:
    if ext in TEXTUAL_EXTS:
        return True, "CSV doctor will heal this into a 3-sheet workbook."
    if ext in MODERN_EXCEL_EXTS:
        details = workbook_mode_details(ext, tabular_rescue=tabular_rescue)
        if details and details["mode"] == "tabular-rescue":
            return True, "CSV doctor tabular rescue will flatten this workbook into a 3-sheet readable workbook."
        return True, "Excel doctor will preserve the workbook structure and add a Change Log."
    if ext in LEGACY_PREVIEW_EXTS:
        return True, "Legacy workbook support uses tabular rescue or a readable export fallback rather than workbook-native healing."
    return False, "This format is not supported."


def inspect_local_bytes(file_bytes: bytes, suffix: str) -> tuple[list[str], bool, Optional[str]]:
    with tempfile.TemporaryDirectory(prefix="sheet_doctor_local_inspect_") as tmpdir:
        tmp_path = Path(tmpdir) / f"inspect{suffix}"
        tmp_path.write_bytes(file_bytes)
        return workbook_sheet_info(tmp_path)


def inspect_local_workbook_semantics(
    file_bytes: bytes,
    suffix: str,
    sheet_name: Optional[str] = None,
    consolidate: Optional[bool] = None,
    header_row_override: Optional[int] = None,
    role_overrides: Optional[dict[int, str]] = None,
) -> tuple[Optional[dict], Optional[str]]:
    with tempfile.TemporaryDirectory(prefix="sheet_doctor_local_semantic_") as tmpdir:
        tmp_path = Path(tmpdir) / f"semantic{suffix}"
        tmp_path.write_bytes(file_bytes)
        return workbook_semantic_info(
            tmp_path,
            sheet_name=sheet_name,
            consolidate=consolidate,
            header_row_override=header_row_override,
            role_overrides=role_overrides,
        )


def inspect_remote_url(url: str) -> tuple[list[str], bool, Optional[str]]:
    with tempfile.TemporaryDirectory(prefix="sheet_doctor_url_inspect_") as tmpdir:
        tmp_path = Path(tmpdir)
        try:
            remote = fetch_remote_source(url, tmp_path)
        except Exception as exc:
            return [], False, str(exc)
        return workbook_sheet_info(remote["path"])


def inspect_remote_workbook_semantics(
    url: str,
    sheet_name: Optional[str] = None,
    consolidate: Optional[bool] = None,
    header_row_override: Optional[int] = None,
    role_overrides: Optional[dict[int, str]] = None,
) -> tuple[Optional[dict], Optional[str]]:
    with tempfile.TemporaryDirectory(prefix="sheet_doctor_url_semantic_") as tmpdir:
        tmp_path = Path(tmpdir)
        try:
            remote = fetch_remote_source(url, tmp_path)
        except Exception as exc:
            return None, str(exc)
        return workbook_semantic_info(
            remote["path"],
            sheet_name=sheet_name,
            consolidate=consolidate,
            header_row_override=header_row_override,
            role_overrides=role_overrides,
        )


def source_label(item: dict) -> str:
    return f"{item['name']}  •  URL" if item["source_kind"] == "url" else item["name"]


def source_note(item: dict) -> Optional[str]:
    if item["source_kind"] != "url":
        return None

    parsed = urlparse(item["source_label"])
    host = parsed.netloc.lower()
    path = parsed.path

    if "docs.google.com" in host and "/spreadsheets/" in path:
        return "Detected Google Sheet. It will be exported to .xlsx before processing."
    if host == "github.com" and "/blob/" in path:
        return "Detected GitHub file page. The URL will be rewritten to the raw file before processing."
    if "dropbox.com" in host:
        return "Detected Dropbox share link. Direct download mode will be requested before processing."
    if "box.com" in host:
        return "Detected Box share link. Direct download mode will be requested before processing."
    if host.endswith("1drv.ms") or "onedrive.live.com" in host:
        return "Detected OneDrive share link. Download mode will be requested before processing."
    return None


def source_items(uploads, raw_urls: str) -> list[dict]:
    items = []
    for upload in uploads or []:
        items.append(
            {
                "name": upload.name,
                "ext": Path(upload.name).suffix.lower(),
                "bytes": upload.getvalue(),
                "source_kind": "upload",
                "source_label": upload.name,
            }
        )
    for idx, url in enumerate(parse_public_urls(raw_urls)):
        inferred_name = Path(urlparse(url).path).name or f"remote-file-{idx + 1}"
        items.append(
            {
                "name": inferred_name,
                "ext": Path(inferred_name).suffix.lower(),
                "bytes": None,
                "source_kind": "url",
                "source_label": url,
            }
        )
    return items


def build_job(prompt: str, intent: str, sources: list[dict]) -> list[dict]:
    jobs = []
    for idx, source in enumerate(sources):
        source_id = f"{idx}_{source['source_label']}"
        role_overrides = {}
        prefix = f"role_{source_id}_"
        for key, value in st.session_state.items():
            if not key.startswith(prefix) or value in (None, "", "auto"):
                continue
            try:
                column_index = int(key.rsplit("_", 1)[1])
            except ValueError:
                continue
            role_overrides[column_index - 1] = value

        header_row_override = st.session_state.get(f"headerrow_{source_id}")
        detected_header_row = st.session_state.get(f"detected_header_{source_id}")
        tabular_rescue = bool(st.session_state.get(f"tabular_{source_id}", source["ext"] in LEGACY_PREVIEW_EXTS))
        if source["ext"] in WORKBOOK_EXTS and (
            role_overrides
            or (
                header_row_override is not None
                and detected_header_row is not None
                and int(header_row_override) != int(detected_header_row)
            )
        ):
            tabular_rescue = True

        item = {
            "name": source["name"],
            "ext": source["ext"],
            "bytes": source["bytes"],
            "intent": intent,
            "prompt": prompt,
            "sheet_name": None,
            "consolidate": False,
            "source_kind": source["source_kind"],
            "source_label": source["source_label"],
            "header_row_override": header_row_override,
            "role_overrides": role_overrides,
            "tabular_rescue": tabular_rescue,
            "plan_confirmed": bool(st.session_state.get(f"confirmplan_{source_id}", False)),
        }
        if item["ext"] in WORKBOOK_EXTS:
            item["sheet_name"] = st.session_state.get(f"sheet_{source_id}")
            item["consolidate"] = bool(st.session_state.get(f"consolidate_{source_id}", False))
        jobs.append(item)
    return jobs


def pending_workbook_plan_confirmations(sources: list[dict], intent: str) -> list[str]:
    if intent != "Make Readable":
        return []

    pending: list[str] = []
    for idx, source in enumerate(sources):
        if source["ext"] not in WORKBOOK_EXTS:
            continue
        source_id = f"{idx}_{source['source_label']}"
        tabular_rescue = bool(st.session_state.get(f"tabular_{source_id}", source["ext"] in LEGACY_PREVIEW_EXTS))
        if not tabular_rescue:
            continue
        if not st.session_state.get(f"confirmplan_{source_id}", False):
            pending.append(source["name"])
    return pending


def process_one_item(item: dict, folder: Path) -> dict:
    if item["source_kind"] == "url":
        remote = fetch_remote_source(item["source_label"], folder)
        source_path = remote["path"]
        item["name"] = remote["name"]
        item["ext"] = remote["ext"]
    else:
        source_path = folder / item["name"]
        source_path.write_bytes(item["bytes"])

    ext = item["ext"]
    if ext in WORKBOOK_EXTS and not item.get("sheet_name") and not item.get("consolidate"):
        sheet_names, _, workbook_error = workbook_sheet_info(source_path)
        if workbook_error:
            raise ValueError(f"Could not inspect workbook sheets: {workbook_error}")
        if len(sheet_names) > 1:
            item["sheet_name"] = sheet_names[0]

    support_heal, support_message = heal_support_message(ext, tabular_rescue=bool(item.get("tabular_rescue")))
    mode_details = workbook_mode_details(ext, tabular_rescue=bool(item.get("tabular_rescue")))
    result = {
        "name": item["name"],
        "ext": ext,
        "intent": item["intent"],
        "support_message": support_message,
        "mode_details": mode_details,
        "status": "success",
        "preview": None,
        "report": None,
        "report_type": None,
        "stdout": "",
        "stderr": "",
        "download_name": None,
        "download_bytes": None,
        "messages": [],
        "workbook_semantics": None,
        "heal_summary": None,
    }

    workbook_semantics = None
    if ext in WORKBOOK_EXTS:
        workbook_semantics, semantic_error = workbook_semantic_info(
            source_path,
            sheet_name=item.get("sheet_name") if not item.get("consolidate") else None,
            consolidate=item.get("consolidate"),
            header_row_override=item.get("header_row_override"),
            role_overrides=item.get("role_overrides"),
        )
        if semantic_error:
            result["messages"].append(f"Workbook interpretation preview failed: {semantic_error}")
        else:
            result["workbook_semantics"] = workbook_semantics

    preview_loaded = load_preview(
        source_path,
        sheet_name=item.get("sheet_name") if not item.get("consolidate") else None,
        consolidate=item.get("consolidate"),
    )
    result["preview"] = preview_payload(preview_loaded, workbook_semantics=workbook_semantics)
    result["messages"].extend(preview_loaded.get("warnings") or [])

    if item["intent"] == "Diagnose Only":
        if ext in TEXTUAL_EXTS:
            completed = run_script(CSV_DIAGNOSE, str(source_path))
            result["stdout"] = completed.stdout.strip()
            result["stderr"] = completed.stderr.strip()
            if completed.returncode != 0:
                result["status"] = "error"
                result["messages"].append(result["stdout"] or result["stderr"] or "Diagnosis failed.")
                return result
            result["report"] = json.loads(completed.stdout)
            result["report_type"] = "csv"
            return result

        if ext in MODERN_EXCEL_EXTS:
            completed = run_script(EXCEL_DIAGNOSE, str(source_path))
            result["stdout"] = completed.stdout.strip()
            result["stderr"] = completed.stderr.strip()
            if completed.returncode != 0:
                result["status"] = "error"
                result["messages"].append(result["stdout"] or result["stderr"] or "Diagnosis failed.")
                return result
            result["report"] = json.loads(completed.stdout)
            result["report_type"] = "excel"
            return result

        result["status"] = "info"
        result["messages"].append("Structured diagnose is not wired for this format yet. Preview is available.")
        return result

    if not support_heal:
        result["status"] = "error"
        result["messages"].append("This format is not supported.")
        return result
    if ext in WORKBOOK_EXTS and item.get("tabular_rescue") and not item.get("plan_confirmed"):
        result["status"] = "error"
        result["messages"].append("Workbook rescue plan must be confirmed before healing runs.")
        return result

    if ext in MODERN_EXCEL_EXTS and not item.get("tabular_rescue"):
        output_name = f"{source_path.stem}_healed{source_path.suffix}"
        download_mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if ext == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        output_name = f"{source_path.stem}_readable.xlsx"
        download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    output_path = folder / output_name

    if ext in TEXTUAL_EXTS or (ext in WORKBOOK_EXTS and item.get("tabular_rescue")):
        heal_args = [str(source_path), str(output_path)]
        summary_path = folder / f"{source_path.stem}_heal_summary.json"
        heal_args.extend(["--json-summary", str(summary_path)])
        if ext in WORKBOOK_EXTS:
            if item.get("consolidate"):
                heal_args.append("--all-sheets")
            elif item.get("sheet_name"):
                heal_args.extend(["--sheet", str(item["sheet_name"])])
            if item.get("header_row_override"):
                heal_args.extend(["--header-row", str(item["header_row_override"])])
            for column_index, role in sorted((item.get("role_overrides") or {}).items()):
                heal_args.extend(["--role-override", f"{column_index + 1}={role}"])
            if item.get("plan_confirmed"):
                heal_args.append("--confirm-plan")
        completed = run_script(CSV_HEAL, *heal_args)
        result["stdout"] = completed.stdout.strip()
        result["stderr"] = completed.stderr.strip()
        if completed.returncode != 0:
            result["status"] = "error"
            result["messages"].append(result["stdout"] or result["stderr"] or "Healing failed.")
            return result
        if summary_path.exists():
            result["heal_summary"] = json.loads(summary_path.read_text(encoding="utf-8"))
    elif ext in MODERN_EXCEL_EXTS:
        summary_path = folder / f"{source_path.stem}_excel_heal_summary.json"
        completed = run_script(EXCEL_HEAL, str(source_path), str(output_path), "--json-summary", str(summary_path))
        result["stdout"] = completed.stdout.strip()
        result["stderr"] = completed.stderr.strip()
        if completed.returncode != 0:
            result["status"] = "error"
            result["messages"].append(result["stdout"] or result["stderr"] or "Healing failed.")
            return result
        if summary_path.exists():
            result["heal_summary"] = json.loads(summary_path.read_text(encoding="utf-8"))
    else:
        loaded = create_readable_export(
            source_path,
            output_path,
            prompt=item["prompt"],
            sheet_name=item.get("sheet_name") if not item.get("consolidate") else None,
            consolidate=item.get("consolidate"),
        )
        result["stdout"] = "Readable export created."
        result["messages"].extend(loaded.get("warnings") or [])

    result["download_name"] = output_name
    result["download_mime"] = download_mime
    result["download_bytes"] = output_path.read_bytes()
    return result


def process_job(items: list[dict]) -> list[dict]:
    results = []
    status_box = st.empty()
    progress_box = st.empty()
    progress = progress_box.progress(0.0, text="Preparing files...")

    with tempfile.TemporaryDirectory(prefix="sheet_doctor_ui_") as tmpdir:
        tmp_path = Path(tmpdir)
        total = len(items)
        for idx, item in enumerate(items, start=1):
            status_box.markdown(
                f"""
                <div class="work-status">
                    <div class="work-status__spinner"></div>
                    <div>
                        <div class="work-status__title">Working on {item['name']}</div>
                        <div class="work-status__text">File {idx} of {total}. Please be patient while the file is analyzed and prepared.</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            try:
                results.append(process_one_item(item, tmp_path))
            except Exception as exc:
                results.append(
                    {
                        "name": item["name"],
                        "ext": item["ext"],
                        "intent": item["intent"],
                        "support_message": "",
                        "status": "error",
                        "preview": None,
                        "report": None,
                        "report_type": None,
                        "stdout": "",
                        "stderr": "",
                        "download_name": None,
                        "download_bytes": None,
                        "messages": [str(exc)],
                    }
                )
            progress.progress(idx / total, text=f"Processed {idx} of {total} file(s)")

    status_box.empty()
    progress_box.empty()
    return results


def render_csv_report(report: dict) -> None:
    summary = report.get("summary", {})
    st.markdown(
        f"**Verdict:** `{summary.get('verdict', 'UNKNOWN')}`  \n"
        f"**Issues found:** `{summary.get('issue_count', 0)}`"
    )
    st.json(report)


def render_excel_report(report: dict) -> None:
    summary = report.get("summary", {})
    st.markdown(
        f"**Verdict:** `{summary.get('verdict', 'UNKNOWN')}`  \n"
        f"**Issues found:** `{summary.get('issue_count', 0)}`  \n"
        f"**Categories triggered:** `{summary.get('issue_categories_triggered', 0)}`"
    )
    warnings = report.get("manual_review_warnings") or []
    if warnings:
        st.warning("Workbook-native manual review:\n- " + "\n- ".join(warnings))
    st.json(report)


def render_preview(preview: dict) -> None:
    left, right = st.columns(2)
    left.metric("Rows", preview["rows"])
    left.metric("Columns", preview["columns"])
    right.metric("Format", preview.get("detected_format") or "-")
    right.metric("Sheet", preview.get("sheet_name") or "-")
    if preview.get("warnings"):
        st.warning(" | ".join(preview["warnings"]))
    if preview.get("head_records"):
        st.dataframe(pd.DataFrame(preview["head_records"]), width="stretch")
    else:
        st.info("No table rows were loaded for preview.")
    if preview.get("workbook_semantics"):
        render_workbook_semantics(preview["workbook_semantics"])


def render_workbook_semantics(inspection: dict) -> None:
    st.markdown("**Workbook interpretation**")
    cols = st.columns(4)
    cols[0].metric("Mode", inspection.get("healing_mode_candidate", "-"))
    cols[1].metric("Metadata rows", inspection.get("metadata_rows_removed", 0))
    band_rows = inspection.get("detected_header_band_rows") or []
    band_label = ", ".join(str(row) for row in band_rows) if band_rows else "-"
    cols[2].metric("Header band rows", band_label)
    cols[3].metric("Header merged", "Yes" if inspection.get("header_band_merged") else "No")

    headers = inspection.get("effective_headers") or []
    if headers:
        st.caption("Effective headers after workbook interpretation")
        st.code(" | ".join(headers))

    suggested_columns = inspection.get("suggested_semantic_columns") or []
    semantic_columns = inspection.get("semantic_columns") or []
    semantic_comparison = inspection.get("semantic_comparison") or []
    if semantic_columns:
        st.caption("Chosen semantic columns")
        st.dataframe(pd.DataFrame(semantic_columns), width="stretch", hide_index=True)
    else:
        st.info("No confident semantic column mapping was detected for this workbook preview.")
    if suggested_columns and semantic_comparison:
        st.caption("Detected mapping vs your overrides")
        comparison_df = pd.DataFrame(semantic_comparison).rename(
            columns={
                "column_index": "#",
                "header": "Header",
                "detected_role": "Detected role",
                "detected_confidence": "Detected confidence",
                "override_role": "Override",
                "final_role": "Final role",
                "final_confidence": "Final confidence",
            }
        )
        st.dataframe(comparison_df, width="stretch", hide_index=True)
    if inspection.get("applied_role_overrides"):
        st.caption(f"Applied overrides: {inspection['applied_role_overrides']}")


def render_mode_details(mode_details: Optional[dict]) -> None:
    if not mode_details:
        return
    st.markdown(
        f"**Mode:** `{mode_details['mode']}`  \n"
        f"**Why:** {mode_details['why']}  \n"
        f"**Tradeoff:** {mode_details['tradeoff']}"
    )


def render_results() -> None:
    results = st.session_state.get("results") or []
    if not results:
        return

    st.subheader("Results")
    metrics = st.columns(3)
    metrics[0].metric("Files processed", len(results))
    metrics[1].metric("Downloads ready", sum(1 for item in results if item.get("download_bytes")))
    metrics[2].metric("Needs review", sum(1 for item in results if item["status"] != "success"))

    for item in results:
        with st.expander(f"{item['name']}  •  {item['status'].upper()}", expanded=item["status"] != "success"):
            st.caption(item["support_message"])
            render_mode_details(item.get("mode_details"))
            for message in item.get("messages", []):
                if item["status"] == "error":
                    st.error(message)
                elif item["status"] == "info":
                    st.info(message)
                else:
                    st.warning(message)
            if item.get("preview"):
                render_preview(item["preview"])
            if item.get("report"):
                if item["report_type"] == "csv":
                    render_csv_report(item["report"])
                elif item["report_type"] == "excel":
                    render_excel_report(item["report"])
            if item.get("stdout"):
                st.code(item["stdout"])
            if item.get("heal_summary"):
                workbook_plan = item["heal_summary"].get("workbook_plan") or {}
                if workbook_plan:
                    st.caption("Structured healing summary captured")
                    st.json(workbook_plan)
                heal_warnings = item["heal_summary"].get("warnings") or []
                if heal_warnings:
                    st.warning("Workbook-native healing warnings:\n- " + "\n- ".join(heal_warnings))
            if item.get("download_bytes") and item.get("download_name"):
                st.download_button(
                    "Download fixed file",
                    data=item["download_bytes"],
                    file_name=item["download_name"],
                    mime=item.get("download_mime") or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key=f"download_{item['name']}_{item['status']}",
                )


def set_visuals() -> None:
    st.set_page_config(page_title="sheet-doctor UI", page_icon="🩺", layout="wide", initial_sidebar_state="collapsed")
    st.markdown(
        """
        <style>
        :root {
            --qt-bg: #ffffff;
            --qt-bg-secondary: #fafafe;
            --qt-surface: rgba(255, 255, 255, 0.96);
            --qt-surface-strong: rgba(250, 250, 254, 1);
            --qt-text: #2a2a32;
            --qt-text-secondary: #5a5a70;
            --qt-text-muted: #8b8ba3;
            --qt-border: #e8e8f0;
            --qt-border-strong: #d6d6e1;
            --qt-primary: #9d72ff;
            --qt-primary-strong: #8b4cf7;
            --qt-primary-deep: #7c3aed;
            --qt-shadow: 0 20px 25px -5px rgba(99, 102, 241, 0.10), 0 10px 10px -5px rgba(99, 102, 241, 0.04);
        }
        @media (prefers-color-scheme: dark) {
            :root {
                --qt-bg: #1a1a1a;
                --qt-bg-secondary: #22222b;
                --qt-surface: rgba(36, 36, 48, 0.96);
                --qt-surface-strong: rgba(30, 30, 39, 1);
                --qt-text: #f4f4f8;
                --qt-text-secondary: #d6d6e1;
                --qt-text-muted: #aeaec0;
                --qt-border: rgba(66, 66, 79, 0.8);
                --qt-border-strong: #5a5a70;
                --qt-primary: #b89fff;
                --qt-primary-strong: #9d72ff;
                --qt-primary-deep: #8b4cf7;
            }
        }
        .stApp {
            background: linear-gradient(180deg, var(--qt-bg) 0%, var(--qt-bg-secondary) 100%);
            color: var(--qt-text);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            max-width: 1200px;
        }
        [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
            display: none !important;
        }
        [data-testid="stHeader"], [data-testid="stAppViewContainer"] {
            background: transparent;
        }
        h1, h2, h3, p, label, .stCaption, .stMarkdown, .stText, .stRadio, .stMetric, .stAlert {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            color: var(--qt-text);
        }
        code, pre, .stCode, .stJson {
            font-family: "SFMono-Regular", Menlo, Consolas, monospace !important;
        }
        .doctor-panel {
            background: var(--qt-surface);
            border: 1px solid var(--qt-border);
            border-radius: 24px;
            padding: 1.25rem 1.25rem 0.75rem;
            box-shadow: var(--qt-shadow);
        }
        .queue-note {
            margin: 0.4rem 0 1rem;
            color: var(--qt-text-muted);
            font-size: 0.95rem;
        }
        .work-status {
            display: flex;
            align-items: center;
            gap: 0.9rem;
            padding: 0.95rem 1rem;
            margin: 0.75rem 0 1rem;
            background: var(--qt-surface-strong);
            border: 1px solid var(--qt-border);
            border-radius: 18px;
        }
        .work-status__spinner {
            width: 22px;
            height: 22px;
            border-radius: 999px;
            border: 3px solid rgba(157, 114, 255, 0.20);
            border-top-color: var(--qt-primary);
            animation: spin 0.85s linear infinite;
            flex: 0 0 auto;
        }
        .work-status__title {
            font-weight: 600;
            color: var(--qt-text);
        }
        .work-status__text {
            color: var(--qt-text-secondary);
            font-size: 0.94rem;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .stTextArea textarea,
        .stTextInput input,
        .stSelectbox div[data-baseweb="select"] > div {
            background: var(--qt-surface-strong) !important;
            color: var(--qt-text) !important;
            border: 1px solid var(--qt-border) !important;
            border-radius: 16px !important;
        }
        .stFileUploader section {
            background: var(--qt-surface-strong);
            border: 1px dashed var(--qt-border-strong);
            border-radius: 18px;
        }
        .stButton > button, .stDownloadButton > button {
            border-radius: 999px !important;
            border: 1px solid transparent !important;
            background: linear-gradient(135deg, var(--qt-primary) 0%, var(--qt-primary-strong) 100%) !important;
            color: #ffffff !important;
            font-weight: 600 !important;
        }
        .stButton > button:disabled {
            opacity: 0.55 !important;
        }
        .stRadio [role="radiogroup"] { gap: 0.5rem; }
        .stRadio [role="radiogroup"] label {
            background: var(--qt-surface-strong);
            border: 1px solid var(--qt-border);
            border-radius: 999px;
            padding: 0.4rem 0.8rem;
        }
        [data-testid="stMetric"] {
            background: var(--qt-surface-strong);
            border: 1px solid var(--qt-border);
            border-radius: 18px;
            padding: 0.85rem 1rem;
        }
        .stExpander, .stAlert, .stDataFrame, .stTable, .stJson {
            border-radius: 18px;
            border: 1px solid var(--qt-border) !important;
            overflow: hidden;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_file_configuration(sources: list[dict], disabled: bool) -> None:
    if not sources:
        return

    count = len(sources)
    st.markdown(
        f'<div class="queue-note">{count} file(s) selected. They will be processed sequentially and each finished file will get its own download button.</div>',
        unsafe_allow_html=True,
    )
    if count > 10:
        st.warning("Large batch detected. Processing will run one file at a time to keep the session stable.")

    for idx, item in enumerate(sources):
        ext = item["ext"]
        tabular_default = ext in LEGACY_PREVIEW_EXTS
        support_heal, support_message = heal_support_message(ext, tabular_rescue=tabular_default)
        with st.expander(source_label(item), expanded=idx == 0 and count <= 3):
            st.caption(support_message)
            note = source_note(item)
            if note:
                st.info(note)
            if ext and ext not in SUPPORTED_EXTS:
                st.error(f"Unsupported format: {ext}")
                continue
            if ext not in WORKBOOK_EXTS:
                continue

            if item["source_kind"] == "url":
                sheet_names, same_columns, workbook_error = inspect_remote_url(item["source_label"])
            else:
                sheet_names, same_columns, workbook_error = inspect_local_bytes(item["bytes"], ext)

            if workbook_error:
                st.error(f"Could not inspect workbook sheets: {workbook_error}")
                continue

            st.write(f"Sheets found: {', '.join(sheet_names)}")
            source_id = f"{idx}_{item['source_label']}"
            consolidate_key = f"consolidate_{source_id}"
            sheet_key = f"sheet_{source_id}"
            tabular_key = f"tabular_{source_id}"

            if len(sheet_names) > 1 and same_columns:
                st.checkbox(
                    f"Consolidate sheets for {item['name']}",
                    value=False,
                    key=consolidate_key,
                    disabled=disabled,
                )
            else:
                st.session_state[consolidate_key] = False

            st.selectbox(
                f"Sheet to use for {item['name']}",
                options=sheet_names,
                index=0,
                key=sheet_key,
                disabled=disabled or bool(st.session_state.get(consolidate_key, False)),
            )

            selected_sheet = None if st.session_state.get(consolidate_key, False) else st.session_state.get(sheet_key) or sheet_names[0]
            consolidate = bool(st.session_state.get(consolidate_key, False))
            st.checkbox(
                f"Use tabular rescue mode for {item['name']}",
                value=tabular_default,
                key=tabular_key,
                disabled=disabled,
                help="Use csv-doctor semantic rescue to create a 3-sheet readable workbook. Leave this off to keep workbook-preserving healing for .xlsx/.xlsm.",
            )
            render_mode_details(workbook_mode_details(ext, tabular_rescue=bool(st.session_state.get(tabular_key, tabular_default))))

            if item["source_kind"] == "url":
                inspection, semantic_error = inspect_remote_workbook_semantics(
                    item["source_label"],
                    sheet_name=selected_sheet,
                    consolidate=consolidate,
                )
            else:
                inspection, semantic_error = inspect_local_workbook_semantics(
                    item["bytes"],
                    ext,
                    sheet_name=selected_sheet,
                    consolidate=consolidate,
                )

            if semantic_error:
                st.info(f"Workbook interpretation unavailable: {semantic_error}")
            elif inspection:
                detected_header_key = f"detected_header_{source_id}"
                st.session_state[detected_header_key] = inspection["detected_header_row_number"]
                header_row_key = f"headerrow_{source_id}"
                if header_row_key not in st.session_state:
                    st.session_state[header_row_key] = inspection["detected_header_row_number"]

                header_override = st.number_input(
                    f"Header row for {item['name']}",
                    min_value=1,
                    max_value=max(1, int(inspection.get("original_rows_total", 1))),
                    step=1,
                    key=header_row_key,
                    disabled=disabled,
                    help="Choose which 1-based row should be treated as the true header row before healing.",
                )

                if item["source_kind"] == "url":
                    inspection, semantic_error = inspect_remote_workbook_semantics(
                        item["source_label"],
                        sheet_name=selected_sheet,
                        consolidate=consolidate,
                        header_row_override=int(header_override),
                    )
                else:
                    inspection, semantic_error = inspect_local_workbook_semantics(
                        item["bytes"],
                        ext,
                        sheet_name=selected_sheet,
                        consolidate=consolidate,
                        header_row_override=int(header_override),
                    )

                if semantic_error:
                    st.info(f"Workbook interpretation unavailable: {semantic_error}")
                    continue

                role_overrides: dict[int, str] = {}
                headers = inspection.get("effective_headers") or []
                if headers:
                    st.caption("Override semantic roles")
                    for col_index, header in enumerate(headers, start=1):
                        role_key = f"role_{source_id}_{col_index}"
                        st.selectbox(
                            f"{col_index}. {header}",
                            options=SEMANTIC_ROLE_OPTIONS,
                            key=role_key,
                            disabled=disabled,
                        )
                        selected_role = st.session_state.get(role_key, "auto")
                        if selected_role != "auto":
                            role_overrides[col_index - 1] = selected_role

                if role_overrides:
                    if item["source_kind"] == "url":
                        inspection, semantic_error = inspect_remote_workbook_semantics(
                            item["source_label"],
                            sheet_name=selected_sheet,
                            consolidate=consolidate,
                            header_row_override=int(header_override),
                            role_overrides=role_overrides,
                        )
                    else:
                        inspection, semantic_error = inspect_local_workbook_semantics(
                            item["bytes"],
                            ext,
                            sheet_name=selected_sheet,
                            consolidate=consolidate,
                            header_row_override=int(header_override),
                            role_overrides=role_overrides,
                        )
                    if semantic_error:
                        st.info(f"Workbook interpretation unavailable: {semantic_error}")
                        continue

                render_workbook_semantics(inspection)
                if st.session_state.get(tabular_key, ext in LEGACY_PREVIEW_EXTS):
                    confirm_key = f"confirmplan_{source_id}"
                    signature_key = f"plansignature_{source_id}"
                    plan_signature = json.dumps(
                        {
                            "sheet_name": selected_sheet,
                            "consolidate": consolidate,
                            "header_row_override": int(header_override),
                            "role_overrides": {str(k + 1): v for k, v in sorted(role_overrides.items())},
                            "final_roles": {
                                str(entry["column_index"]): entry["role"]
                                for entry in inspection.get("semantic_columns", [])
                            },
                        },
                        sort_keys=True,
                    )
                    if st.session_state.get(signature_key) != plan_signature:
                        st.session_state[signature_key] = plan_signature
                        st.session_state[confirm_key] = False

                    confirmed = st.checkbox(
                        f"Confirm workbook rescue plan for {item['name']}",
                        key=confirm_key,
                        disabled=disabled,
                        help="Required before tabular rescue runs. The confirmed header row and semantic overrides will be persisted into the JSON healing summary.",
                    )
                    if confirmed:
                        st.success("Workbook rescue plan confirmed.")
                    else:
                        st.warning("Confirm this workbook plan before running tabular rescue.")


def main() -> None:
    set_visuals()
    ensure_state()

    st.title("sheet-doctor")
    st.caption("Upload files or paste public file URLs, describe what you want, and get either a diagnosis or a human-readable output workbook.")

    processing = st.session_state["processing"]

    with st.container():
        st.markdown('<div class="doctor-panel">', unsafe_allow_html=True)
        prompt = st.text_area("What do you want done?", key="prompt_input", height=110, disabled=processing)
        st.file_uploader(
            "Upload files",
            type=[ext.lstrip(".") for ext in sorted(SUPPORTED_EXTS)],
            accept_multiple_files=True,
            key="uploads_input",
            disabled=processing,
        )
        st.text_area(
            "Public file URLs",
            key="public_urls_input",
            height=110,
            disabled=processing,
            placeholder="One public file URL per line. Supports direct links and common public share links from GitHub, Dropbox, Google Drive, OneDrive, Box, and similar hosts.",
        )
        st.caption(
            f"Public URL mode makes outbound network requests and rejects remote files above {MAX_REMOTE_FILE_MB} MB."
        )

        sources = source_items(st.session_state.get("uploads_input") or [], st.session_state.get("public_urls_input", ""))
        if sources:
            render_file_configuration(sources, disabled=processing)

        default_intent = infer_intent(prompt)
        intent = st.radio(
            "Intent",
            options=["Make Readable", "Diagnose Only"],
            index=0 if default_intent == "Make Readable" else 1,
            horizontal=True,
            key="intent_input",
            disabled=processing,
        )
        pending_confirmations = pending_workbook_plan_confirmations(sources, intent) if sources else []
        if pending_confirmations:
            st.warning(
                "Confirm the workbook rescue plan before running: "
                + ", ".join(sorted(pending_confirmations))
            )
        submit = st.button(
            "Run",
            type="primary",
            width="stretch",
            disabled=processing or not sources or bool(pending_confirmations),
        )
        st.markdown("</div>", unsafe_allow_html=True)

    if submit and sources:
        st.session_state["job"] = build_job(prompt, intent, sources)
        st.session_state["results"] = []
        st.session_state["processing"] = True
        st.rerun()

    if st.session_state["processing"]:
        st.info("The system is working on the uploaded files. Please be patient.")
        st.session_state["results"] = process_job(st.session_state["job"])
        st.session_state["processing"] = False
        st.session_state["job"] = []
        st.rerun()

    if not st.session_state.get("results"):
        st.info("Supported here: local uploads or public file URLs for .csv .tsv .txt .xlsx .xls .xlsm .ods .json .jsonl")
        return

    render_results()


if __name__ == "__main__":
    main()
