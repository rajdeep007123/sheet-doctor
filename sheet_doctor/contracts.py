"""Shared versioned contracts for deployable sheet-doctor outputs."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

CONTRACT_VERSIONS = {
    "csv_doctor.diagnose": "1.0.0",
    "csv_doctor.report": "1.0.0",
    "csv_doctor.heal_summary": "1.0.0",
    "excel_doctor.diagnose": "1.0.0",
    "excel_doctor.heal_summary": "1.0.0",
}


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def build_contract(name: str) -> dict[str, str]:
    version = CONTRACT_VERSIONS[name]
    return {"name": name, "version": version}



def build_run_summary(
    *,
    tool: str,
    script: str,
    input_path: Path,
    status: str = "ok",
    output_path: Path | None = None,
    metrics: dict[str, Any] | None = None,
    warnings: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "tool": tool,
        "script": script,
        "status": status,
        "generated_at": utc_now_iso(),
        "input_file": str(input_path),
        "output_file": str(output_path) if output_path else None,
        "warnings_count": len(warnings or []),
        "warnings": list(warnings or []),
        "metrics": metrics or {},
    }
