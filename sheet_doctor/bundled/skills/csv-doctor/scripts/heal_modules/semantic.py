from __future__ import annotations

import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

from column_detector import analyse_dataframe
from heal_modules.normalization import (
    apply_normalisations,
    apply_semantic_normalisations,
    clean_row,
    forward_fill_merged_cell_gaps,
    forward_fill_merged_cell_gaps_generic,
    needs_review,
    needs_review_generic,
    needs_review_semantic,
    normalise_amount,
    normalise_currency,
    normalise_date,
    normalise_name,
    normalise_status,
    parse_amount_like,
)
from heal_modules.preprocessing import (
    GENERIC_QUARANTINE_REASONS,
    QUARANTINE_REASONS,
    classify_raw_row,
    classify_raw_row_generic,
    clean_row_generic,
    detect_formula_row,
    detect_header_band_start_index,
    detect_header_row_index,
    fix_alignment,
    fix_alignment_generic,
    normalise_headers_generic,
    preprocess_rows,
    read_file,
    row_amount_totalish,
    sparse_total_label_row,
)
from heal_modules.shared import (
    ASSUMPTIONS,
    COL,
    GENERIC_ASSUMPTIONS,
    HEADERS,
    LARGE_FILE_SKIP_EXTRAS,
    N_COLS,
    ROLE_HEADER_HINTS,
    SEMANTIC_ASSUMPTIONS,
    STATUS_MAP,
    STATUS_VALUE_HINTS,
    VALID_SEMANTIC_ROLES,
    Change,
    CleanRow,
    QuarantineRow,
    SemanticPlan,
    is_schema_specific_header,
)

def flag_near_duplicates_semantic(
    clean_data: list[CleanRow],
    changelog: list[Change],
    headers: list[str],
    semantic_plan: SemanticPlan,
) -> None:
    if semantic_plan.date_idx is None or semantic_plan.amount_idx is None:
        return

    key_indices = [
        idx for idx, role in semantic_plan.roles_by_index.items()
        if role in {"name", "amount", "currency", "category", "department"}
    ]
    if len(key_indices) < 2:
        return

    nd_index: dict[tuple[str, ...], int] = {}
    for idx, entry in enumerate(clean_data):
        row = entry.row
        key = tuple(row[i] for i in key_indices)
        if key in nd_index:
            prev = clean_data[nd_index[key]]
            d1 = prev.row[semantic_plan.date_idx]
            d2 = entry.row[semantic_plan.date_idx]
            if (
                d1
                and d2
                and re.match(r"^\d{4}-\d{2}-\d{2}$", d1)
                and re.match(r"^\d{4}-\d{2}-\d{2}$", d2)
            ):
                try:
                    delta = abs(
                        (
                            datetime.strptime(d2, "%Y-%m-%d")
                            - datetime.strptime(d1, "%Y-%m-%d")
                        ).days
                    )
                except ValueError:
                    continue
                if delta <= 2:
                    prev.needs_review = True
                    entry.needs_review = True
                    label_idx = semantic_plan.label_idx
                    for flagged, other_date, other_row_num in [
                        (entry, d1, prev.row_num),
                        (prev, d2, entry.row_num),
                    ]:
                        changelog.append(
                            Change(
                                flagged.row_num,
                                "[row]",
                                flagged.row[label_idx] if label_idx < len(flagged.row) else "",
                                "",
                                "Flagged",
                                f"Near-duplicate: same semantic key columns; date {flagged.row[semantic_plan.date_idx]} differs by {delta} day(s) from row {other_row_num} ({other_date})",
                            )
                        )
        else:
            nd_index[key] = idx


# ══════════════════════════════════════════════════════════════════════════
# MAIN PROCESSING LOOP
# ══════════════════════════════════════════════════════════════════════════

def process_schema_specific(
    all_rows: list[list[str]],
    initial_changelog: list[Change] | None = None,
) -> tuple[list[CleanRow], list[QuarantineRow], list[Change]]:
    header_sig = tuple(c.strip().lower() for c in all_rows[0])

    clean_data: list[CleanRow]       = []
    quarantine: list[QuarantineRow]  = []
    changelog:  list[Change]         = list(initial_changelog or [])
    seen_exact: dict[tuple, int]     = {}   # normalized row → original row_num
    running_amount_total = 0.0

    # Skip row 0 (actual column headers); start from row 1 (metadata / first data row)
    data_rows = all_rows[1:]

    for i, raw_row in enumerate(data_rows):
        row_num = i + 2   # 1-based; header = row 1

        # ── Classify on raw values ────────────────────────────────────────
        cls = classify_raw_row(raw_row, header_sig)

        if cls != "NORMAL":
            q_reason = QUARANTINE_REASONS[cls]
            # Light-clean the row for display in Quarantine tab
            q_row = [c.strip() for c in raw_row]
            q_row = (q_row + [""] * N_COLS)[:N_COLS]          # normalise length
            row_id = next((c for c in q_row if c), "[empty]")
            column_hint = "[row]"
            if cls == "FORMULA":
                _, formula_column = detect_formula_row(raw_row, HEADERS)
                column_hint = formula_column or "formula_residue"
            quarantine.append(QuarantineRow(q_row, row_num, q_reason))
            changelog.append(Change(
                row_num, column_hint, row_id[:60], "",
                "Quarantined",
                "formula_residue: Excel formula found, not data" if cls == "FORMULA" else q_reason
            ))
            continue

        # ── Fix alignment ─────────────────────────────────────────────────
        aligned, align_chg = fix_alignment(raw_row, row_num)
        was_padded = align_chg is not None and "padded" in align_chg.reason.lower()
        if align_chg:
            changelog.append(align_chg)

        # ── Clean cells ───────────────────────────────────────────────────
        cleaned, cell_chgs = clean_row(aligned, row_num)
        changelog.extend(cell_chgs)

        # ── Normalise values ──────────────────────────────────────────────
        fixed, norm_chgs = apply_normalisations(cleaned, row_num)
        changelog.extend(norm_chgs)

        label_cell = fixed[COL["Employee Name"]] or fixed[COL["Department"]] or fixed[COL["Category"]]
        if row_amount_totalish(label_cell, fixed[COL["Amount"]], running_amount_total) or sparse_total_label_row(fixed, COL["Employee Name"], COL["Amount"]):
            quarantine.append(QuarantineRow(fixed, row_num, QUARANTINE_REASONS["CALCULATED_SUBTOTAL"]))
            changelog.append(
                Change(
                    row_num,
                    "Amount",
                    fixed[COL["Amount"]],
                    "",
                    "Quarantined",
                    "Calculated subtotal row",
                )
            )
            continue

        was_modified = bool(align_chg or cell_chgs or norm_chgs)

        # ── Exact-duplicate removal ───────────────────────────────────────
        row_key = tuple(fixed)
        if row_key in seen_exact:
            first_num = seen_exact[row_key]
            changelog.append(Change(
                row_num, "[row]", fixed[0], "",
                "Removed", f"Exact duplicate of row {first_num}"
            ))
            continue
        seen_exact[row_key] = row_num

        clean_data.append(CleanRow(
            row          = fixed,
            row_num      = row_num,
            was_modified = was_modified,
            needs_review = needs_review(fixed, was_padded),
        ))
        parsed_amount = parse_amount_like(fixed[COL["Amount"]])
        if parsed_amount is not None:
            running_amount_total += parsed_amount

    # Skip expensive post-processing for very large files
    if len(data_rows) <= LARGE_FILE_SKIP_EXTRAS:
        forward_fill_merged_cell_gaps(clean_data, changelog)

        # ── Near-duplicate detection (second pass on clean_data) ─────────────
        nd_index: dict[tuple, int] = {}   # key → index in clean_data
        for idx, entry in enumerate(clean_data):
            r = entry.row
            key = (r[COL["Employee Name"]], r[COL["Amount"]],
                   r[COL["Currency"]],     r[COL["Category"]])
            if key in nd_index:
                j       = nd_index[key]
                prev    = clean_data[j]
                d1, d2  = prev.row[COL["Date"]], entry.row[COL["Date"]]
                if (d1 and d2
                        and re.match(r"^\d{4}-\d{2}-\d{2}$", d1)
                        and re.match(r"^\d{4}-\d{2}-\d{2}$", d2)):
                    try:
                        delta = abs((datetime.strptime(d2, "%Y-%m-%d")
                                     - datetime.strptime(d1, "%Y-%m-%d")).days)
                        if delta <= 2:
                            prev.needs_review  = True
                            entry.needs_review = True
                            for flagged, other_date, other_row_num in [
                                (entry, d1, prev.row_num),
                                (prev,  d2, entry.row_num),
                            ]:
                                changelog.append(Change(
                                    flagged.row_num, "[row]",
                                    flagged.row[COL["Employee Name"]], "",
                                    "Flagged",
                                    f"Near-duplicate: same Name/Amount/Currency/Category; "
                                    f"date {flagged.row[COL['Date']]} differs by {delta} day(s) "
                                    f"from row {other_row_num} ({other_date})"
                                ))
                    except ValueError:
                        pass
            else:
                nd_index[key] = idx

    return clean_data, quarantine, changelog

def _header_text(header: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (header or "").strip().lower()).strip()


def _header_matches_role(header: str, role: str) -> bool:
    lowered = _header_text(header)
    return any(token in lowered for token in ROLE_HEADER_HINTS.get(role, ()))


def _status_like_column(column_stats: dict) -> bool:
    values = [
        (entry.get("value") or "").strip().lower()
        for entry in column_stats.get("most_common_values", [])
        if (entry.get("value") or "").strip()
    ]
    return bool(values) and sum(1 for value in values if value in STATUS_VALUE_HINTS) >= max(1, len(values) // 2)


def _average_sample_length(column_stats: dict) -> float:
    samples = [value for value in column_stats.get("sample_values", []) if value]
    if not samples:
        return 0.0
    return sum(len(value) for value in samples) / len(samples)


def _semantic_role_scores(header: str, column_stats: dict) -> dict[str, float]:
    detected_type = column_stats.get("detected_type", "unknown")
    header_text = _header_text(header)
    scores = {
        "identifier": 0.0,
        "name": 0.0,
        "date": 0.0,
        "amount": 0.0,
        "measurement": 0.0,
        "currency": 0.0,
        "status": 0.0,
        "department": 0.0,
        "category": 0.0,
        "notes": 0.0,
    }

    if detected_type == "name":
        scores["name"] += 0.45
    if detected_type == "date":
        scores["date"] += 0.72
    if detected_type == "ID/code":
        scores["identifier"] += 0.72
    if detected_type == "currency/amount":
        scores["amount"] += 0.72
        scores["measurement"] += 0.24
    if detected_type == "plain number":
        scores["amount"] += 0.42
        scores["measurement"] += 0.45
    if detected_type == "percentage":
        scores["measurement"] += 0.35
    if detected_type == "currency code":
        scores["currency"] += 0.72
    if detected_type == "boolean":
        scores["status"] += 0.20
    if detected_type == "categorical":
        scores["status"] += 0.12
        scores["department"] += 0.12
        scores["category"] += 0.12
    if detected_type == "free text":
        scores["notes"] += 0.20

    for role in scores:
        if _header_matches_role(header, role):
            if role == "identifier":
                scores[role] += 0.82 if re.search(r"\b(id|code)\b", header_text) else 0.68
            elif role == "name":
                scores[role] += 0.82 if re.search(r"\bname\b", header_text) else 0.68
            elif role == "currency":
                scores[role] += 0.68
            elif role in {"date", "amount"}:
                scores[role] += 0.32
            elif role == "measurement":
                scores[role] += 0.72
            elif role == "department":
                scores[role] += 0.82 if re.search(r"\b(ward|clinic|division|department|dept|team|unit|function|location)\b", header_text) else 0.68
            else:
                scores[role] += 0.68

    if _status_like_column(column_stats):
        scores["status"] += 0.28
    if detected_type == "free text" and _average_sample_length(column_stats) >= 20:
        scores["notes"] += 0.12

    if not _header_matches_role(header, "name") and any(
        _header_matches_role(header, role_name)
        for role_name in ("identifier", "measurement", "department", "category", "status", "date")
    ):
        scores["name"] = min(scores["name"], 0.40)
    if re.search(r"\b(month|day|year)\b", header_text) and not re.search(r"\b(date|dob|dofb)\b", header_text):
        scores["date"] = min(scores["date"], 0.40)
    if _header_matches_role(header, "measurement"):
        scores["notes"] = min(scores["notes"], 0.20)

    return {role: min(score, 0.99) for role, score in scores.items()}


def build_semantic_plan(
    headers: list[str],
    raw_rows: list[list[str]],
    delimiter: str,
    role_overrides: dict[int, str] | None = None,
) -> SemanticPlan:
    n_cols = len(headers)
    preview_rows: list[list[str]] = []

    for idx, raw_row in enumerate(raw_rows[:1000], start=2):
        aligned, _, _ = fix_alignment_generic(raw_row, idx, n_cols, delimiter)
        cleaned, _ = clean_row_generic(aligned, idx, headers)
        preview_rows.append(cleaned)

    if not preview_rows:
        return SemanticPlan(False, {}, {}, 0, None, None, None, [])

    analysis = analyse_dataframe(pd.DataFrame(preview_rows, columns=headers))
    columns = analysis.get("columns", {})
    candidate_scores = {
        index: _semantic_role_scores(header, columns.get(header, {}))
        for index, header in enumerate(headers)
    }

    thresholds = {
        "identifier": 0.60,
        "name": 0.60,
        "date": 0.60,
        "amount": 0.60,
        "measurement": 0.60,
        "currency": 0.60,
        "status": 0.72,
        "department": 0.72,
        "category": 0.72,
        "notes": 0.72,
    }

    assignments: dict[int, str] = {}
    confidences: dict[int, float] = {}
    taken_indices: set[int] = set()

    for role in ("identifier", "name", "date", "amount", "measurement", "currency", "status", "department", "category", "notes"):
        best_idx = None
        best_score = 0.0
        for idx, scores in candidate_scores.items():
            if idx in taken_indices:
                continue
            score = scores.get(role, 0.0)
            if score > best_score:
                best_idx = idx
                best_score = score
        if best_idx is not None and best_score >= thresholds[role]:
            assignments[best_idx] = role
            confidences[best_idx] = round(best_score, 2)
            taken_indices.add(best_idx)

    def assign_unique_detected_type(
        role: str, detected_type: str, minimum: float = 0.58
    ) -> None:
        if role in assignments.values():
            return
        candidates = [
            idx for idx, header in enumerate(headers)
            if idx not in taken_indices and columns.get(header, {}).get("detected_type") == detected_type
        ]
        if len(candidates) == 1:
            idx = candidates[0]
            score = candidate_scores[idx].get(role, 0.0)
            if score >= minimum:
                assignments[idx] = role
                confidences[idx] = round(score, 2)
                taken_indices.add(idx)

    assign_unique_detected_type("identifier", "ID/code")
    assign_unique_detected_type("name", "name")
    assign_unique_detected_type("date", "date")
    assign_unique_detected_type("amount", "currency/amount")
    assign_unique_detected_type("currency", "currency code")
    assign_unique_detected_type("notes", "free text", minimum=0.55)

    for idx, scores in candidate_scores.items():
        if idx in taken_indices:
            continue
        score = scores.get("measurement", 0.0)
        if score >= thresholds["measurement"]:
            assignments[idx] = "measurement"
            confidences[idx] = round(score, 2)
            taken_indices.add(idx)

    role_overrides = role_overrides or {}
    if role_overrides:
        for idx, role in role_overrides.items():
            if idx < 0 or idx >= len(headers):
                continue
            assignments.pop(idx, None)
            confidences.pop(idx, None)
            if role == "ignore":
                continue
            for assigned_idx, assigned_role in list(assignments.items()):
                if assigned_role == role:
                    assignments.pop(assigned_idx, None)
                    confidences.pop(assigned_idx, None)
            assignments[idx] = role
            confidences[idx] = 1.0

    primary_roles = set(assignments.values())
    enabled = False
    if "amount" in primary_roles and len(primary_roles.intersection({"name", "date", "currency", "status", "department", "category"})) >= 2:
        enabled = True
    elif len(primary_roles.intersection({"identifier", "name", "date", "status", "department", "category", "notes", "measurement"})) >= 3 and primary_roles.intersection({"identifier", "date", "measurement"}):
        enabled = True
    if not enabled:
        return SemanticPlan(False, {}, {}, 0, None, None, None, [])

    label_idx = next(
        (idx for role_name in ("name", "identifier", "department", "category", "notes") for idx, assigned in assignments.items() if assigned == role_name),
        0,
    )
    fill_down_indices = [
        idx for idx, role in assignments.items()
        if role in {"department", "category", "status", "currency"}
    ]

    return SemanticPlan(
        enabled=True,
        roles_by_index=assignments,
        confidence_by_index=confidences,
        label_idx=label_idx,
        amount_idx=next((idx for idx, role in assignments.items() if role == "amount"), None),
        currency_idx=next((idx for idx, role in assignments.items() if role == "currency"), None),
        date_idx=next((idx for idx, role in assignments.items() if role == "date"), None),
        fill_down_indices=fill_down_indices,
    )


def _semantic_columns_from_plan(headers: list[str], semantic_plan: SemanticPlan) -> list[dict]:
    return [
        {
            "column_index": idx + 1,
            "header": headers[idx],
            "role": role,
            "confidence": semantic_plan.confidence_by_index.get(idx, 0.0),
        }
        for idx, role in sorted(semantic_plan.roles_by_index.items())
    ]


def _semantic_mapping_comparison(
    headers: list[str],
    suggested_plan: SemanticPlan,
    effective_plan: SemanticPlan,
    role_overrides: dict[int, str] | None = None,
) -> list[dict]:
    overrides = role_overrides or {}
    rows: list[dict] = []
    for idx, header in enumerate(headers):
        suggested_role = suggested_plan.roles_by_index.get(idx)
        effective_role = effective_plan.roles_by_index.get(idx)
        override_role = overrides.get(idx)
        rows.append(
            {
                "column_index": idx + 1,
                "header": header,
                "detected_role": suggested_role or "",
                "detected_confidence": suggested_plan.confidence_by_index.get(idx, 0.0),
                "override_role": override_role or "",
                "final_role": effective_role or "",
                "final_confidence": effective_plan.confidence_by_index.get(idx, 0.0),
            }
        )
    return rows


def inspect_healing_plan(
    input_path: Path,
    *,
    sheet_name: str | None = None,
    consolidate_sheets: bool | None = None,
    header_row_override: int | None = None,
    role_overrides: dict[int, str] | None = None,
) -> dict:
    all_rows, delimiter = read_file(
        input_path,
        sheet_name=sheet_name,
        consolidate_sheets=consolidate_sheets,
    )
    if not all_rows:
        raise ValueError("File is empty.")

    header_idx = detect_header_row_index(all_rows, explicit_header_row=header_row_override)
    header_band_start = detect_header_band_start_index(all_rows, header_idx)
    preprocessed_rows, preprocessing_changes = preprocess_rows(
        all_rows,
        explicit_header_row=header_row_override,
    )
    if not preprocessed_rows:
        raise ValueError("No usable rows remain after preprocessing.")

    metadata_rows_removed = sum(
        1 for change in preprocessing_changes if change.column_affected == "[file metadata]"
    )
    header_band_merged = any(
        change.column_affected == "[header band]" for change in preprocessing_changes
    )

    raw_header = preprocessed_rows[0]
    if is_schema_specific_header(raw_header) and not role_overrides:
        headers = HEADERS[:]
        mode = "schema-specific"
        semantic_columns = [
            {
                "column_index": idx + 1,
                "header": header,
                "role": role,
                "confidence": 0.99,
            }
            for idx, (header, role) in enumerate(
                zip(
                    headers,
                    ["name", "department", "date", "amount", "currency", "category", "status", "notes"],
                )
            )
        ]
        suggested_semantic_columns = list(semantic_columns)
        semantic_comparison = [
            {
                "column_index": idx + 1,
                "header": header,
                "detected_role": role,
                "detected_confidence": 0.99,
                "override_role": (role_overrides or {}).get(idx, ""),
                "final_role": (role_overrides or {}).get(idx, role) if (role_overrides or {}).get(idx) != "ignore" else "",
                "final_confidence": 1.0 if idx in (role_overrides or {}) and (role_overrides or {}).get(idx) != "ignore" else 0.99,
            }
            for idx, (header, role) in enumerate(
                zip(
                    headers,
                    ["name", "department", "date", "amount", "currency", "category", "status", "notes"],
                )
            )
        ]
    else:
        headers, _ = normalise_headers_generic(raw_header)
        suggested_plan = build_semantic_plan(
            headers,
            preprocessed_rows[1:],
            delimiter,
            role_overrides=None,
        )
        semantic_plan = build_semantic_plan(headers, preprocessed_rows[1:], delimiter, role_overrides=role_overrides)
        mode = "semantic" if semantic_plan.enabled else "generic"
        suggested_semantic_columns = _semantic_columns_from_plan(headers, suggested_plan)
        semantic_columns = _semantic_columns_from_plan(headers, semantic_plan)
        semantic_comparison = _semantic_mapping_comparison(headers, suggested_plan, semantic_plan, role_overrides=role_overrides)

    return {
        "delimiter": delimiter,
        "original_rows_total": len(all_rows),
        "detected_header_row_number": header_idx + 1,
        "detected_header_band_rows": list(range(header_band_start + 1, header_idx + 2)),
        "metadata_rows_removed": metadata_rows_removed,
        "header_band_merged": header_band_merged,
        "effective_headers": headers,
        "healing_mode_candidate": mode,
        "suggested_semantic_columns": suggested_semantic_columns,
        "semantic_columns": semantic_columns,
        "semantic_comparison": semantic_comparison,
        "applied_role_overrides": {
            str(idx + 1): role for idx, role in sorted((role_overrides or {}).items())
        },
    }



def process_generic(
    all_rows: list[list[str]],
    delimiter: str,
    initial_changelog: list[Change] | None = None,
    role_overrides: dict[int, str] | None = None,
) -> tuple[list[CleanRow], list[QuarantineRow], list[Change], list[str], str]:
    headers, header_changes = normalise_headers_generic(all_rows[0])
    n_cols = len(headers)
    header_sig = tuple(h.strip().lower() for h in headers)
    semantic_plan = build_semantic_plan(headers, all_rows[1:], delimiter, role_overrides=role_overrides)
    applied_mode = "semantic" if semantic_plan.enabled else "generic"

    clean_data: list[CleanRow] = []
    quarantine: list[QuarantineRow] = []
    changelog: list[Change] = list(initial_changelog or []) + list(header_changes)
    seen_exact: dict[tuple, int] = {}
    running_amount_total = 0.0
    amount_idx = semantic_plan.amount_idx
    if amount_idx is None:
        amount_idx = next((i for i, header in enumerate(headers) if "amount" in header.lower() or "total" in header.lower()), None)
    label_idx = semantic_plan.label_idx if semantic_plan.enabled else 0

    for i, raw_row in enumerate(all_rows[1:], start=2):
        cls = classify_raw_row_generic(raw_row, header_sig, n_cols)
        if cls != "NORMAL":
            q_reason = GENERIC_QUARANTINE_REASONS[cls]
            q_row = [c.strip() for c in raw_row]
            q_row = (q_row + [""] * n_cols)[:n_cols]
            row_id = next((c for c in q_row if c), "[empty]")
            column_hint = "[row]"
            if cls == "FORMULA":
                _, formula_column = detect_formula_row(raw_row, headers)
                column_hint = formula_column or "formula_residue"
            quarantine.append(QuarantineRow(q_row, i, q_reason))
            changelog.append(
                Change(
                    i,
                    column_hint,
                    row_id[:60],
                    "",
                    "Quarantined",
                    "formula_residue: Excel formula found, not data" if cls == "FORMULA" else q_reason,
                )
            )
            continue

        aligned, align_change, structure_changed = fix_alignment_generic(raw_row, i, n_cols, delimiter)
        if align_change:
            changelog.append(align_change)

        cleaned, cell_changes = clean_row_generic(aligned, i, headers)
        changelog.extend(cell_changes)
        semantic_changes: list[Change] = []
        if semantic_plan.enabled:
            cleaned, semantic_changes = apply_semantic_normalisations(cleaned, i, headers, semantic_plan)
            changelog.extend(semantic_changes)
        was_modified = bool(align_change or cell_changes or semantic_changes)

        label_text = cleaned[label_idx] if label_idx < len(cleaned) else ""
        amount_text = cleaned[amount_idx] if amount_idx is not None and amount_idx < len(cleaned) else ""
        if amount_idx is not None and (
            row_amount_totalish(label_text, amount_text, running_amount_total)
            or sparse_total_label_row(cleaned, label_idx, amount_idx)
        ):
            quarantine.append(QuarantineRow(cleaned, i, GENERIC_QUARANTINE_REASONS["CALCULATED_SUBTOTAL"]))
            changelog.append(Change(i, headers[amount_idx], amount_text, "", "Quarantined", "Calculated subtotal row"))
            continue

        row_key = tuple(cleaned)
        if row_key in seen_exact:
            first_row = seen_exact[row_key]
            changelog.append(
                Change(i, "[row]", cleaned[0] if cleaned else "", "", "Removed", f"Exact duplicate of row {first_row}")
            )
            continue
        seen_exact[row_key] = i

        clean_data.append(
            CleanRow(
                row=cleaned,
                row_num=i,
                was_modified=was_modified,
                needs_review=(
                    needs_review_semantic(cleaned, structure_changed, semantic_plan)
                    if semantic_plan.enabled
                    else needs_review_generic(cleaned, structure_changed)
                ),
            )
        )
        if amount_idx is not None:
            parsed_amount = parse_amount_like(cleaned[amount_idx])
            if parsed_amount is not None:
                running_amount_total += parsed_amount

    if semantic_plan.enabled and semantic_plan.fill_down_indices:
        if len(all_rows) - 1 <= LARGE_FILE_SKIP_EXTRAS:
            forward_fill_merged_cell_gaps_generic(clean_data, changelog, headers, semantic_plan.fill_down_indices)
            flag_near_duplicates_semantic(clean_data, changelog, headers, semantic_plan)

    return clean_data, quarantine, changelog, headers, applied_mode


def execute_healing_pipeline(
    input_path: Path,
    *,
    sheet_name: str | None = None,
    consolidate_sheets: bool | None = None,
    header_row_override: int | None = None,
    role_overrides: dict[int, str] | None = None,
) -> dict:
    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")

    all_rows, delimiter = read_file(
        input_path,
        sheet_name=sheet_name,
        consolidate_sheets=consolidate_sheets,
    )
    if len(all_rows) < 2:
        raise ValueError("File is empty or has only a header.")

    original_total_in = len(all_rows)
    all_rows, metadata_changes = preprocess_rows(
        all_rows,
        explicit_header_row=header_row_override,
    )
    if len(all_rows) < 2:
        raise ValueError("File is empty after metadata/header detection.")

    if is_schema_specific_header(all_rows[0]) and not role_overrides:
        mode = "schema-specific"
        clean_data, quarantine, changelog = process_schema_specific(
            all_rows,
            initial_changelog=metadata_changes,
        )
        headers = HEADERS
        assumptions = ASSUMPTIONS
    else:
        clean_data, quarantine, changelog, headers, mode = process_generic(
            all_rows,
            delimiter,
            initial_changelog=metadata_changes,
            role_overrides=role_overrides,
        )
        assumptions = SEMANTIC_ASSUMPTIONS if mode == "semantic" else GENERIC_ASSUMPTIONS

    action_counts = Counter(c.action_taken for c in changelog)
    quarantine_reason_counts = {
        reason: sum(1 for q in quarantine if q.reason == reason)
        for reason in {q.reason for q in quarantine}
    }

    return {
        "input_path": input_path,
        "delimiter": delimiter,
        "total_in": original_total_in,
        "mode": mode,
        "headers": headers,
        "assumptions": assumptions,
        "clean_data": clean_data,
        "quarantine": quarantine,
        "changelog": changelog,
        "action_counts": action_counts,
        "quarantine_reason_counts": dict(sorted(quarantine_reason_counts.items())),
    }
