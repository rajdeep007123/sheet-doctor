#!/usr/bin/env python3
"""
csv-doctor column_detector.py

Analyses each column in a tabular file and infers what the column likely
contains even when headers are weak or missing. Outputs structured JSON.

Usage:
    python column_detector.py <path-to-file>
"""

from __future__ import annotations

import json
import math
import random
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from loader import load_file


TYPE_ORDER = [
    "date",
    "currency/amount",
    "plain number",
    "percentage",
    "email address",
    "phone number",
    "URL",
    "country name or code",
    "currency code",
    "name",
    "categorical",
    "free text",
    "boolean",
    "ID/code",
    "unknown",
]

DATE_FORMAT_PATTERNS = [
    ("%Y-%m-%d", re.compile(r"^\d{4}-\d{2}-\d{2}$"), "YYYY-MM-DD"),
    ("%Y/%m/%d", re.compile(r"^\d{4}/\d{2}/\d{2}$"), "YYYY/MM/DD"),
    ("%d/%m/%Y", re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$"), "DD/MM/YYYY"),
    ("%m/%d/%Y", re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$"), "MM/DD/YYYY"),
    ("%d-%m-%Y", re.compile(r"^\d{1,2}-\d{1,2}-\d{4}$"), "DD-MM-YYYY"),
    ("%m-%d-%Y", re.compile(r"^\d{1,2}-\d{1,2}-\d{4}$"), "MM-DD-YYYY"),
    ("%d/%m/%y", re.compile(r"^\d{1,2}/\d{1,2}/\d{2}$"), "DD/MM/YY"),
    ("%m/%d/%y", re.compile(r"^\d{1,2}/\d{1,2}/\d{2}$"), "MM/DD/YY"),
    ("%d-%m-%y", re.compile(r"^\d{1,2}-\d{1,2}-\d{2}$"), "DD-MM-YY"),
    ("%m-%d-%y", re.compile(r"^\d{1,2}-\d{1,2}-\d{2}$"), "MM-DD-YY"),
    ("%B %d %Y", re.compile(r"^[A-Za-z]+\s+\d{1,2}\s+\d{4}$"), "Month D YYYY"),
    ("%b %d %Y", re.compile(r"^[A-Za-z]{3}\s+\d{1,2}\s+\d{4}$"), "Mon D YYYY"),
    ("%d %B %Y", re.compile(r"^\d{1,2}\s+[A-Za-z]+\s+\d{4}$"), "D Month YYYY"),
    ("%d %b %Y", re.compile(r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}$"), "D Mon YYYY"),
    ("%B %d, %Y", re.compile(r"^[A-Za-z]+\s+\d{1,2},\s+\d{4}$"), "Month D, YYYY"),
    ("%b %d, %Y", re.compile(r"^[A-Za-z]{3}\s+\d{1,2},\s+\d{4}$"), "Mon D, YYYY"),
]

BOOLEAN_TRUE = {"true", "yes", "y", "1", "approved", "approve"}
BOOLEAN_FALSE = {"false", "no", "n", "0", "rejected", "reject"}

CURRENCY_CODES = {
    "AED", "AUD", "BRL", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "INR",
    "JPY", "KRW", "MXN", "NOK", "NZD", "PLN", "RUB", "SEK", "SGD", "TRY",
    "USD", "ZAR",
}

COUNTRY_CODES = {
    "AE", "AU", "BR", "CA", "CH", "CN", "DE", "ES", "FR", "GB", "HK", "IN",
    "IT", "JP", "KR", "MX", "NL", "NZ", "PL", "RU", "SE", "SG", "TR", "US",
    "ZA",
}

COUNTRY_NAMES = {
    "argentina", "australia", "austria", "belgium", "brazil", "canada", "china",
    "denmark", "finland", "france", "germany", "hong kong", "india", "indonesia",
    "ireland", "italy", "japan", "kenya", "malaysia", "mexico", "netherlands",
    "new zealand", "norway", "pakistan", "poland", "portugal", "russia",
    "saudi arabia", "singapore", "south africa", "south korea", "spain", "sweden",
    "switzerland", "thailand", "turkey", "uae", "uk", "united arab emirates",
    "united kingdom", "united states", "united states of america", "usa", "us",
    "vietnam",
}

EMAIL_RE = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.IGNORECASE)
PHONE_RE = re.compile(
    r"^(?:\+?\d{1,3}[\s().-]*)?(?:\(?\d{2,4}\)?[\s().-]*)?\d(?:[\d\s().-]{5,}\d)$"
)
URL_RE = re.compile(r"^(?:https?://|www\.)\S+$", re.IGNORECASE)
PERCENT_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?%$")
ID_RE = re.compile(r"^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9][A-Za-z0-9._/\-]{2,}$")
NAME_RE = re.compile(r"^[A-Za-z][A-Za-z'`.-]*(?:\s+[A-Za-z][A-Za-z'`.-]*){1,3}$")
WHITESPACE_RE = re.compile(r"^\s+|\s+$")
SENTINEL_NULLS = {"", "na", "n/a", "none", "null", "nil", "nan", "tbd", "-"}
NAME_STOPWORDS = {
    "a", "an", "and", "at", "before", "by", "for", "from", "in", "into", "of",
    "on", "or", "said", "the", "to", "with", "was", "were", "worth",
}


def normalize_scalar(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        if value.tzinfo is not None:
            value = value.tz_convert(None)
        return value.to_pydatetime()
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone(timezone.utc).replace(tzinfo=None)
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return None
        return value
    text = str(value).replace("\x00", "")
    return text


def normalize_datetime(value: datetime) -> datetime:
    if value.tzinfo is not None:
        return value.astimezone(timezone.utc).replace(tzinfo=None)
    return value


def is_effective_null(value: Any) -> bool:
    normalized = normalize_scalar(value)
    if normalized is None:
        return True
    if isinstance(normalized, str):
        return normalized.strip().lower() in SENTINEL_NULLS
    return False


def stringify(value: Any) -> str:
    normalized = normalize_scalar(value)
    if normalized is None:
        return ""
    if isinstance(normalized, datetime):
        return normalized.isoformat(sep=" ")
    return str(normalized)


def canonical_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def header_hint(column_name: str) -> str | None:
    lowered = column_name.strip().lower()
    hints = [
        ("date", "date"),
        ("amount", "currency/amount"),
        ("price", "currency/amount"),
        ("cost", "currency/amount"),
        ("currency", "currency code"),
        ("email", "email address"),
        ("phone", "phone number"),
        ("mobile", "phone number"),
        ("url", "URL"),
        ("website", "URL"),
        ("country", "country name or code"),
        ("nation", "country name or code"),
        ("name", "name"),
        ("notes", "free text"),
        ("comment", "free text"),
        ("description", "free text"),
        ("message", "free text"),
        ("status", "categorical"),
        ("type", "categorical"),
        ("category", "categorical"),
        ("id", "ID/code"),
        ("code", "ID/code"),
        ("percent", "percentage"),
        ("ratio", "percentage"),
        ("flag", "boolean"),
        ("is_", "boolean"),
        ("has_", "boolean"),
    ]
    for needle, inferred in hints:
        if needle in lowered:
            return inferred
    return None


def looks_like_name(text: str) -> bool:
    if not NAME_RE.fullmatch(text):
        return False
    tokens = text.split()
    lower_tokens = [token.lower().strip(".'`-") for token in tokens]
    if any(token in NAME_STOPWORDS for token in lower_tokens):
        return False
    if sum(1 for token in tokens if token[:1].isupper()) < 2 and not all(token.isupper() for token in tokens):
        return False
    return True


def maybe_parse_number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    normalized = normalize_scalar(value)
    if normalized is None:
        return None
    if isinstance(normalized, (int, float)):
        return float(normalized)
    text = str(normalized).strip()
    if not text:
        return None
    if text.lower() in SENTINEL_NULLS:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = text.replace(" ", "")
    text = re.sub(r"^[A-Z]{3}", "", text)
    text = re.sub(r"(USD|EUR|INR|GBP|JPY|CAD|AUD|AED|CHF)$", "", text, flags=re.I)
    text = text.replace("$", "").replace("€", "").replace("£", "").replace("¥", "").replace("₹", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") == 1 and text.count(".") == 0:
        left, right = text.split(",", 1)
        if len(right) == 2:
            text = f"{left}.{right}"
        elif len(right) == 3:
            text = text.replace(",", "")
    else:
        text = text.replace(",", "")

    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text):
        return None

    number = float(text)
    return -number if negative else number


def maybe_parse_percentage(value: Any) -> float | None:
    text = stringify(value).strip()
    if not PERCENT_RE.fullmatch(text):
        return None
    try:
        return float(text[:-1])
    except ValueError:
        return None


def maybe_parse_date(value: Any) -> tuple[datetime, str] | None:
    normalized = normalize_scalar(value)
    if normalized is None:
        return None
    if isinstance(normalized, datetime):
        return normalize_datetime(normalized), "native datetime"
    if isinstance(normalized, (int, float)) and not isinstance(normalized, bool):
        number = float(normalized)
        if 25000 <= number <= 60000:
            parsed = pd.to_datetime(number, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(parsed):
                return normalize_datetime(parsed.to_pydatetime()), "excel serial"
        if 946684800 <= number <= 4102444800:
            parsed = pd.to_datetime(number, unit="s", errors="coerce")
            if not pd.isna(parsed):
                return normalize_datetime(parsed.to_pydatetime()), "unix timestamp"
        return None

    text = stringify(normalized).strip()
    if not text or text.lower() in SENTINEL_NULLS:
        return None

    for fmt, pattern, label in DATE_FORMAT_PATTERNS:
        if not pattern.fullmatch(text):
            continue
        try:
            return normalize_datetime(datetime.strptime(text, fmt)), label
        except ValueError:
            continue

    parsed = pd.to_datetime(text, errors="coerce", utc=False)
    if pd.isna(parsed):
        return None
    if isinstance(parsed, pd.Timestamp):
        return normalize_datetime(parsed.to_pydatetime()), "inferred datetime"
    return None


def detect_atomic_type(value: Any) -> str:
    normalized = normalize_scalar(value)
    if normalized is None:
        return "unknown"
    if isinstance(normalized, bool):
        return "boolean"
    if maybe_parse_date(normalized):
        return "date"

    text = stringify(normalized).strip()
    lower = text.lower()
    if not text or lower in SENTINEL_NULLS:
        return "unknown"
    if EMAIL_RE.fullmatch(text):
        return "email address"
    if URL_RE.fullmatch(text):
        return "URL"
    if PHONE_RE.fullmatch(text) and sum(ch.isdigit() for ch in text) >= 7:
        return "phone number"
    if maybe_parse_percentage(text) is not None:
        return "percentage"
    if text.upper() in CURRENCY_CODES:
        return "currency code"
    if lower in COUNTRY_NAMES or text.upper() in COUNTRY_CODES:
        return "country name or code"
    if lower in BOOLEAN_TRUE or lower in BOOLEAN_FALSE:
        return "boolean"
    if any(symbol in text for symbol in ("$", "€", "£", "¥", "₹")) or re.search(
        r"\b(?:USD|EUR|INR|GBP|JPY|CAD|AUD|AED|CHF)\b", text, re.IGNORECASE
    ):
        if maybe_parse_number(text) is not None:
            return "currency/amount"
    if maybe_parse_number(text) is not None:
        return "plain number"
    if ID_RE.fullmatch(text):
        return "ID/code"
    if looks_like_name(text):
        return "name"
    return "free text"


def capitalization_signature(value: str) -> str:
    tokens = re.findall(r"[A-Za-z]+", value)
    if not tokens:
        return "other"
    if all(token.isupper() for token in tokens):
        return "upper"
    if all(token.islower() for token in tokens):
        return "lower"
    if all(token[:1].isupper() and token[1:].islower() for token in tokens if len(token) > 1):
        return "title"
    return "mixed"


def infer_column_type(
    column_name: str,
    series: pd.Series,
    atomic_counts: Counter,
    non_null_texts: list[str],
) -> str:
    non_null_count = len(non_null_texts)
    if non_null_count == 0:
        return "unknown"

    score = Counter()
    for kind, count in atomic_counts.items():
        score[kind] += count / non_null_count

    unique_count = len(set(non_null_texts))
    avg_length = sum(len(text) for text in non_null_texts) / non_null_count
    hint = header_hint(column_name)

    if hint == "free text" and (avg_length >= 20 or score["free text"] >= 0.35):
        return "free text"
    if hint == "currency/amount" and (score["currency/amount"] + score["plain number"]) >= 0.6:
        return "currency/amount"
    if hint == "currency code" and score["currency code"] >= 0.45:
        return "currency code"
    if hint == "date" and score["date"] >= 0.35:
        return "date"
    if hint == "name" and score["name"] >= 0.35:
        return "name"
    if hint == "percentage" and (score["percentage"] + score["plain number"]) >= 0.6:
        return "percentage"
    if hint == "categorical" and unique_count <= max(16, int(non_null_count * 0.4)):
        return "categorical"
    if hint == "boolean" and score["boolean"] >= 0.5:
        return "boolean"
    if hint == "ID/code" and score["ID/code"] >= 0.35:
        return "ID/code"

    if score["boolean"] >= 0.8 and unique_count <= 6:
        return "boolean"
    if score["email address"] >= 0.6:
        return "email address"
    if score["phone number"] >= 0.6:
        return "phone number"
    if score["URL"] >= 0.6:
        return "URL"
    if score["currency code"] >= 0.7:
        return "currency code"
    if score["country name or code"] >= 0.7:
        return "country name or code"
    if score["percentage"] >= 0.7:
        return "percentage"
    if score["currency/amount"] >= 0.55:
        return "currency/amount"
    if score["date"] >= 0.55:
        return "date"
    if score["plain number"] >= 0.75:
        return "plain number"
    if score["ID/code"] >= 0.55 and unique_count / non_null_count >= 0.6:
        return "ID/code"
    if score["name"] >= 0.55:
        return "name"
    if unique_count <= max(12, int(non_null_count * 0.2)) and avg_length <= 24:
        return "categorical"
    if avg_length >= 35 or score["free text"] >= 0.55:
        return "free text"
    return "unknown"


def compute_numeric_range(values: list[float]) -> tuple[float | None, float | None]:
    if not values:
        return None, None
    return min(values), max(values)


def compute_date_range(values: list[datetime]) -> tuple[str | None, str | None]:
    if not values:
        return None, None
    return min(values).date().isoformat(), max(values).date().isoformat()


def sample_examples(values: list[str], count: int = 3) -> list[str]:
    if not values:
        return []
    unique_values = list(dict.fromkeys(values))
    if len(unique_values) <= count:
        return unique_values
    rng = random.Random(42)
    return rng.sample(unique_values, count)


def detect_suspected_issues(
    series: pd.Series,
    detected_type: str,
    non_null_texts: list[str],
    atomic_types: list[str],
    date_labels: list[str],
    numeric_values: list[float],
) -> list[str]:
    issues: list[str] = []
    non_null_count = len(non_null_texts)
    if non_null_count == 0:
        return issues

    if len(set(label for label in date_labels if label)) > 1:
        issues.append("Mixed date formats detected")

    raw_non_null = [stringify(value) for value in series if not is_effective_null(value)]
    whitespace_count = sum(1 for value in raw_non_null if isinstance(value, str) and WHITESPACE_RE.search(value))
    if whitespace_count:
        pct = round((whitespace_count / non_null_count) * 100, 1)
        issues.append(f"Trailing/leading whitespace in {pct}% of values")

    case_map: dict[str, set[str]] = defaultdict(set)
    for value in non_null_texts:
        lowered = value.lower()
        case_map[lowered].add(value)
    if any(len(variants) > 1 for variants in case_map.values()):
        issues.append("Inconsistent capitalisation")
    else:
        caps = Counter(capitalization_signature(value) for value in non_null_texts if re.search(r"[A-Za-z]", value))
        if len([kind for kind, count in caps.items() if count >= 2]) > 1:
            issues.append("Inconsistent capitalisation")

    canonical_map: dict[str, set[str]] = defaultdict(set)
    for value in non_null_texts:
        canonical = canonical_text(value)
        if canonical:
            canonical_map[canonical].add(value.strip())
    if any(len(variants) > 1 for variants in canonical_map.values()):
        issues.append("Possible duplicates with slight differences")

    value_counts = Counter(value.strip().lower() for value in non_null_texts)
    top_count = value_counts.most_common(1)[0][1]
    if top_count / non_null_count >= 0.9:
        issues.append("Values suspiciously all the same")

    if len(numeric_values) >= 5:
        mean = sum(numeric_values) / len(numeric_values)
        variance = sum((value - mean) ** 2 for value in numeric_values) / len(numeric_values)
        stddev = math.sqrt(variance)
        if stddev > 0:
            outlier_count = sum(1 for value in numeric_values if abs(value - mean) > 3 * stddev)
            if outlier_count:
                issues.append("Outliers detected (values outside 3 standard deviations)")

    pii_types = {"email address", "phone number", "name"}
    pii_hits = sum(1 for kind in atomic_types if kind in pii_types)
    if detected_type in pii_types or (non_null_count and pii_hits / non_null_count >= 0.4):
        issues.append("Possible PII detected (emails/phones/names)")

    return issues


def analyse_column(series: pd.Series) -> dict[str, Any]:
    raw_values = list(series)
    total_count = len(raw_values)
    non_null_values = [value for value in raw_values if not is_effective_null(value)]
    non_null_texts = [stringify(value).strip() for value in non_null_values]
    null_count = total_count - len(non_null_values)
    null_percentage = round((null_count / total_count) * 100, 2) if total_count else 0.0

    atomic_types: list[str] = []
    atomic_counts = Counter()
    date_labels: list[str] = []
    numeric_values: list[float] = []
    date_values: list[datetime] = []

    for value in non_null_values:
        atomic_type = detect_atomic_type(value)
        atomic_types.append(atomic_type)
        atomic_counts[atomic_type] += 1

        parsed_date = maybe_parse_date(value)
        if parsed_date:
            date_value, date_label = parsed_date
            date_values.append(date_value)
            date_labels.append(date_label)

        parsed_number = maybe_parse_percentage(value)
        if parsed_number is None:
            parsed_number = maybe_parse_number(value)
        if parsed_number is not None:
            numeric_values.append(parsed_number)

    detected_type = infer_column_type(str(series.name), series, atomic_counts, non_null_texts)

    value_counts = Counter(non_null_texts)
    unique_count = len(value_counts)
    unique_percentage = round((unique_count / len(non_null_texts)) * 100, 2) if non_null_texts else 0.0

    mixed_types = False
    material_types = {
        kind for kind, count in atomic_counts.items()
        if kind != "unknown" and count >= max(2, math.ceil(len(non_null_values) * 0.2))
    }
    if len(material_types) > 1:
        mixed_types = True

    min_value = None
    max_value = None
    if detected_type in {"currency/amount", "plain number", "percentage"}:
        min_value, max_value = compute_numeric_range(numeric_values)
    elif detected_type == "date":
        min_value, max_value = compute_date_range(date_values)

    most_common_values = [
        {"value": value, "count": count}
        for value, count in value_counts.most_common(5)
    ]

    return {
        "detected_type": detected_type,
        "type_scores": {
            kind: round((atomic_counts[kind] / len(non_null_values)) * 100, 2)
            for kind in TYPE_ORDER
            if len(non_null_values) and atomic_counts[kind]
        },
        "null_count": null_count,
        "null_percentage": null_percentage,
        "unique_count": unique_count,
        "unique_percentage": unique_percentage,
        "most_common_values": most_common_values,
        "min_value": min_value,
        "max_value": max_value,
        "sample_values": sample_examples(non_null_texts, 3),
        "has_mixed_types": mixed_types,
        "suspected_issues": detect_suspected_issues(
            series=series,
            detected_type=detected_type,
            non_null_texts=non_null_texts,
            atomic_types=atomic_types,
            date_labels=date_labels,
            numeric_values=numeric_values,
        ),
    }


def analyse_dataframe(df: pd.DataFrame) -> dict[str, Any]:
    columns: dict[str, Any] = {}
    detected_counter = Counter()
    issue_counter = Counter()

    for column in df.columns:
        column_report = analyse_column(df[column])
        columns[str(column)] = column_report
        detected_counter[column_report["detected_type"]] += 1
        for issue in column_report["suspected_issues"]:
            issue_counter[issue] += 1

    return {
        "columns": columns,
        "summary": {
            "total_rows": int(len(df)),
            "total_columns": int(len(df.columns)),
            "detected_types": dict(sorted(detected_counter.items())),
            "issue_counts": dict(sorted(issue_counter.items())),
        },
    }


def build_report(file_path: Path) -> dict[str, Any]:
    loaded = load_file(file_path)
    df = loaded["dataframe"]
    analysis = analyse_dataframe(df)
    return {
        "file": str(file_path),
        "detected_format": loaded["detected_format"],
        "detected_encoding": loaded["detected_encoding"],
        "delimiter": loaded["delimiter"],
        "original_rows": loaded["original_rows"],
        "original_columns": loaded["original_columns"],
        "warnings": loaded["warnings"],
        **analysis,
    }


def main() -> int:
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: column_detector.py <file>"}))
        return 1

    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(json.dumps({"error": f"File not found: {file_path}"}))
        return 1

    try:
        report = build_report(file_path)
    except Exception as exc:
        print(json.dumps({"error": str(exc)}))
        return 1

    print(json.dumps(report, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
