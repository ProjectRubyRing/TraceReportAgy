#!/usr/bin/env python3
"""
csv_transformer.py
Transform CloudWatch Logs Insights query result JSON into a structured CSV file.
Designed for RHEL 9.8 / Python 3.9+.
"""

import sys
import json
import csv
import argparse
from typing import List, Dict, Any


COLUMNS = [
    "timestamp",
    "trace_id",
    "span_id",
    "parent_span_id",
    "service_name",
    "operation_name",
    "duration_ms",
    "duration_nano",
    "http_status",
    "http_method",
    "http_url",
    "status_code",
    "status_message",
    "kind",
]


def extract_field_dict(row: List[Dict[str, str]]) -> Dict[str, str]:
    """Convert CloudWatch Logs Insights row [{'field': k, 'value': v}] to dict."""
    d = {}
    for item in row:
        field = item.get("field", "")
        val = item.get("value", "")
        d[field] = val
    return d


def normalize_record(raw: Dict[str, str]) -> Dict[str, Any]:
    """Normalize fields from raw dictionary into target schema."""
    # Timestamp
    ts = raw.get("@timestamp") or raw.get("timestamp", "")

    # Trace & Span IDs
    trace_id = raw.get("traceId") or raw.get("trace_id", "")
    span_id = raw.get("spanId") or raw.get("span_id", "")
    parent_span_id = raw.get("parentSpanId") or raw.get("parent_span_id", "")

    # Service & Operation
    service_name = (
        raw.get("service_name")
        or raw.get("attributes.aws.local.service")
        or raw.get("resource.attributes.service.name")
        or raw.get("service.name")
        or "UnknownService"
    )
    operation_name = (
        raw.get("operation_name")
        or raw.get("attributes.aws.local.operation")
        or raw.get("name")
        or "UnknownOperation"
    )

    # Duration
    duration_nano_str = raw.get("durationNano") or raw.get("duration_nano", "")
    duration_ms_val = ""
    duration_nano_val = ""

    if duration_nano_str:
        try:
            nano = float(duration_nano_str)
            duration_nano_val = int(nano)
            duration_ms_val = round(nano / 1_000_000.0, 3)
        except ValueError:
            duration_nano_val = duration_nano_str
    elif raw.get("duration_ms"):
        try:
            ms = float(raw["duration_ms"])
            duration_ms_val = round(ms, 3)
            duration_nano_val = int(ms * 1_000_000)
        except ValueError:
            duration_ms_val = raw["duration_ms"]
    elif raw.get("duration"):
        try:
            d_val = float(raw["duration"])
            # If duration is in seconds (e.g. 0.045)
            if d_val < 1000.0:
                duration_ms_val = round(d_val * 1000.0, 3)
                duration_nano_val = int(d_val * 1_000_000_000)
            else:
                duration_ms_val = round(d_val, 3)
                duration_nano_val = int(d_val * 1_000_000)
        except ValueError:
            pass

    # HTTP & Status
    http_status = (
        raw.get("http_status")
        or raw.get("attributes.http.response.status_code")
        or raw.get("http.status_code", "")
    )
    http_method = (
        raw.get("http_method")
        or raw.get("attributes.http.method")
        or raw.get("http.method", "")
    )
    http_url = (
        raw.get("http_url")
        or raw.get("attributes.http.target")
        or raw.get("attributes.url.full")
        or raw.get("http.url", "")
    )
    status_code = (
        raw.get("status_code")
        or raw.get("status.code", "")
    )
    status_message = (
        raw.get("status_message")
        or raw.get("status.message")
        or raw.get("attributes.error.message", "")
    )
    kind = raw.get("kind", "")

    # Derive status_code if unset
    if not status_code:
        if http_status:
            try:
                code_num = int(http_status)
                if code_num >= 400:
                    status_code = "ERROR"
                else:
                    status_code = "OK"
            except ValueError:
                status_code = "UNSET"
        else:
            status_code = "OK"

    return {
        "timestamp": ts,
        "trace_id": trace_id,
        "span_id": span_id,
        "parent_span_id": parent_span_id,
        "service_name": service_name,
        "operation_name": operation_name,
        "duration_ms": duration_ms_val,
        "duration_nano": duration_nano_val,
        "http_status": http_status,
        "http_method": http_method,
        "http_url": http_url,
        "status_code": status_code,
        "status_message": status_message,
        "kind": kind,
    }


def transform_json_to_csv(input_path: str, output_path: str) -> int:
    """Parse CloudWatch query results JSON and write standard CSV."""
    if input_path == "-":
        data = json.load(sys.stdin)
    else:
        with open(input_path, "r", encoding="utf-8") as f:
            data = json.load(f)

    # CloudWatch Logs Insights result is inside "results" key
    results = data.get("results", [])
    if not results and isinstance(data, list):
        # Already a list of rows
        results = data

    normalized_rows = []
    for row in results:
        if isinstance(row, list):
            raw_dict = extract_field_dict(row)
        elif isinstance(row, dict):
            raw_dict = row
        else:
            continue
        normalized_rows.append(normalize_record(raw_dict))

    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(normalized_rows)

    return len(normalized_rows)


def main():
    parser = argparse.ArgumentParser(description="Transform CloudWatch Logs Insights JSON to CSV.")
    parser.add_argument("input", help="Path to input JSON file, or '-' for stdin")
    parser.add_argument("output", help="Path to output CSV file")
    args = parser.parse_args()

    count = transform_json_to_csv(args.input, args.output)
    print(f"Exported {count} span records to {args.output}")


if __name__ == "__main__":
    main()
