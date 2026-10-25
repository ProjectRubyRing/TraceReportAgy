#!/usr/bin/env python3
"""
test_csv_transformer.py
Unit tests for csv_transformer.py.
"""

import unittest
import tempfile
import os
import csv
import json
import sys

# Add scripts directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts")))
from csv_transformer import normalize_record, transform_json_to_csv, COLUMNS


class TestCSVTransformer(unittest.TestCase):
    def test_normalize_record_standard(self):
        raw = {
            "@timestamp": "2026-10-26 04:30:00.000",
            "traceId": "trace-123",
            "spanId": "span-001",
            "parentSpanId": "",
            "service_name": "order-service",
            "operation_name": "POST /order",
            "durationNano": "50000000",
            "http_status": "200",
            "kind": "SERVER",
        }
        res = normalize_record(raw)
        self.assertEqual(res["timestamp"], "2026-10-26 04:30:00.000")
        self.assertEqual(res["trace_id"], "trace-123")
        self.assertEqual(res["span_id"], "span-001")
        self.assertEqual(res["parent_span_id"], "")
        self.assertEqual(res["duration_ms"], 50.0)
        self.assertEqual(res["duration_nano"], 50000000)
        self.assertEqual(res["status_code"], "OK")

    def test_normalize_record_fallback_fields(self):
        raw = {
            "timestamp": "2026-10-26T04:30:00Z",
            "trace_id": "trace-456",
            "span_id": "span-002",
            "parent_span_id": "span-001",
            "attributes.aws.local.service": "payment-svc",
            "attributes.aws.local.operation": "Charge",
            "duration": "0.125",  # 125 ms in seconds
            "attributes.http.response.status_code": "500",
        }
        res = normalize_record(raw)
        self.assertEqual(res["service_name"], "payment-svc")
        self.assertEqual(res["operation_name"], "Charge")
        self.assertEqual(res["duration_ms"], 125.0)
        self.assertEqual(res["http_status"], "500")
        self.assertEqual(res["status_code"], "ERROR")

    def test_transform_json_to_csv(self):
        sample_json = {
            "results": [
                [
                    {"field": "@timestamp", "value": "2026-10-26 01:00:00"},
                    {"field": "traceId", "value": "t-1"},
                    {"field": "spanId", "value": "s-1"},
                    {"field": "parentSpanId", "value": ""},
                    {"field": "service_name", "value": "auth"},
                    {"field": "operation_name", "value": "login"},
                    {"field": "durationNano", "value": "20000000"},
                    {"field": "http_status", "value": "200"},
                ]
            ]
        }
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".json") as f_json:
            json.dump(sample_json, f_json)
            json_path = f_json.name

        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv") as f_csv:
            csv_path = f_csv.name

        try:
            count = transform_json_to_csv(json_path, csv_path)
            self.assertEqual(count, 1)

            with open(csv_path, "r", encoding="utf-8") as f:
                reader = list(csv.DictReader(f))
                self.assertEqual(len(reader), 1)
                self.assertEqual(reader[0]["trace_id"], "t-1")
                self.assertEqual(reader[0]["duration_ms"], "20.0")
        finally:
            if os.path.exists(json_path):
                os.remove(json_path)
            if os.path.exists(csv_path):
                os.remove(csv_path)


if __name__ == "__main__":
    unittest.main()
