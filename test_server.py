from __future__ import annotations

import json
import unittest
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
from io import BytesIO
from threading import Thread
from typing import Any
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from server import MAX_XLSX_BYTES, Handler, analyze_payload, parse_xlsx_bytes


def payload(
    headers: list[str],
    rows: list[list[str]],
    outcome: str = "Attrition",
    departed_value: str = "Yes",
    group_by: str | None = "Department",
) -> dict[str, Any]:
    return {
        "headers": headers,
        "rows": rows,
        "outcome": outcome,
        "departed_value": departed_value,
        "group_by": group_by,
    }


def sample_xlsx() -> bytes:
    data_rows = [("Sales", "No"), ("Sales", "No"), ("Sales", "Yes"), ("Sales", "No"), ("Sales", "Yes"), ("Tiny", "Yes")]
    worksheet_rows = ['<row r="1"><c r="A1" t="inlineStr"><is><t>Department</t></is></c><c r="B1" t="inlineStr"><is><t>Attrition</t></is></c></row>']
    worksheet_rows.extend(
        f'<row r="{index}"><c r="A{index}" t="inlineStr"><is><t>{department}</t></is></c><c r="B{index}" t="inlineStr"><is><t>{outcome}</t></is></c></row>'
        for index, (department, outcome) in enumerate(data_rows, start=2)
    )
    parts = {
        "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + "".join(worksheet_rows) + "</sheetData></worksheet>",
    }
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as workbook:
        for name, content in parts.items():
            workbook.writestr(name, content)
    return stream.getvalue()


class ExcelParserTests(unittest.TestCase):
    def test_xlsx_upload_limit_is_80_mb(self) -> None:
        self.assertEqual(MAX_XLSX_BYTES, 80_000_000)

    def test_reads_first_sheet_cells(self) -> None:
        self.assertEqual(
            parse_xlsx_bytes(sample_xlsx()),
            [["Department", "Attrition"], ["Sales", "No"], ["Sales", "No"], ["Sales", "Yes"], ["Sales", "No"], ["Sales", "Yes"], ["Tiny", "Yes"]],
        )

    def test_rejects_invalid_workbook(self) -> None:
        with self.assertRaisesRegex(ValueError, "invalid or uses an unsupported"):
            parse_xlsx_bytes(b"not an Excel workbook")


class AnalyzePayloadTests(unittest.TestCase):
    def test_aggregate_summary_suppresses_small_groups_and_identifiers(self) -> None:
        headers = ["EmployeeID", "Department", "Attrition"]
        rows = [
            [f"E{index}", "Core" if index < 5 else "Tiny", "Yes" if index in (2, 4, 5) else "No"]
            for index in range(6)
        ]

        result = analyze_payload(payload(headers, rows))

        self.assertEqual(result["records"], 6)
        self.assertEqual(result["departures"], 3)
        self.assertEqual(result["attrition_rate"], 0.5)
        self.assertEqual(len(result["groups"]), 1)
        self.assertEqual(result["groups"][0]["label"], "Core")
        self.assertEqual(result["groups"][0]["records"], 5)
        self.assertEqual(result["groups"][0]["departures"], 2)
        self.assertEqual(result["groups"][0]["attrition_rate"], 0.4)
        lower, upper = result["groups"][0]["confidence_interval"]
        self.assertLessEqual(lower, 0.4)
        self.assertGreaterEqual(upper, 0.4)
        self.assertLess(result["confidence_interval"][0], result["confidence_interval"][1])
        self.assertEqual(result["suppressed_groups"], 1)
        serialized = json.dumps(result)
        self.assertNotIn("E0", serialized)
        self.assertNotIn('"rows"', serialized)

    def test_blank_outcomes_are_excluded_from_rates_and_group_counts(self) -> None:
        headers = ["Department", "Attrition"]
        rows = [["Core", "Yes"], ["Core", "No"], ["Core", "Yes"], ["Core", "No"], ["Core", "No"], ["Core", ""]]

        result = analyze_payload(payload(headers, rows))

        self.assertEqual(result["records"], 5)
        self.assertEqual(result["excluded_outcome_rows"], 1)
        self.assertEqual(result["attrition_rate"], 0.4)
        self.assertEqual(result["groups"][0]["records"], 5)

    def test_small_trend_periods_are_suppressed(self) -> None:
        headers = ["Year", "Department", "Attrition"]
        rows = [["2019", "Core", "Yes" if index < 2 else "No"] for index in range(5)]
        rows += [["2020", "Core", "Yes" if index < 2 else "No"] for index in range(4)]

        result = analyze_payload(payload(headers, rows))

        self.assertEqual(result["trend_column"], "Year")
        self.assertEqual(len(result["trend"]), 1)
        period = result["trend"][0]
        self.assertEqual((period["period"], period["records"], period["departures"], period["attrition_rate"]), ("2019", 5, 2, 0.4))
        self.assertLessEqual(period["confidence_interval"][0], 0.4)
        self.assertGreaterEqual(period["confidence_interval"][1], 0.4)
        self.assertEqual(result["suppressed_periods"], 1)

    def test_numeric_factor_is_returned_as_privacy_safe_bands(self) -> None:
        headers = ["Department", "MonthlyIncome", "Attrition"]
        rows = [["Core", str(value), "Yes" if value % 2 else "No"] for value in range(1, 26)]
        rows += [["Core", "Unknown", "No"] for _ in range(3)]

        result = analyze_payload(
            payload(headers, rows, group_by="MonthlyIncome")
            | {"group_mode": "quantile"}
        )

        self.assertEqual(len(result["groups"]), 5)
        self.assertTrue(all(group["records"] == 5 for group in result["groups"]))
        self.assertEqual(result["excluded_group_rows"], 3)
        self.assertNotIn("Unknown", json.dumps(result))

    def test_invalid_group_mode_is_rejected(self) -> None:
        headers = ["Department", "Attrition"]
        rows = [["Core", "Yes"], ["Core", "No"]]

        with self.assertRaisesRegex(ValueError, "supported group analysis mode"):
            analyze_payload(payload(headers, rows) | {"group_mode": []})

    def test_two_factor_cells_are_aggregated_and_small_cells_suppressed(self) -> None:
        rows = [["Sales", "Day", "Yes" if index < 2 else "No"] for index in range(5)]
        rows += [["Sales", "Night", "Yes" if index < 3 else "No"] for index in range(5)]
        rows += [["Support", "Day", "Yes"]]

        result = analyze_payload(
            payload(["Department", "Shift", "Attrition"], rows)
            | {"secondary_group_by": "Shift"}
        )

        self.assertEqual(len(result["cross_groups"]), 2)
        self.assertEqual(result["suppressed_cross_groups"], 1)
        self.assertTrue(all(group["records"] >= result["min_group_size"] for group in result["cross_groups"]))
        self.assertNotIn("rows", result)

    def test_maximum_supported_row_count_returns_aggregate_summaries(self) -> None:
        headers = ["Year", "Department", "Shift", "Attrition"]
        rows = [
            [str(2017 + index % 6), f"Dept {index % 4}", "Day" if (index // 4) % 2 else "Night", "Yes" if index % 7 == 0 else "No"]
            for index in range(100_000)
        ]

        result = analyze_payload(
            payload(headers, rows, group_by="Department")
            | {"secondary_group_by": "Shift"}
        )

        self.assertEqual(result["records"], 100_000)
        self.assertEqual(len(result["groups"]), 4)
        self.assertEqual(len(result["cross_groups"]), 8)
        self.assertEqual(len(result["trend"]), 6)
        self.assertNotIn("rows", result)

    def test_second_breakdown_requires_categorical_primary(self) -> None:
        headers = ["Income", "Department", "Attrition"]
        rows = [[str(index), "Sales", "Yes" if index % 2 else "No"] for index in range(10)]

        with self.assertRaisesRegex(ValueError, "categorical primary"):
            analyze_payload(
                payload(headers, rows, group_by="Income")
                | {"group_mode": "quantile", "secondary_group_by": "Department"}
            )

    def test_identifier_columns_cannot_be_outcomes_or_groups(self) -> None:
        headers = ["EmployeeID", "Department", "Attrition"]
        rows = [[f"E{index}", "Core", "Yes" if index == 0 else "No"] for index in range(6)]

        with self.assertRaisesRegex(ValueError, "identifier"):
            analyze_payload(payload(headers, rows, outcome="EmployeeID", departed_value="E0"))
        with self.assertRaisesRegex(ValueError, "Identifier"):
            analyze_payload(payload(headers, rows, group_by="EmployeeID"))

    def test_duplicate_headers_are_rejected_case_insensitively(self) -> None:
        with self.assertRaisesRegex(ValueError, "unique"):
            analyze_payload(payload(["Department", "department", "Attrition"], []))

    def test_group_identifier_cardinality_is_rejected(self) -> None:
        headers = ["Department", "Attrition"]
        rows = [[f"Team {index}", "Yes" if index % 2 else "No"] for index in range(6)]

        with self.assertRaisesRegex(ValueError, "lower-cardinality"):
            analyze_payload(payload(headers, rows))

    def test_outcome_requires_both_departure_and_stayer_values(self) -> None:
        headers = ["Department", "Attrition"]
        rows = [["Core", "Yes"] for _ in range(6)]

        with self.assertRaisesRegex(ValueError, "between 2 and 20 values"):
            analyze_payload(payload(headers, rows))

    def test_row_width_must_match_headers(self) -> None:
        with self.assertRaisesRegex(ValueError, "match the number of columns"):
            analyze_payload(payload(["Department", "Attrition"], [["Core"]]))


class LocalHTTPTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.thread = Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        cls.authority = f"127.0.0.1:{cls.httpd.server_address[1]}"

    @classmethod
    def tearDownClass(cls) -> None:
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=2)

    def request(
        self,
        method: str,
        path: str,
        body: str | bytes | None = None,
        host: str | None = None,
        origin: str | None = None,
        content_type: str = "application/json",
    ) -> tuple[int, dict[str, str], bytes]:
        connection = HTTPConnection("127.0.0.1", self.httpd.server_address[1])
        headers = {"Host": host or self.authority}
        if origin is not None:
            headers["Origin"] = origin
        if body is not None:
            headers["Content-Type"] = content_type
        connection.request(method, path, body=body, headers=headers)
        response = connection.getresponse()
        result = response.status, {key.lower(): value for key, value in response.getheaders()}, response.read()
        connection.close()
        return result

    def test_health_is_local_and_no_store(self) -> None:
        status, headers, body = self.request("GET", "/api/health")

        self.assertEqual(status, 200)
        self.assertEqual(json.loads(body), {"ok": True})
        self.assertEqual(headers["cache-control"], "no-store")
        self.assertEqual(headers["x-content-type-options"], "nosniff")

    def test_dashboard_is_not_cached_or_frameable(self) -> None:
        status, headers, body = self.request("GET", "/")

        self.assertEqual(status, 200)
        self.assertEqual(headers["cache-control"], "no-store")
        self.assertIn("frame-ancestors 'none'", headers["content-security-policy"])
        self.assertIn(b"Employee Attrition Analytics", body)

    def test_foreign_host_and_origin_are_rejected(self) -> None:
        self.assertEqual(self.request("GET", "/api/health", host="evil.example")[0], 403)
        self.assertEqual(
            self.request("GET", "/api/health", origin="http://evil.example")[0],
            403,
        )

    def test_analysis_endpoint_returns_aggregate_data_only(self) -> None:
        headers = ["EmployeeID", "Department", "Attrition"]
        rows = [[f"E{index}", "Core", "Yes" if index < 2 else "No"] for index in range(6)]
        status, _, body = self.request(
            "POST",
            "/api/analyze",
            json.dumps(payload(headers, rows)),
            origin=f"http://{self.authority}",
        )

        self.assertEqual(status, 200)
        result = json.loads(body)
        self.assertEqual(result["records"], 6)
        self.assertNotIn("rows", result)
        self.assertNotIn("E0", body.decode())

    def test_malformed_json_is_rejected(self) -> None:
        status, _, _ = self.request("POST", "/api/analyze", "not-json")

        self.assertEqual(status, 400)

    def test_non_json_analysis_request_is_rejected(self) -> None:
        status, _, _ = self.request(
            "POST",
            "/api/analyze",
            "headers=Department",
            content_type="text/plain",
        )

        self.assertEqual(status, 415)

    def test_xlsx_conversion_endpoint_returns_a_matrix(self) -> None:
        status, _, body = self.request(
            "POST",
            "/api/parse-xlsx",
            sample_xlsx(),
            origin=f"http://{self.authority}",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        self.assertEqual(status, 200)
        self.assertEqual(json.loads(body), {"matrix": [["Department", "Attrition"], ["Sales", "No"], ["Sales", "No"], ["Sales", "Yes"], ["Sales", "No"], ["Sales", "Yes"], ["Tiny", "Yes"]]})

    def test_xlsx_conversion_endpoint_rejects_files_over_the_limit(self) -> None:
        workbook = sample_xlsx()
        with patch("server.MAX_XLSX_BYTES", len(workbook) - 1):
            status, _, body = self.request(
                "POST",
                "/api/parse-xlsx",
                workbook,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        self.assertEqual(status, 413)
        self.assertEqual(json.loads(body), {"error": "Excel files must be 80 MB or smaller."})


if __name__ == "__main__":
    unittest.main()
