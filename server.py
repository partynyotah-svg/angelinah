from __future__ import annotations

import json
import math
import posixpath
import re
from io import BytesIO
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET


HOST = "127.0.0.1"
PORT = 8770
MIN_GROUP_SIZE = 5
MAX_BODY_BYTES = 20_000_000
MAX_XLSX_BYTES = 80_000_000
MAX_XLSX_UNCOMPRESSED_BYTES = 50_000_000
MAX_ROWS = 100_000
MAX_COLUMNS = 200
PAGE = Path(__file__).with_name("attritioniq.html")
CONTENT_SECURITY_POLICY = (
    "default-src 'self'; script-src 'self' 'unsafe-inline'; "
    "style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; "
    "base-uri 'self'; form-action 'self'; frame-ancestors 'none'"
)
XLSX_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XLSX_DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XLSX_PACKAGE_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def wilson_interval(successes: int, total: int) -> list[float]:
    if total <= 0:
        return [0.0, 1.0]
    z = 1.959963984540054
    rate = successes / total
    scale = 1 + z * z / total
    center = (rate + z * z / (2 * total)) / scale
    margin = z * math.sqrt(rate * (1 - rate) / total + z * z / (4 * total * total)) / scale
    return [max(0.0, center - margin), min(1.0, center + margin)]


def parse_xlsx_bytes(data: bytes) -> list[list[str]]:
    if len(data) > MAX_XLSX_BYTES:
        raise ValueError("Excel files must be 8 MB or smaller.")
    try:
        with ZipFile(BytesIO(data)) as workbook:
            entries = workbook.infolist()
            if sum(entry.file_size for entry in entries) > MAX_XLSX_UNCOMPRESSED_BYTES:
                raise ValueError("The expanded Excel workbook exceeds the 50 MB processing limit.")
            names = {entry.filename for entry in entries}
            book_root = ET.fromstring(workbook.read("xl/workbook.xml"))
            relations_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
            first_sheet = book_root.find(f"{XLSX_MAIN_NS}sheets/{XLSX_MAIN_NS}sheet")
            if first_sheet is None:
                raise ValueError("The Excel workbook has no worksheets.")
            relation_id = first_sheet.get(f"{XLSX_DOC_REL_NS}id")
            relations = {
                relation.get("Id"): relation.get("Target")
                for relation in relations_root.findall(f"{XLSX_PACKAGE_REL_NS}Relationship")
            }
            target = relations.get(relation_id or "")
            if not target:
                raise ValueError("The first worksheet could not be found.")
            sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            if not sheet_path.startswith("xl/") or sheet_path not in names:
                raise ValueError("The first worksheet path is invalid.")

            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in names:
                shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
                shared_strings = [
                    "".join(text.text or "" for text in item.findall(f".//{XLSX_MAIN_NS}t"))
                    for item in shared_root.findall(f"{XLSX_MAIN_NS}si")
                ]

            sheet_root = ET.fromstring(workbook.read(sheet_path))
            sheet_rows = sheet_root.findall(f"{XLSX_MAIN_NS}sheetData/{XLSX_MAIN_NS}row")
            if len(sheet_rows) > MAX_ROWS + 1:
                raise ValueError("The Excel worksheet exceeds the 100,000-row limit.")
            matrix: list[list[str]] = []
            for sheet_row in sheet_rows:
                cells: dict[int, str] = {}
                for cell in sheet_row.findall(f"{XLSX_MAIN_NS}c"):
                    reference = re.match(r"[A-Z]+", cell.get("r", "").upper())
                    if not reference:
                        continue
                    column = 0
                    for letter in reference.group():
                        column = column * 26 + ord(letter) - ord("A") + 1
                    column -= 1
                    if column >= MAX_COLUMNS:
                        raise ValueError("The Excel worksheet exceeds the 200-column limit.")
                    cell_type = cell.get("t", "")
                    value_node = cell.find(f"{XLSX_MAIN_NS}v")
                    if cell_type == "inlineStr":
                        inline = cell.find(f"{XLSX_MAIN_NS}is")
                        value = "".join(text.text or "" for text in inline.findall(f".//{XLSX_MAIN_NS}t")) if inline is not None else ""
                    elif value_node is None:
                        value = ""
                    elif cell_type == "s":
                        value = shared_strings[int(value_node.text or "0")]
                    elif cell_type == "b":
                        value = "TRUE" if value_node.text == "1" else "FALSE"
                    else:
                        value = value_node.text or ""
                    cells[column] = value
                row_width = max(cells, default=-1) + 1
                matrix.append([cells.get(index, "") for index in range(row_width)])
            return matrix
    except ValueError:
        raise
    except (BadZipFile, ET.ParseError, KeyError, IndexError, RuntimeError) as error:
        raise ValueError("The Excel workbook is invalid or uses an unsupported layout.") from error


def analyze_payload(payload: dict[str, Any]) -> dict[str, Any]:
    headers = payload.get("headers")
    rows = payload.get("rows")
    outcome = payload.get("outcome")
    departed_value = payload.get("departed_value")
    group_by = payload.get("group_by")
    group_mode = payload.get("group_mode", "category")
    secondary_group_by = payload.get("secondary_group_by")

    if not isinstance(headers, list) or not headers or len(headers) > MAX_COLUMNS:
        raise ValueError("Provide between 1 and 200 column headers.")
    headers = [str(value).strip() for value in headers]
    if any(not value for value in headers):
        raise ValueError("Column headers cannot be blank.")
    if len({value.casefold() for value in headers}) != len(headers):
        raise ValueError("Column headers must be unique.")
    if not isinstance(rows, list) or len(rows) > MAX_ROWS:
        raise ValueError("Provide no more than 100,000 data rows.")
    if outcome not in headers:
        raise ValueError("Choose an outcome column from the uploaded file.")
    if not isinstance(departed_value, str) or not departed_value.strip():
        raise ValueError("Choose the value that means an employee left.")
    if group_by is not None and group_by not in headers:
        raise ValueError("Choose a group column from the uploaded file.")
    if secondary_group_by is not None and (
        not isinstance(secondary_group_by, str) or secondary_group_by not in headers
    ):
        raise ValueError("Choose a second group column from the uploaded file.")
    if not isinstance(group_mode, str) or group_mode not in {"category", "quantile"}:
        raise ValueError("Choose a supported group analysis mode.")
    if group_mode == "quantile" and group_by is None:
        raise ValueError("Choose a numeric column to create bands.")
    if secondary_group_by is not None and (group_by is None or group_mode != "category"):
        raise ValueError("A second breakdown requires a categorical primary group.")

    outcome_index = headers.index(outcome)
    group_index = headers.index(group_by) if group_by else None
    secondary_index = headers.index(secondary_group_by) if secondary_group_by else None
    identifier_name = re.compile(
        r"\b(id|name|email|phone|address|ssn)\b|(?:employee|staff|person)[\s_-]?(?:id|number)",
        re.I,
    )
    if identifier_name.search(outcome):
        raise ValueError("An identifier column cannot be used as the attrition outcome.")
    if group_by and identifier_name.search(group_by):
        raise ValueError("Identifier columns cannot be used to group results.")
    if secondary_group_by and (
        secondary_group_by == outcome
        or secondary_group_by == group_by
        or identifier_name.search(secondary_group_by)
    ):
        raise ValueError("Choose a distinct, non-identifier second group column.")
    departure_value = departed_value.strip().casefold()
    clean_rows: list[list[str]] = []
    for row in rows:
        if not isinstance(row, list) or len(row) != len(headers):
            raise ValueError("Each data row must match the number of columns.")
        values = [str(value if value is not None else "").strip() for value in row]
        if any(len(value) > 10_000 for value in values):
            raise ValueError("A cell exceeds the 10,000 character limit.")
        if any(values):
            clean_rows.append(values)

    if not clean_rows:
        raise ValueError("The uploaded file has no data rows.")
    analyzed_rows = [row for row in clean_rows if row[outcome_index]]
    if not analyzed_rows:
        raise ValueError("The outcome column has no nonblank values.")
    if not any(row[outcome_index].casefold() == departure_value for row in analyzed_rows):
        raise ValueError("The selected departure value does not appear in the outcome column.")
    outcome_values = {row[outcome_index].casefold() for row in analyzed_rows}
    if len(outcome_values) < 2 or len(outcome_values) > 20:
        raise ValueError("Choose a categorical outcome with between 2 and 20 values.")
    if group_index is not None and group_mode == "category":
        group_values = {row[group_index] for row in analyzed_rows if row[group_index]}
        if len(group_values) > 30 or len(group_values) / len(analyzed_rows) > 0.5:
            raise ValueError("Choose a lower-cardinality group column; identifier-like groups are not allowed.")
    if secondary_index is not None:
        secondary_values = {row[secondary_index] for row in analyzed_rows if row[secondary_index]}
        if len(secondary_values) > 30 or len(secondary_values) / len(analyzed_rows) > 0.5:
            raise ValueError("Choose a lower-cardinality second group column.")

    departures = sum(row[outcome_index].casefold() == departure_value for row in analyzed_rows)
    result: dict[str, Any] = {
        "records": len(analyzed_rows),
        "excluded_outcome_rows": len(clean_rows) - len(analyzed_rows),
        "departures": departures,
        "attrition_rate": departures / len(analyzed_rows),
        "confidence_interval": wilson_interval(departures, len(analyzed_rows)),
        "min_group_size": MIN_GROUP_SIZE,
        "group_by": group_by,
        "group_mode": group_mode,
        "secondary_group_by": secondary_group_by,
        "cross_groups": [],
        "suppressed_cross_groups": 0,
        "excluded_group_rows": 0,
        "groups": [],
        "suppressed_groups": 0,
        "missing_by_column": [
            {
                "column": header,
                "missing": sum(not row[index] for row in analyzed_rows),
            }
            for index, header in enumerate(headers)
        ],
    }

    period_index = next(
        (
            index
            for index, header in enumerate(headers)
            if re.search(r"\b(year|date|period|month)\b", header, re.I)
        ),
        None,
    )
    trend_counts: dict[str, list[int]] = {}
    if period_index is not None:
        for row in analyzed_rows:
            match = re.search(r"\b(?:19|20)\d{2}\b", row[period_index])
            if match:
                period = match.group()
                counts = trend_counts.setdefault(period, [0, 0])
                counts[0] += 1
                counts[1] += row[outcome_index].casefold() == departure_value
    result["trend_column"] = headers[period_index] if period_index is not None else None
    result["trend"] = [
        {
            "period": period,
            "records": records,
            "departures": period_departures,
            "attrition_rate": period_departures / records,
            "confidence_interval": wilson_interval(period_departures, records),
        }
        for period, (records, period_departures) in sorted(trend_counts.items())
        if records >= MIN_GROUP_SIZE
    ]
    result["suppressed_periods"] = sum(
        records < MIN_GROUP_SIZE for records, _ in trend_counts.values()
    )

    if group_index is not None and group_mode == "quantile":
        numeric_rows: list[tuple[float, list[str]]] = []
        for row in analyzed_rows:
            try:
                value = float(row[group_index])
            except ValueError:
                continue
            if math.isfinite(value):
                numeric_rows.append((value, row))
        if len(numeric_rows) < MIN_GROUP_SIZE * 2:
            raise ValueError("The numeric column needs at least 10 valid values for comparison.")
        numeric_rows.sort(key=lambda item: item[0])
        band_count = min(5, len(numeric_rows) // MIN_GROUP_SIZE)
        band_values: dict[int, list[float | int]] = {}
        for offset, (value, row) in enumerate(numeric_rows):
            band = min(band_count - 1, offset * band_count // len(numeric_rows))
            counts = band_values.setdefault(band, [0, 0, value, value])
            counts[0] += 1
            counts[1] += row[outcome_index].casefold() == departure_value
            counts[2] = min(counts[2], value)
            counts[3] = max(counts[3], value)
        result["excluded_group_rows"] = len(analyzed_rows) - len(numeric_rows)
        for band, (records, band_departures, minimum, maximum) in sorted(band_values.items()):
            result["groups"].append(
                {
                    "label": f"Band {band + 1} ({minimum:g}–{maximum:g})",
                    "records": records,
                    "departures": band_departures,
                    "attrition_rate": band_departures / records,
                    "confidence_interval": wilson_interval(band_departures, records),
                }
            )
    elif group_index is not None:
        grouped: dict[str, list[int]] = {}
        for row in analyzed_rows:
            label = row[group_index] or "(Blank)"
            counts = grouped.setdefault(label, [0, 0])
            counts[0] += 1
            counts[1] += row[outcome_index].casefold() == departure_value

        for label, (records, group_departures) in sorted(
            grouped.items(), key=lambda item: (-item[1][0], item[0].casefold())
        ):
            if records < MIN_GROUP_SIZE:
                result["suppressed_groups"] += 1
                continue
            result["groups"].append(
                {
                    "label": label,
                    "records": records,
                    "departures": group_departures,
                    "attrition_rate": group_departures / records,
                    "confidence_interval": wilson_interval(group_departures, records),
                }
            )

    if group_index is not None and secondary_index is not None:
        cross_counts: dict[tuple[str, str], list[int]] = {}
        for row in analyzed_rows:
            key = (row[group_index] or "(Blank)", row[secondary_index] or "(Blank)")
            counts = cross_counts.setdefault(key, [0, 0])
            counts[0] += 1
            counts[1] += row[outcome_index].casefold() == departure_value
        for (primary, secondary), (records, cross_departures) in sorted(
            cross_counts.items(), key=lambda item: (-item[1][0], item[0][0].casefold(), item[0][1].casefold())
        ):
            if records < MIN_GROUP_SIZE:
                result["suppressed_cross_groups"] += 1
                continue
            result["cross_groups"].append(
                {
                    "primary": primary,
                    "secondary": secondary,
                    "label": f"{primary} / {secondary}",
                    "records": records,
                    "departures": cross_departures,
                    "attrition_rate": cross_departures / records,
                    "confidence_interval": wilson_interval(cross_departures, records),
                }
            )

    return result


class Handler(BaseHTTPRequestHandler):
    def local_request(self) -> bool:
        host_header = self.headers.get("Host", "")
        origin = self.headers.get("Origin")
        try:
            host = urlsplit(f"//{host_header}")
            if host.hostname not in {"127.0.0.1", "localhost"}:
                return False
            if origin is None:
                return True
            parsed_origin = urlsplit(origin)
        except ValueError:
            return False
        return (
            parsed_origin.scheme == "http"
            and parsed_origin.netloc.casefold() == host_header.casefold()
        )

    def send_json(self, status: int, data: dict[str, Any]) -> None:
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if not self.local_request():
            self.send_error(403, "Requests must come from this local application.")
            return
        if self.path == "/api/health":
            self.send_json(200, {"ok": True})
            return
        if self.path not in ("/", "/attritioniq.html"):
            self.send_error(404)
            return
        try:
            body = PAGE.read_bytes()
        except OSError:
            self.send_error(500, "The AttritionIQ page could not be read.")
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", CONTENT_SECURITY_POLICY)
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        if not self.local_request():
            self.send_error(403, "Requests must come from this local application.")
            return
        if self.path == "/api/parse-xlsx":
            content_type = self.headers.get("Content-Type", "").split(";", 1)[0].strip().casefold()
            if content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                self.send_json(415, {"error": "Excel uploads must use the .xlsx workbook format."})
                return
            try:
                size = int(self.headers.get("Content-Length", "0"))
                if size <= 0 or size > MAX_XLSX_BYTES:
                    self.send_json(413, {"error": "Excel files must be 8 MB or smaller."})
                    return
                self.send_json(200, {"matrix": parse_xlsx_bytes(self.rfile.read(size))})
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
            return
        if self.path != "/api/analyze":
            self.send_error(404)
            return
        try:
            content_type = self.headers.get("Content-Type", "").split(";", 1)[0].strip().casefold()
            if content_type != "application/json":
                self.send_json(415, {"error": "Analysis requests must use application/json."})
                return
            size = int(self.headers.get("Content-Length", "0"))
            if size <= 0 or size > MAX_BODY_BYTES:
                self.send_json(413, {"error": "Request must be under 20 MB."})
                return
            payload = json.loads(self.rfile.read(size))
            if not isinstance(payload, dict):
                raise ValueError("Request body must be a JSON object.")
            self.send_json(200, analyze_payload(payload))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self.send_json(400, {"error": "Request body must be valid UTF-8 JSON."})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})

    def log_message(self, _format: str, *_args: Any) -> None:
        return


if __name__ == "__main__":
    if not PAGE.is_file():
        raise SystemExit(f"Place attritioniq.html beside {Path(__file__).name}.")
    with ThreadingHTTPServer((HOST, PORT), Handler) as httpd:
        httpd.daemon_threads = True
        print(f"AttritionIQ is available at http://{HOST}:{PORT}/")
        httpd.serve_forever()