import copy
import unittest
from unittest.mock import Mock, patch

import requests

from onedrive_excel_logger import DEFAULT_COLUMNS, GRAPH, OneDriveExcelLogger


class TableAppendTests(unittest.TestCase):
    def setUp(self):
        auth = patch.object(OneDriveExcelLogger, "_init_auth")
        auth.start()
        self.addCleanup(auth.stop)
        token = patch.object(OneDriveExcelLogger, "_token", return_value="mock-token")
        token.start()
        self.addCleanup(token.stop)
        http = patch("onedrive_excel_logger.requests.request", side_effect=self.graph)
        self.request = http.start()
        self.addCleanup(http.stop)
        self.headers = list(DEFAULT_COLUMNS.values())
        self.table_id = "table-id"
        self.columns_url = (
            f"{GRAPH}/drives/drive-id/items/item-id/workbook"
            f"/worksheets/Sheet1/tables/table-id/columns"
        )
        self.rows = [["original run", 0, 3, 0.1, "old detector", "old spot"]]
        self.original_rows = copy.deepcopy(self.rows)
        self.posts = []
        self.post_error = None
        self.column_pages = None
        self.logger = self.make_logger()

    def make_logger(self, *args, **kwargs):
        logger = OneDriveExcelLogger(
            "https://example.invalid/workbook", "example-client-id", *args, **kwargs
        )
        self.addCleanup(logger._pool.shutdown, wait=True)
        return logger

    def graph(self, method, url, **kwargs):
        if method == "GET" and "/shares/" in url:
            data = {"parentReference": {"driveId": "drive-id"}, "id": "item-id"}
        elif method == "GET" and "/tables/" in url:
            if "/columns" in url:
                data = self.column_pages[url] if self.column_pages else {
                    "value": [
                        {"name": name, "index": index}
                        for index, name in reversed(list(enumerate(self.headers)))
                    ]
                }
            else:
                data = {"id": self.table_id, "name": "BeamtimeRuns"}
        elif method == "POST" and url.endswith("/rows/add"):
            payload = kwargs["json"]
            self.assertIsNone(payload["index"])
            self.assertEqual(len(payload["values"]), 1)
            self.assertEqual(len(payload["values"][0]), len(self.headers))
            self.posts.append(copy.deepcopy(payload))
            if self.post_error:
                raise self.post_error
            index = len(self.rows)
            self.rows.extend(copy.deepcopy(payload["values"]))
            data = {"value": [{"index": index, "values": payload["values"]}]}
        else:
            self.fail(f"Unexpected Graph request: {method} {url}")
        response = Mock(content=b"json")
        response.json.return_value = data
        return response

    def run_documents(self, uid, scan_id, **metadata):
        self.logger.start({
            "uid": uid, "scan_id": scan_id, "detectors": ["det"], **metadata
        })
        self.logger.descriptor({
            "uid": f"{uid}-primary", "name": "primary",
            "configuration": {"det": {"data": {"det_cam_acquire_time": 0.25}}},
        })
        self.logger.event({"descriptor": f"{uid}-primary"})
        self.logger.event_page({"descriptor": f"{uid}-primary", "seq_num": [2, 3]})
        self.logger.stop({})

    def test_two_runs_append_in_order_without_overwriting_existing_rows(self):
        self.run_documents("run-1", 1, beam_spot_size="5 x 5 um")
        self.run_documents("run-2", 2, integration_time=0.5, beam_spot_size="10 x 10 um")
        self.logger._pool.shutdown(wait=True)
        self.assertEqual(self.posts, [
            {"index": None, "values": [["run-1", 1, 3, 0.25, "det", "5 x 5 um"]]},
            {"index": None, "values": [["run-2", 2, 3, 0.5, "det", "10 x 10 um"]]},
        ])
        self.assertEqual(self.rows[:1], self.original_rows)
        self.assertEqual(self.rows[1:], [post["values"][0] for post in self.posts])
        calls = self.request.call_args_list
        self.assertEqual(sum(call.args[0] == "POST" for call in calls), 2)
        self.assertTrue(all(call.args[0] in ("GET", "POST") for call in calls))
        self.assertTrue(all("usedRange" not in call.args[1] for call in calls))
        self.assertTrue(all("/range(" not in call.args[1] for call in calls))

    def test_reordered_normalized_headers_and_extra_columns(self):
        self.headers = [
            "Notes", " beam   SPOT size ", " detector ", "SCANID",
            " Integration\tTime per Frame ", " uid ", "Number  of frames",
        ]
        self.logger._append_row({
            "uid": "run", "scan_id": 42, "num_frames": 9,
            "integration_time": 0.2, "detector": "det", "beam_spot_size": "spot",
        })
        self.assertEqual(self.posts[0]["values"], [["", "spot", "det", 42, 0.2, "run", 9]])

    def test_custom_columns(self):
        logger = self.make_logger(columns={"uid": "Run ID", "scan_id": "Scan"})
        self.headers = ["Scan", "Extra", "RUN  ID"]
        logger._append_row({"uid": "run", "scan_id": 4})
        self.assertEqual(self.posts[0]["values"], [[4, "", "run"]])

    def test_missing_required_header_fails_before_write_and_at_login(self):
        self.headers.remove("Detector")
        for operation in (self.logger.login, lambda: self.logger._append_row({"uid": "run"})):
            with self.subTest(operation=operation):
                with self.assertRaisesRegex(ValueError, "Required column 'Detector' missing"):
                    operation()
        self.assertEqual(self.posts, [])
        self.assertTrue(all(call.args[0] == "GET" for call in self.request.call_args_list))

    def test_ambiguous_table_header_fails_before_write(self):
        self.headers.append(" uid ")
        with self.assertRaisesRegex(ValueError, "Ambiguous column 'UID'"):
            self.logger._append_row({"uid": "run"})
        self.assertEqual(self.posts, [])

    def test_ambiguous_configured_mapping_fails_before_write(self):
        logger = self.make_logger(columns={"uid": "UID", "scan_id": " uid "})
        with self.assertRaisesRegex(ValueError, "Ambiguous column"):
            logger.login()
        self.assertEqual(self.posts, [])

    def test_login_resolves_table_without_writing(self):
        self.logger.login()
        self.assertEqual(self.posts, [])
        self.assertEqual(self.request.call_args_list[-1].args, ("GET", self.columns_url))

    def test_headers_are_revalidated_after_login(self):
        self.logger.login()
        self.headers.remove("UID")
        with self.assertRaisesRegex(ValueError, "Required column 'UID' missing"):
            self.logger._append_row({"uid": "run"})
        self.assertEqual(self.posts, [])

    def test_url_encoding_retains_worksheet_and_resolved_table_id(self):
        logger = self.make_logger(worksheet="Scan / #?", table_name="Runs / #?")
        logger._drive_item = ("drive / #?", "item / #?")
        self.table_id = "{table / #?}"
        logger._append_row({"uid": "run"})
        prefix = (
            f"{GRAPH}/drives/drive%20%2F%20%23%3F/items/item%20%2F%20%23%3F"
            "/workbook/worksheets/Scan%20%2F%20%23%3F/tables/"
        )
        self.assertEqual(self.request.call_args_list[0].args,
                         ("GET", prefix + "Runs%20%2F%20%23%3F"))
        self.assertEqual(self.request.call_args_list[-1].args,
                         ("POST", prefix + "%7Btable%20%2F%20%23%3F%7D/rows/add"))

    def test_paginated_columns_are_sorted_by_actual_index(self):
        self.column_pages = {
            self.columns_url: {
                "value": [{"name": name, "index": index}
                          for index, name in enumerate(self.headers[3:], start=3)],
                "@odata.nextLink": self.columns_url + "?page=2",
            },
            self.columns_url + "?page=2": {
                "value": [{"name": name, "index": index}
                          for index, name in enumerate(self.headers[:3])]
            },
        }
        self.logger._append_row({"uid": "run", "scan_id": 7})
        self.assertEqual(self.posts[0]["values"], [["run", 7, "", "", "", ""]])

    def test_append_timeout_is_logged_without_retry(self):
        self.post_error = requests.Timeout("Ambiguous append outcome")
        with self.assertLogs("onedrive_excel_logger", level="ERROR") as logs:
            self.logger._safe_append({"uid": "run"})
        self.assertEqual(len(self.posts), 1)
        self.assertIn("Failed to write run run", logs.output[0])
        self.assertEqual(self.rows, self.original_rows)

    def test_missing_table_does_not_create_or_fall_back(self):
        self.logger._drive_item = ("drive-id", "item-id")
        self.request.side_effect = requests.HTTPError("Table not found")
        with self.assertRaises(requests.HTTPError):
            self.logger.login()
        self.assertEqual(self.posts, [])
        self.assertIn("/tables/BeamtimeRuns", self.request.call_args.args[1])
        self.assertTrue(all(call.args[0] == "GET" for call in self.request.call_args_list))

    def test_returned_row_index_is_logged(self):
        with self.assertLogs("onedrive_excel_logger", level="INFO") as logs:
            self.logger._append_row({"uid": "run"})
        self.assertIn("returned row index: 1", logs.output[0])

    def test_original_positional_arguments_are_preserved(self):
        logger = self.make_logger(
            "organizations", None, "hint", "CustomSheet", {"uid": "UID"},
            ("exposure",), "spot", "secondary", "/tmp/unused-token-cache", True,
            table_name="CustomTable",
        )
        self.assertEqual(logger.worksheet, "CustomSheet")
        self.assertEqual(logger.table_name, "CustomTable")
        self.assertEqual(logger.stream_name, "secondary")
        self.assertTrue(logger.blocking)
        self.assertEqual(logger.beam_spot_size, "spot")

    def test_stop_event_count_and_spot_callable_are_preserved(self):
        self.logger.blocking = True
        self.logger.beam_spot_size = lambda: "callable spot"
        self.logger.start({"uid": "run", "scan_id": 4})
        self.logger.stop({"num_events": {"primary": 12}})
        self.assertEqual(self.posts[0]["values"], [["run", 4, 12, "", "", "callable spot"]])

    def test_unsupported_auth_and_empty_table_rejected_before_auth(self):
        for kwargs, message in [
            ({"client_secret": "unsupported"}, "application permissions"),
            ({"client_secret": ""}, "application permissions"),
            ({"tenant_id": "consumers"}, "personal Microsoft accounts"),
            ({"table_name": " "}, "existing Excel table"),
        ]:
            with self.subTest(kwargs=kwargs):
                with patch.object(OneDriveExcelLogger, "_init_auth") as auth:
                    with self.assertRaisesRegex(ValueError, message):
                        self.make_logger(**kwargs)
                    auth.assert_not_called()


if __name__ == "__main__":
    unittest.main()
