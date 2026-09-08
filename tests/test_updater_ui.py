"""Test worker-to-Tk handoff without displaying dialogs or launching executables."""

from contextlib import ExitStack
from pathlib import Path
from queue import Queue
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch
from urllib.error import URLError

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src-pyloid"))
import updater  # noqa: E402


class UpdaterFlowTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(patch.object(updater.core.LOGGER, "disabled", True))
        temporary = self.stack.enter_context(tempfile.TemporaryDirectory(
            prefix="updater-ui-test-", dir=Path(__file__).resolve().parents[1]))
        self.installation = Path(temporary)
        self.stack.enter_context(patch.object(updater, "ROOT", self.installation))
        self.app = updater.Updater.__new__(updater.Updater)
        self.app.root = Mock()
        self.app.events = Queue()
        self.app.closed = False
        self.app.exit_code = 0
        self.app.worker = None
        self.app.fatal_ui_error = False
        self.app.status_var = Mock()
        self.app.progress_var = Mock()
        self.running = self.mock_core("is_main_running", return_value=False)
        self.recover = self.mock_core("recover_pending", return_value=[])
        self.version = self.mock_core("read_local_version", return_value="1.0.0")
        self.fetch = self.mock_core("fetch_latest_zip_asset", return_value=("1.0.0", {}, None))
        self.launch = self.mock_core("launch_main", return_value=Mock())
        self.dialog = self.stack.enter_context(patch.object(updater, "show_dialog"))

    def mock_core(self, name, **kwargs):
        return self.stack.enter_context(patch.object(updater.core, name, **kwargs))

    def terminal_event(self):
        events = []
        while not self.app.events.empty():
            event = self.app.events.get_nowait()
            if event[0] != "status":
                events.append(event)
        self.assertEqual(len(events), 1)
        return events[0]

    def prepare_update(self):
        self.fetch.return_value = ("2.0.0", {}, None)
        self.mock_core("download_asset", return_value=self.installation / "update.zip")
        self.mock_core("safe_extract_zip")
        self.mock_core("validate_payload")
        transaction = Mock()
        transaction.record = {"phase": "pending"}
        transaction.directory = self.installation / ".update_tmp/backup/example"
        transaction.commit.return_value = None
        self.mock_core("Transaction", return_value=transaction)
        return transaction

    def test_network_failure_waits_for_dialog_acknowledgement_before_fallback(self):
        self.fetch.side_effect = URLError("offline")
        self.app.run_update_flow()
        self.launch.assert_not_called()
        self.dialog.assert_not_called()  # Worker cannot call Tk.
        order = []
        self.dialog.side_effect = lambda *args: order.append("dialog")
        with patch.object(self.app, "_start_worker", side_effect=lambda target: order.append("launch")):
            self.app._drain_events()
        self.assertEqual(order, ["dialog", "launch"])
        self.assertIn("서버에 연결", self.dialog.call_args.args[3])

    def test_version_read_permission_error_is_reported_and_can_fallback(self):
        self.version.side_effect = PermissionError("version is locked")
        self.app.run_update_flow()
        kind, exc, can_launch = self.terminal_event()
        self.assertEqual(kind, "update_failed")
        self.assertTrue(can_launch)
        self.assertIsInstance(exc, PermissionError)
        self.fetch.assert_not_called()

    def test_process_inspection_failure_never_starts_install_or_launch(self):
        self.running.side_effect = PermissionError("cannot inspect")
        self.app.run_update_flow()
        self.assertFalse(self.terminal_event()[2])
        self.recover.assert_not_called()
        self.launch.assert_not_called()

    def test_interrupted_recovery_failure_disables_fallback(self):
        self.recover.side_effect = updater.core.InstallError(
            "restore blocked", rollback_ok=False, backup_dir=self.installation / "backup")
        self.app.run_update_flow()
        self.app._drain_events()
        self.assertEqual(self.app.exit_code, 1)
        self.assertTrue(self.app.closed)
        self.assertIn("백업 위치", self.dialog.call_args.args[3])
        self.launch.assert_not_called()

    def test_already_running_main_is_explained_without_updating(self):
        self.running.return_value = True
        self.app.run_update_flow()
        self.app._drain_events()
        self.recover.assert_not_called()
        self.fetch.assert_not_called()
        self.launch.assert_not_called()
        self.assertEqual(self.dialog.call_args.args[1], "showinfo")

    def test_latest_version_launch_error_shows_error_once_without_retry(self):
        self.launch.side_effect = updater.core.UpdateError("early exit")
        self.app.run_update_flow()
        self.app._drain_events()
        self.launch.assert_called_once()
        self.assertEqual(self.dialog.call_count, 1)
        self.assertEqual(self.dialog.call_args.args[2], "tdm 실행 실패")
        self.assertEqual(self.app.exit_code, 1)

    def test_updated_main_launch_failure_rolls_back_before_fallback(self):
        transaction = self.prepare_update()
        self.launch.side_effect = updater.core.UpdateError("early exit")
        self.app.run_update_flow()
        transaction.rollback.assert_called_once()
        transaction.commit.assert_not_called()
        event = self.terminal_event()
        self.assertEqual(event[0], "update_failed")
        self.assertTrue(event[2])
        self.assertIn("복구했습니다", str(event[1]))

    def test_updated_main_rollback_failure_retains_backup_and_stops(self):
        transaction = self.prepare_update()
        self.launch.side_effect = updater.core.UpdateError("early exit")
        transaction.rollback.side_effect = PermissionError("restore locked")
        self.app.run_update_flow()
        event = self.terminal_event()
        self.assertFalse(event[2])
        self.assertFalse(event[1].rollback_ok)
        self.assertEqual(event[1].backup_dir, transaction.directory)

    def test_commit_failure_while_main_lives_must_not_roll_back(self):
        transaction = self.prepare_update()
        transaction.commit.side_effect = PermissionError("journal locked")
        self.app.run_update_flow()
        transaction.rollback.assert_not_called()
        event = self.terminal_event()
        self.assertEqual(event[0], "finished")
        self.assertIn("완료 기록", event[1][0])

    def test_process_started_during_download_stops_installation(self):
        transaction = self.prepare_update()
        self.running.side_effect = [False, True]
        self.app.run_update_flow()
        transaction.install.assert_not_called()
        self.launch.assert_not_called()
        self.assertFalse(self.terminal_event()[2])

    def test_successful_update_commits_after_launch(self):
        transaction = self.prepare_update()
        order = []
        self.launch.side_effect = lambda root: order.append("launch") or Mock()
        transaction.commit.side_effect = lambda: order.append("commit")
        self.app.run_update_flow()
        self.assertEqual(order, ["launch", "commit"])
        self.assertEqual(self.terminal_event(), ("finished", []))

    def test_failed_fallback_launch_is_reported_then_closes(self):
        self.launch.side_effect = OSError("Popen failed")
        self.app._launch_existing()
        self.app._drain_events()
        self.assertEqual(self.app.exit_code, 1)
        self.assertEqual(self.dialog.call_args.args[2], "tdm 실행 실패")

    def test_worker_start_failure_reports_error_and_closes(self):
        with patch.object(updater.threading.Thread, "start", side_effect=RuntimeError("cannot start")):
            self.app.start_update_thread()
        self.app._drain_events()
        self.assertTrue(self.app.closed)
        self.assertEqual(self.app.exit_code, 1)

    def test_ui_failure_waits_for_active_file_worker_before_closing(self):
        self.app.worker = Mock()
        self.app.worker.is_alive.return_value = True
        error = RuntimeError("UI callback failed")
        self.app._callback_error(type(error), error, None)
        self.assertFalse(self.app.closed)
        self.app.worker.is_alive.return_value = False
        self.app._drain_events()
        self.assertTrue(self.app.closed)
        self.assertEqual(self.app.exit_code, 1)

    def test_tk_startup_failure_uses_error_dialog_and_releases_lock(self):
        lock = self.mock_core("UpdaterLock").return_value
        self.mock_core("configure_logging")
        with patch.object(updater, "Updater", side_effect=updater.tk.TclError("Tk init failed")):
            self.assertEqual(updater.main(), 1)
        self.dialog.assert_called_once()
        lock.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()
