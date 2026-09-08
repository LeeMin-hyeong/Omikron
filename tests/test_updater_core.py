"""Exercise updater failures with disposable files; never start an application."""

import hashlib
import io
import json
import stat
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

SOURCE_DIR = Path(__file__).resolve().parents[1] / "src-pyloid"
sys.path.insert(0, str(SOURCE_DIR))
import updater_core as core  # noqa: E402


class SimulatedInterruption(BaseException):
    """Model abrupt termination without invoking normal exception recovery."""


class TemporaryUpdaterTests(unittest.TestCase):
    def setUp(self):
        logger_patch = patch.object(core.LOGGER, "disabled", True)
        logger_patch.start()
        self.addCleanup(logger_patch.stop)
        workspace = Path(__file__).resolve().parents[1]
        self.temp = tempfile.TemporaryDirectory(prefix="updater-test-", dir=workspace)
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.root = self.base / "installed"
        self.new_root = self.base / "staged"
        self.write_payload(self.root, "old", "1.0.0")
        self.write_payload(self.new_root, "new", "2.0.0")
        (self.root / "license.json").write_text("{}", encoding="utf-8")
        (self.root / "user-data.txt").write_text("keep user data", encoding="utf-8")

    @staticmethod
    def write_payload(root, marker, version):
        files = {
            "main.exe": f"{marker} executable",
            "_internal/python312.dll": f"{marker} python",
            "_internal/dist-front/index.html": f"<html>{marker}</html>",
            "_internal/PySide6/QtWebEngineProcess.exe": f"{marker} renderer",
            "version.txt": version,
        }
        for relative, content in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content, encoding="utf-8")

    def assert_old_installation(self, version="1.0.0"):
        self.assertEqual((self.root / "main.exe").read_text(), "old executable")
        self.assertEqual((self.root / "_internal/python312.dll").read_text(), "old python")
        if version is None:
            self.assertFalse((self.root / "version.txt").exists())
        else:
            self.assertEqual((self.root / "version.txt").read_text().strip(), version)
        self.assertEqual((self.root / "license.json").read_text(), "{}")
        self.assertEqual((self.root / "user-data.txt").read_text(), "keep user data")

    def assert_new_installation(self):
        self.assertEqual((self.root / "main.exe").read_text(), "new executable")
        self.assertEqual((self.root / "_internal/python312.dll").read_text(), "new python")
        self.assertEqual((self.root / "version.txt").read_text().strip(), "2.0.0")
        self.assertEqual((self.root / "license.json").read_text(), "{}")
        self.assertEqual((self.root / "user-data.txt").read_text(), "keep user data")

    def transaction(self):
        return core.Transaction(self.root, self.new_root, "2.0.0")


class TransactionFailureTests(TemporaryUpdaterTests):
    def test_install_retains_backup_until_commit(self):
        transaction = self.transaction().install()
        self.assert_new_installation()
        self.assertTrue(transaction.directory.exists())
        transaction.commit()
        self.assert_new_installation()
        core.recover_pending(self.root)
        self.assert_new_installation()

    def test_explicit_rollback_restores_version_and_files(self):
        transaction = self.transaction().install()
        transaction.rollback()
        self.assert_old_installation()

    def test_rollback_preserves_originally_missing_version_file(self):
        (self.root / "version.txt").unlink()
        transaction = self.transaction().install()
        self.assert_new_installation()
        transaction.rollback()
        self.assert_old_installation(version=None)
        core.recover_pending(self.root)
        self.assert_old_installation(version=None)

    def test_commit_cleanup_failure_never_reverts_successful_update(self):
        transaction = self.transaction().install()
        with patch.object(core.shutil, "rmtree", side_effect=PermissionError("backup is locked")):
            warning = transaction.commit()
            self.assertIsInstance(warning, str)
            self.assertTrue(transaction.directory.exists())
            core.recover_pending(self.root)
            self.assert_new_installation()
        core.recover_pending(self.root)
        self.assert_new_installation()

    def test_commit_journal_failure_leaves_recoverable_backup(self):
        transaction = self.transaction().install()
        original_replace = Path.replace

        def fail_commit_record(source, destination):
            if source.name == "transaction.tmp":
                raise PermissionError("commit record is locked")
            return original_replace(source, destination)

        with patch.object(Path, "replace", fail_commit_record):
            with self.assertRaises(PermissionError):
                transaction.commit()
        self.assertTrue(transaction.directory.exists())
        core.recover_pending(self.root)
        self.assert_old_installation()

    def test_incomplete_payload_never_moves_original_files(self):
        (self.new_root / "_internal/dist-front/index.html").unlink()
        original_replace = Path.replace
        live_moves = []

        def track_replace(source, destination):
            if source.parent == self.root:
                live_moves.append(source)
            return original_replace(source, destination)

        with patch.object(Path, "replace", track_replace):
            with self.assertRaises(core.UpdateError):
                self.transaction().install()
        self.assertEqual(live_moves, [])
        self.assert_old_installation()

    def test_invalid_version_never_changes_original_installation(self):
        with self.assertRaises(core.UpdateError):
            core.Transaction(self.root, self.new_root, "not-a-version").install()
        self.assert_old_installation()

    def test_partial_backup_failure_preserves_unmoved_internal_directory(self):
        original_replace = Path.replace

        def fail_second_backup(source, destination):
            if source == self.root / "_internal":
                raise PermissionError("old DLL is locked")
            return original_replace(source, destination)

        with patch.object(Path, "replace", fail_second_backup):
            with self.assertRaises(core.InstallError) as caught:
                self.transaction().install()
        self.assertTrue(caught.exception.rollback_ok)
        self.assert_old_installation()

    def test_partial_install_failure_restores_previous_installation(self):
        original_replace = Path.replace

        def fail_new_internal(source, destination):
            if source == self.new_root / "_internal":
                raise PermissionError("new directory cannot be installed")
            return original_replace(source, destination)

        with patch.object(Path, "replace", fail_new_internal):
            with self.assertRaises(core.InstallError) as caught:
                self.transaction().install()
        self.assertTrue(caught.exception.rollback_ok)
        self.assert_old_installation()

    def test_version_install_failure_restores_previous_version(self):
        original_replace = Path.replace

        def fail_new_version(source, destination):
            if source == self.new_root / "version.txt":
                raise PermissionError("version cannot be installed")
            return original_replace(source, destination)

        with patch.object(Path, "replace", fail_new_version):
            with self.assertRaises(core.InstallError) as caught:
                self.transaction().install()
        self.assertTrue(caught.exception.rollback_ok)
        self.assert_old_installation()

    def test_rollback_failure_preserves_backup_for_later_recovery(self):
        original_replace = Path.replace
        transaction = self.transaction()

        def fail_install_and_main_restore(source, destination):
            destination = Path(destination)
            if source == self.new_root / "_internal":
                raise PermissionError("installation interrupted")
            if (destination == self.root / "main.exe"
                    and source != self.new_root / "main.exe"):
                raise PermissionError("restoration blocked")
            return original_replace(source, destination)

        with patch.object(Path, "replace", fail_install_and_main_restore):
            with self.assertRaises(core.InstallError) as caught:
                transaction.install()
        self.assertFalse(caught.exception.rollback_ok)
        backup_dir = Path(caught.exception.backup_dir)
        self.assertTrue(backup_dir.exists())
        self.assertTrue(any(path.read_text() == "old executable"
                            for path in backup_dir.rglob("main.exe")))
        core.recover_pending(self.root)
        self.assert_old_installation()

    def test_interrupted_partial_install_recovers_on_next_run(self):
        original_replace = Path.replace

        def interrupt_new_internal(source, destination):
            if source == self.new_root / "_internal":
                raise SimulatedInterruption("process ended unexpectedly")
            return original_replace(source, destination)

        with patch.object(Path, "replace", interrupt_new_internal):
            with self.assertRaises(SimulatedInterruption):
                self.transaction().install()
        messages = core.recover_pending(self.root)
        self.assertTrue(messages)
        self.assert_old_installation()
        core.recover_pending(self.root)
        self.assert_old_installation()

    def test_interruption_after_move_before_return_recovers_on_next_run(self):
        original_replace = Path.replace

        def interrupt_after_new_version(source, destination):
            result = original_replace(source, destination)
            if source == self.new_root / "version.txt":
                raise SimulatedInterruption("crashed after the final file move")
            return result

        with patch.object(Path, "replace", interrupt_after_new_version):
            with self.assertRaises(SimulatedInterruption):
                self.transaction().install()
        self.assert_new_installation()
        core.recover_pending(self.root)
        self.assert_old_installation()

    def test_corrupt_recovery_record_preserves_backup_and_reports_failure(self):
        transaction = self.transaction().install()
        journal = transaction.directory / "transaction.json"
        journal.write_text("{incomplete", encoding="utf-8")
        with self.assertRaises(core.InstallError) as caught:
            core.recover_pending(self.root)
        self.assertFalse(caught.exception.rollback_ok)
        self.assertTrue(transaction.directory.exists())
        self.assert_new_installation()


class ArchiveValidationTests(TemporaryUpdaterTests):
    def test_valid_archive_extracts_expected_payload(self):
        archive = self.base / "update.zip"
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr("tdm-win/main.exe", b"executable")
        destination = self.base / "extract"
        core.safe_extract_zip(archive, destination)
        self.assertEqual((destination / "tdm-win/main.exe").read_bytes(), b"executable")

    def test_unsafe_archive_paths_are_rejected(self):
        for index, name in enumerate(("../outside", "/outside", "C:/outside",
                                      "tdm-win/../../outside", "tdm-win/main.exe:stream")):
            with self.subTest(name=name):
                archive = self.base / f"unsafe-{index}.zip"
                with zipfile.ZipFile(archive, "w") as zipped:
                    zipped.writestr(name, b"outside")
                with self.assertRaises(core.UpdateError):
                    core.safe_extract_zip(archive, self.base / f"extract-{index}")

    def test_archive_symlink_is_rejected(self):
        archive = self.base / "symlink.zip"
        member = zipfile.ZipInfo("tdm-win/link")
        member.create_system = 3
        member.external_attr = (stat.S_IFLNK | 0o777) << 16
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr(member, "../../outside")
        with self.assertRaises(core.UpdateError):
            core.safe_extract_zip(archive, self.base / "extract")

    def test_case_colliding_archive_members_are_rejected_before_extraction(self):
        archive = self.base / "duplicate.zip"
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr("tdm-win/main.exe", b"one")
            zipped.writestr("tdm-win/MAIN.EXE", b"two")
        destination = self.base / "extract"
        with self.assertRaises(core.UpdateError):
            core.safe_extract_zip(archive, destination)
        self.assertFalse(destination.exists())

    def test_checksum_mismatch_and_malformed_checksum_are_rejected(self):
        downloaded = self.base / "download.zip"
        downloaded.write_bytes(b"downloaded payload")
        valid_hash = hashlib.sha256(downloaded.read_bytes()).hexdigest()
        core.verify_sha256(downloaded, valid_hash + "  tdm-win.zip\n")
        for text in ("0" * 64, "", "not-a-hash"):
            with self.subTest(checksum=text):
                with self.assertRaises(core.UpdateError):
                    core.verify_sha256(downloaded, text)


class LaunchFailureTests(TemporaryUpdaterTests):
    def test_missing_executable_prevents_launch(self):
        (self.root / "main.exe").unlink()
        with patch.object(core.subprocess, "Popen") as spawn:
            with self.assertRaises(core.UpdateError):
                core.launch_main(self.root, monitor_seconds=0)
        spawn.assert_not_called()

    def test_missing_license_prevents_launch(self):
        (self.root / "license.json").unlink()
        with patch.object(core.subprocess, "Popen") as spawn:
            with self.assertRaises(core.UpdateError):
                core.launch_main(self.root, monitor_seconds=0)
        spawn.assert_not_called()

    def test_os_launch_error_is_reported(self):
        with patch.object(core.subprocess, "Popen", side_effect=OSError("bad executable")):
            with self.assertRaises(core.UpdateError):
                core.launch_main(self.root, monitor_seconds=0)

    def test_successful_launch_uses_install_directory_and_independent_environment(self):
        process = Mock()
        process.poll.return_value = None
        with patch.object(core.subprocess, "Popen", return_value=process) as spawn:
            result = core.launch_main(self.root, monitor_seconds=0)
        self.assertIs(result, process)
        args, kwargs = spawn.call_args
        self.assertEqual(args[0], [str(self.root / "main.exe")])
        self.assertEqual(kwargs["cwd"], str(self.root))
        self.assertEqual(kwargs["env"]["PYINSTALLER_RESET_ENVIRONMENT"], "1")

    def test_early_exit_is_failure_even_with_zero_exit_code(self):
        for exit_code in (0, 1, 3221225781):
            with self.subTest(exit_code=exit_code):
                process = Mock()
                process.poll.return_value = exit_code
                process.wait.return_value = exit_code
                with patch.object(core.subprocess, "Popen", return_value=process):
                    with self.assertRaises(core.UpdateError):
                        core.launch_main(self.root, monitor_seconds=0.001)


class DownloadFailureTests(TemporaryUpdaterTests):
    def test_incomplete_download_preserves_previous_complete_file(self):
        destination = self.base / "update.zip"
        destination.write_bytes(b"complete previous file")
        response = io.BytesIO(b"short")
        response.headers = {"Content-Length": "100"}
        asset = {"browser_download_url": "https://example.invalid/update.zip", "size": 100}
        with patch.object(core, "urlopen", return_value=response):
            with self.assertRaises(core.UpdateError):
                core.download_asset(asset, destination)
        self.assertEqual(destination.read_bytes(), b"complete previous file")
        self.assertFalse(destination.with_suffix(".part").exists())

    def test_release_missing_zip_asset_reports_failure(self):
        response = json.dumps({"tag_name": "v2.0.0", "assets": []}).encode()
        with patch.object(core, "gh_get", return_value=response):
            with self.assertRaises(core.UpdateError):
                core.fetch_latest_zip_asset()


@unittest.skipUnless(sys.platform == "win32", "Requires Windows kernel mutexes")
class NativeWindowsLockTests(TemporaryUpdaterTests):
    def test_duplicate_lock_is_rejected_and_release_allows_next_updater(self):
        first = core.UpdaterLock(self.root)
        duplicate = core.UpdaterLock(self.root)
        replacement = core.UpdaterLock(self.root)
        for lock in (first, duplicate, replacement):
            self.addCleanup(lock.close)
        first.acquire()
        with self.assertRaises(core.UpdateError):
            duplicate.acquire()
        self.assertIsNone(duplicate.handle)
        self.assertIsNotNone(first.handle)
        first.close()
        replacement.acquire()
        self.assertIsNotNone(replacement.handle)

    def test_different_installations_have_independent_locks(self):
        first = core.UpdaterLock(self.root)
        second = core.UpdaterLock(self.new_root)
        self.addCleanup(first.close)
        self.addCleanup(second.close)
        first.acquire()
        second.acquire()
        self.assertIsNotNone(first.handle)
        self.assertIsNotNone(second.handle)


@unittest.skipUnless(sys.platform == "win32", "Requires Windows ctypes structures")
class WindowsProcessInspectionTests(TemporaryUpdaterTests):
    def mock_processes(self, entries, inaccessible=None):
        """Mock Win32 calls while exercising actual ctypes structure writes."""
        kernel = Mock()
        snapshot = 900
        kernel.CreateToolhelp32Snapshot.return_value = snapshot
        rows = iter(entries)

        def next_entry(handle, pointer):
            self.assertEqual(handle, snapshot)
            try:
                pid, name, _ = next(rows)
            except StopIteration:
                core.ctypes.set_last_error(18)  # ERROR_NO_MORE_FILES
                return False
            entry = pointer._obj
            self.assertEqual(entry.dwSize, core.ctypes.sizeof(entry))
            entry.th32ProcessID = pid
            entry.szExeFile = name
            return True

        def open_process(access, inherit, pid):
            self.assertEqual(access, 0x1000)
            self.assertFalse(inherit)
            if inaccessible and pid in inaccessible:
                core.ctypes.set_last_error(inaccessible[pid])
                return None
            return pid + 1000

        paths = {pid + 1000: str(path) for pid, _, path in entries}

        def query_path(handle, flags, buffer, length_pointer):
            self.assertEqual(flags, 0)
            buffer.value = paths[handle]
            length_pointer._obj.value = len(buffer.value)
            return True

        kernel.Process32FirstW.side_effect = next_entry
        kernel.Process32NextW.side_effect = next_entry
        kernel.OpenProcess.side_effect = open_process
        kernel.QueryFullProcessImageNameW.side_effect = query_path
        kernel.CloseHandle.return_value = True
        return kernel, snapshot

    def test_same_name_in_another_folder_is_skipped_until_exact_path_matches(self):
        kernel, snapshot = self.mock_processes([
            (1, "python.exe", Path(sys.executable)),
            (2, "main.exe", self.new_root / "main.exe"),
            (3, "MAIN.EXE", self.root / "main.exe"),
        ])
        with patch.object(core.ctypes, "WinDLL", return_value=kernel):
            self.assertTrue(core.is_main_running(self.root))
        self.assertEqual(kernel.OpenProcess.call_count, 2)
        self.assertEqual([call.args[0] for call in kernel.CloseHandle.call_args_list],
                         [1002, 1003, snapshot])

    def test_unrelated_main_does_not_block_this_installation(self):
        kernel, snapshot = self.mock_processes([
            (2, "main.exe", self.new_root / "main.exe"),
        ])
        with patch.object(core.ctypes, "WinDLL", return_value=kernel):
            self.assertFalse(core.is_main_running(self.root))
        self.assertEqual([call.args[0] for call in kernel.CloseHandle.call_args_list],
                         [1002, snapshot])

    def test_access_denied_reports_unknown_process_state_and_closes_snapshot(self):
        kernel, snapshot = self.mock_processes([
            (2, "main.exe", self.root / "main.exe"),
        ], inaccessible={2: 5})
        with patch.object(core.ctypes, "WinDLL", return_value=kernel):
            with self.assertRaises(core.UpdateError):
                core.is_main_running(self.root)
        kernel.QueryFullProcessImageNameW.assert_not_called()
        kernel.CloseHandle.assert_called_once_with(snapshot)

    def test_process_that_exited_after_snapshot_is_skipped(self):
        kernel, snapshot = self.mock_processes([
            (2, "main.exe", self.root / "main.exe"),
        ], inaccessible={2: 87})
        with patch.object(core.ctypes, "WinDLL", return_value=kernel):
            self.assertFalse(core.is_main_running(self.root))
        kernel.QueryFullProcessImageNameW.assert_not_called()
        kernel.CloseHandle.assert_called_once_with(snapshot)

    def test_query_failure_preserves_uncertainty_and_closes_all_handles(self):
        kernel, snapshot = self.mock_processes([
            (2, "main.exe", self.root / "main.exe"),
        ])
        kernel.QueryFullProcessImageNameW.side_effect = None
        kernel.QueryFullProcessImageNameW.return_value = False
        with patch.object(core.ctypes, "WinDLL", return_value=kernel):
            with self.assertRaises(core.UpdateError):
                core.is_main_running(self.root)
        self.assertEqual([call.args[0] for call in kernel.CloseHandle.call_args_list],
                         [1002, snapshot])


class VersionValidationTests(unittest.TestCase):
    def test_compares_numeric_components_instead_of_lexicographic_order(self):
        self.assertGreater(core.cmp_semver("1.10.0", "1.9.0"), 0)
        self.assertLess(core.cmp_semver("1.9.0", "2.0.0"), 0)
        self.assertEqual(core.cmp_semver("1.2.3", "1.2.3"), 0)

    def test_malformed_versions_are_not_silently_treated_as_zero(self):
        for version in ("", "garbage", "1..2", "1.2.invalid"):
            with self.subTest(version=version):
                with self.assertRaises(core.UpdateError):
                    core.parse_semver(version)


if __name__ == "__main__":
    unittest.main()
