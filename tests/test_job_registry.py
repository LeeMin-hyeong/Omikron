from tdm_host.jobs.registry import JobRegistry
from tdm.domain.errors import InvalidOperationError, JobAlreadyRunningError


def test_job_registry_only_increments_revision_for_meaningful_changes():
    registry = JobRegistry(finished_ttl_seconds=60)
    registry.create(
        "job-1",
        "save",
        {
            "status": "running",
            "level": "info",
            "message": "대기 중",
            "warnings": [],
            "ts": 1.0,
        },
    )

    registry.emit("job-1", {"status": "running", "message": "대기 중", "ts": 2.0})
    assert registry.get("job-1")["revision"] == 1

    registry.emit("job-1", {"status": "running", "message": "처리 중", "ts": 3.0})
    assert registry.get("job-1")["revision"] == 2


def test_job_registry_batch_omits_unchanged_state():
    registry = JobRegistry(finished_ttl_seconds=60)
    registry.create("job-1", "save", {"status": "running", "warnings": []})

    unchanged = registry.get_many([{"jobId": "job-1", "revision": 1}])[0]
    changed = registry.get_many([{"jobId": "job-1", "revision": 0}])[0]

    assert unchanged == {"jobId": "job-1", "revision": 1, "changed": False}
    assert changed["changed"] is True
    assert changed["state"]["jobType"] == "save"


def test_job_registry_removes_finished_state_after_ttl():
    registry = JobRegistry(finished_ttl_seconds=1)
    registry.create("job-1", "save", {"status": "running", "warnings": []})
    registry.emit("job-1", {"status": "done", "level": "success"})
    finished_at = registry.get("job-1")["finishedAt"]

    registry.cleanup_expired(now=finished_at + 1.1)

    assert registry.get("job-1")["status"] == "unknown"


def test_job_registry_rejects_second_job_in_same_exclusive_group():
    registry = JobRegistry(finished_ttl_seconds=60)
    registry.reserve(
        "job-1",
        "save",
        {"status": "running", "warnings": []},
        exclusive_group="excel-write",
    )

    try:
        registry.reserve(
            "job-2",
            "update",
            {"status": "running", "warnings": []},
            exclusive_group="excel-write",
        )
    except JobAlreadyRunningError:
        pass
    else:
        raise AssertionError("같은 실행 그룹의 두 번째 작업이 허용되었습니다.")


def test_job_registry_allows_group_again_after_terminal_state():
    registry = JobRegistry(finished_ttl_seconds=60)
    registry.reserve(
        "job-1",
        "save",
        {"status": "running", "warnings": []},
        exclusive_group="excel-write",
    )
    registry.emit("job-1", {"status": "done", "level": "success"})

    registry.reserve(
        "job-2",
        "update",
        {"status": "running", "warnings": []},
        exclusive_group="excel-write",
    )
    assert registry.get("job-2")["status"] == "running"


def test_job_registry_rejects_new_jobs_after_shutdown():
    registry = JobRegistry(finished_ttl_seconds=60)
    registry.shutdown(timeout=0)

    try:
        registry.reserve("job-1", "save", {"status": "running"})
    except InvalidOperationError:
        pass
    else:
        raise AssertionError("종료가 시작된 뒤 새 작업이 허용되었습니다.")
