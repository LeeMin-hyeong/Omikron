"""Background-job RPC handlers."""

from typing import Any
from collections.abc import Callable

from pyloid.rpc import RPCContext

from tdm_host.jobs.manager import job_manager
from tdm_host.jobs.registry import get_progress_payload, registry
from tdm_host.jobs.workers import (
    save_exam_job_process,
    send_exam_message_job_process,
    update_class_job_process,
)
from tdm_host.rpc.responses import failure_response, success_response
from tdm_host.rpc.contracts import JobBatchRequest, JobIdRequest, JobUploadRequest
from tdm_host.rpc.error_codes import RpcErrorCode
from tdm_host.rpc.responses import error_response
from tdm_host.rpc.transport import server


def _start_process_job(
    *,
    job_type: str,
    target: Callable[..., None],
    total: int,
    message: str,
    kwargs: dict[str, Any],
) -> dict[str, Any]:
    try:
        return job_manager.start(
            job_type=job_type,
            target=target,
            total=total,
            message=message,
            kwargs=kwargs,
            exclusive_group="user-operation",
        )
    except Exception as exc:
        return failure_response(exc, context=f"{job_type}.start")


@server.method()
async def get_progress(ctx: RPCContext, job_id: str) -> dict[str, Any]:
    return get_progress_payload(job_id)


@server.method()
async def get_job_progress_batch(
    ctx: RPCContext, jobs: list[dict[str, Any]]
) -> dict[str, Any]:
    try:
        request = JobBatchRequest.validate(jobs)
        return success_response(data={"jobs": registry.get_many(request.jobs)})
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))


@server.method()
async def get_job(ctx: RPCContext, job_id: str) -> dict[str, Any]:
    try:
        request = JobIdRequest.validate(job_id)
        state = registry.get(request.job_id)
        if state.get("status") == "unknown":
            return error_response(RpcErrorCode.JOB_NOT_FOUND, "작업을 찾을 수 없습니다.")
        return success_response(data=state)
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))


@server.method()
async def acknowledge_job_completion(ctx: RPCContext, job_id: str) -> dict[str, Any]:
    try:
        request = JobIdRequest.validate(job_id)
        if not registry.acknowledge(request.job_id):
            return error_response(RpcErrorCode.JOB_NOT_FOUND, "작업을 찾을 수 없습니다.")
        return success_response(data={"jobId": request.job_id, "acknowledged": True})
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))


@server.method()
async def cancel_job(ctx: RPCContext, job_id: str) -> dict[str, Any]:
    try:
        request = JobIdRequest.validate(job_id)
        result = registry.request_cancel(request.job_id)
        if result == "not_found":
            return error_response(RpcErrorCode.JOB_NOT_FOUND, "작업을 찾을 수 없습니다.")
        if result == "not_cancellable":
            return error_response(
                RpcErrorCode.JOB_NOT_CANCELLABLE,
                "이미 완료된 작업은 취소할 수 없습니다.",
            )
        return success_response(data={"jobId": request.job_id, "cancellationRequested": True})
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))


@server.method()
async def start_send_exam_message(
    ctx: RPCContext,
    filename: str,
    b64: str,
    makeup_test_date: dict[str, Any],
) -> dict[str, Any]:
    try:
        request = JobUploadRequest.validate(filename, b64, makeup_test_date)
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))
    return _start_process_job(
        job_type="start_send_exam_message",
        target=send_exam_message_job_process,
        total=3,
        message="작업 대기 중...",
        kwargs={
            "filename": request.filename,
            "b64": request.b64,
            "makeup_test_date": request.makeup_test_date,
        },
    )


@server.method()
async def start_save_exam(
    ctx: RPCContext,
    filename: str,
    b64: str,
    makeup_test_date: dict[str, Any],
) -> dict[str, Any]:
    try:
        request = JobUploadRequest.validate(filename, b64, makeup_test_date)
    except (TypeError, ValueError) as exc:
        return error_response(RpcErrorCode.INVALID_REQUEST, str(exc))
    return _start_process_job(
        job_type="start_save_exam",
        target=save_exam_job_process,
        total=6,
        message="작업 대기 중...",
        kwargs={
            "filename": request.filename,
            "b64": request.b64,
            "makeup_test_date": request.makeup_test_date,
        },
    )


@server.method()
async def start_update_class(ctx: RPCContext) -> dict[str, Any]:
    return _start_process_job(
        job_type="start_update_class",
        target=update_class_job_process,
        total=6,
        message="반 업데이트 준비중...",
        kwargs={},
    )
