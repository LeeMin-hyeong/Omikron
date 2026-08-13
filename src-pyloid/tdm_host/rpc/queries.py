"""Read-only data query RPC handlers."""

from typing import Any, Dict

from pyloid.rpc import RPCContext

import tdm.aisosik.reader
import tdm.excel.class_info
import tdm.excel.data_file
import tdm.excel.makeup_test
from tdm_host.rpc.responses import failure_response, success_response
from tdm_host.rpc.contracts import CellRequest, TextRequest
from tdm_host.rpc.transport import server


@server.method()
async def get_datafile_data(ctx: RPCContext, mocktest=False) -> Dict[Any, Any]:
    try:
        return success_response(data=tdm.excel.data_file.get_data_sorted_dict(mocktest))
    except Exception as exc:
        return failure_response(exc, context="get_datafile_data")


@server.method()
async def get_aisosik_data(ctx: RPCContext):
    try:
        return success_response(data=tdm.aisosik.reader.get_class_names())
    except Exception as exc:
        return failure_response(exc, context="get_aisosik_data")


@server.method()
async def get_aisosik_student_data(ctx: RPCContext):
    try:
        return success_response(data=tdm.aisosik.reader.get_class_student_dict())
    except Exception as exc:
        return failure_response(exc, context="get_aisosik_student_data")


@server.method()
async def check_aisosik_difference(ctx: RPCContext):
    try:
        aisosik = tdm.aisosik.reader.get_class_student_dict() or {}
        datafile_raw = tdm.excel.data_file.get_data_sorted_dict()
        if isinstance(datafile_raw, (list, tuple)) and datafile_raw:
            datafile = datafile_raw[0]
        else:
            datafile = datafile_raw

        datafile = datafile or {}
        same = all(
            set((student_dict or {}).keys()) == set(aisosik.get(class_name) or [])
            for class_name, student_dict in datafile.items()
        )
        return success_response(data=same)
    except Exception as exc:
        return failure_response(exc, context="check_aisosik_difference")


@server.method()
async def get_makeuptest_data(ctx: RPCContext):
    try:
        return success_response(
            data=tdm.excel.makeup_test.get_student_test_index_dict()
        )
    except Exception as exc:
        return failure_response(exc, context="get_makeuptest_data")


@server.method()
async def get_class_list(ctx: RPCContext):
    try:
        return success_response(data=tdm.excel.class_info.get_class_names())
    except Exception as exc:
        return failure_response(exc, context="get_class_list")


@server.method()
async def get_class_info(ctx: RPCContext, class_name: str):
    try:
        request = TextRequest.validate(class_name, "class_name")
        return success_response(data=tdm.excel.class_info.get_class_info(request.value))
    except Exception as exc:
        return failure_response(exc, context="get_class_info")


@server.method()
async def get_new_class_list(ctx: RPCContext):
    try:
        return success_response(data=tdm.excel.class_info.get_new_class_names())
    except Exception as exc:
        return failure_response(exc, context="get_new_class_list")


@server.method()
async def is_cell_empty(ctx: RPCContext, row: int, col: int):
    try:
        request = CellRequest.validate(row, col)
        empty, value = tdm.excel.data_file.is_cell_empty(request.row, request.col)
        return success_response(data={"empty": empty, "value": value})
    except Exception as exc:
        return failure_response(exc, context="is_cell_empty")
