"""Read-only data query RPC handlers."""

import traceback
from typing import Any, Dict

from pyloid.rpc import RPCContext

import tdm.aisosik.reader
import tdm.excel.class_info
import tdm.excel.data_file
import tdm.excel.makeup_test
from tdm_host.rpc.transport import server

@server.method()
async def get_datafile_data(ctx: RPCContext, mocktest = False) -> Dict[Any, Any]:
    try:
        return {"ok": True, "data": tdm.excel.data_file.get_data_sorted_dict(mocktest)}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_aisosic_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.aisosik.reader.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_aisosic_student_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.aisosik.reader.get_class_student_dict()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def check_aisosic_difference(ctx: RPCContext):
    try:
        aisosic = tdm.aisosik.reader.get_class_student_dict()
        datafile_raw = tdm.excel.data_file.get_data_sorted_dict()
        if isinstance(datafile_raw, (list, tuple)) and len(datafile_raw) >= 1:
            datafile = datafile_raw[0]
        else:
            datafile = datafile_raw

        aisosic = aisosic or {}
        datafile = datafile or {}

        same = True
        for class_name, student_dict in datafile.items():
            datafile_students = set((student_dict or {}).keys())
            aisosic_students = set(aisosic.get(class_name) or [])
            if datafile_students != aisosic_students:
                same = False
                break
        return {"ok": True, "data": same}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_makeuptest_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.excel.makeup_test.get_studnet_test_index_dict()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.excel.class_info.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_class_info(ctx: RPCContext, class_name:str):
    try:
        return {"ok": True, "data": tdm.excel.class_info.get_class_info(class_name)}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_new_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.excel.class_info.get_new_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def is_cell_empty(ctx: RPCContext, row:int, col:int):
    try:
        empty, value = tdm.excel.data_file.is_cell_empty(row, col)
        return {"ok": True, "empty": empty, "value": value}
    except Exception as e:
            return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


####################################### 작업 API #######################################
