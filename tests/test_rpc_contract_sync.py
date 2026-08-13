"""Keep the Python RPC surface, TypeScript gateway, and shared manifest in sync."""

from __future__ import annotations

import ast
import json
import re
from pathlib import Path

from tdm_host.rpc.error_codes import RpcErrorCode
from tdm.domain.errors import TDMError


ROOT = Path(__file__).resolve().parents[1]


def _async_functions(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    return {node.name for node in tree.body if isinstance(node, ast.AsyncFunctionDef)}


def test_rpc_contract_is_synchronized() -> None:
    contract = json.loads((ROOT / "contracts" / "rpc_contract.json").read_text(encoding="utf-8"))
    rpc_dir = ROOT / "src-pyloid" / "tdm_host" / "rpc"

    backend_general = set()
    for filename in ("commands.py", "queries.py", "startup.py"):
        backend_general.update(_async_functions(rpc_dir / filename))
    assert backend_general == set(contract["generalMethods"])
    assert _async_functions(rpc_dir / "jobs.py") == set(contract["jobMethods"])
    assert {code.value for code in RpcErrorCode} == set(contract["errorCodes"])
    domain_codes = {TDMError.code}
    pending = list(TDMError.__subclasses__())
    while pending:
        error_type = pending.pop()
        domain_codes.add(error_type.code)
        pending.extend(error_type.__subclasses__())
    assert domain_codes <= set(contract["errorCodes"])

    gateway = (ROOT / "src" / "api" / "rpc.ts").read_text(encoding="utf-8")
    methods_block = re.search(
        r"GENERAL_RPC_METHODS: GeneralRpcMethod\[\] = \[(.*?)\];",
        gateway,
        re.DOTALL,
    )
    assert methods_block is not None
    frontend_methods = set(re.findall(r'"([a-z][a-z0-9_]*)"', methods_block.group(1)))
    assert frontend_methods == set(contract["generalMethods"])

    error_block = re.search(
        r"RPC_ERROR_CODES = \[(.*?)\] as const;",
        gateway,
        re.DOTALL,
    )
    assert error_block is not None
    frontend_codes = set(re.findall(r'"([A-Z][A-Z0-9_]*)"', error_block.group(1)))
    assert frontend_codes == set(contract["errorCodes"])
