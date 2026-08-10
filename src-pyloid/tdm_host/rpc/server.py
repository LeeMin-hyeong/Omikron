"""RPC server assembly.

Importing this module registers every public method on the shared transport.
"""

from tdm_host.rpc.transport import server

# Imports are intentionally side-effectful: decorators register the public methods.
from tdm_host.rpc import commands as _commands  # noqa: F401
from tdm_host.rpc import jobs as _jobs  # noqa: F401
from tdm_host.rpc import queries as _queries  # noqa: F401
from tdm_host.rpc import startup as _startup  # noqa: F401

__all__ = ["server"]