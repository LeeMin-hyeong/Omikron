"""Shared RPC transport instance.

Handler modules register methods on this object when imported.
"""

from pyloid.rpc import PyloidRPC

server = PyloidRPC()

__all__ = ["server"]
