"""Return locally detected, masked network identifiers for audit records."""

from __future__ import annotations

import socket
import uuid
from functools import lru_cache


def _masked_local_ip() -> str:
    """Detect the outbound interface address without sending a network packet."""
    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("192.0.2.1", 9))
        address = probe.getsockname()[0]
    except OSError:
        return "tidak tersedia"
    finally:
        probe.close()
    parts = address.split(".")
    if len(parts) == 4:
        return ".".join((*parts[:3], "*"))
    return "tersedia (masked)"


def _masked_mac() -> str:
    """Read the host interface identifier and retain only its first 3 bytes."""
    node = uuid.getnode()
    # uuid.getnode() sets bit 40 when it had to create a random identifier.
    if node == 0 or node & (1 << 40):
        return "tidak tersedia"
    octets = [f"{(node >> shift) & 0xff:02X}" for shift in range(40, -1, -8)]
    return ":".join((*octets[:3], "**", "**", "**"))


@lru_cache(maxsize=1)
def masked_device_identity() -> tuple[str, str]:
    """Return masked local IP and MAC; raw identifiers are never returned."""
    return _masked_local_ip(), _masked_mac()
