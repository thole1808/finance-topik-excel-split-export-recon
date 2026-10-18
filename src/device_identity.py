"""Return locally detected network identifiers for audit records."""

from __future__ import annotations

import socket
import uuid
from functools import lru_cache


def _local_ip() -> str:
    """Detect the outbound interface address without sending a network packet."""
    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("192.0.2.1", 9))
        address = probe.getsockname()[0]
    except OSError:
        return "tidak tersedia"
    finally:
        probe.close()
    return address


def _mac() -> str:
    """Read the host interface identifier if the OS exposes a stable value."""
    node = uuid.getnode()
    # uuid.getnode() sets bit 40 when it had to create a random identifier.
    if node == 0 or node & (1 << 40):
        return "tidak tersedia"
    octets = [f"{(node >> shift) & 0xff:02X}" for shift in range(40, -1, -8)]
    return ":".join(octets)


@lru_cache(maxsize=1)
def device_identity() -> tuple[str, str]:
    """Return the local IP and MAC for the user's explicitly requested audit."""
    return _local_ip(), _mac()


def masked_device_identity() -> tuple[str, str]:
    """Return masked values for the on-screen history view."""
    ip_address, mac_address = device_identity()
    parts = ip_address.split(".")
    if len(parts) == 4:
        ip_address = ".".join((*parts[:3], "*"))
    elif ip_address != "tidak tersedia":
        ip_address = "tersedia (masked)"
    octets = mac_address.split(":")
    if len(octets) == 6:
        mac_address = ":".join((*octets[:3], "**", "**", "**"))
    return ip_address, mac_address
