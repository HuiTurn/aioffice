"""Stable, sortable identifiers for artifacts and semantic nodes."""

from __future__ import annotations

import re
import secrets
import time

_CROCKFORD = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"
_PREFIX_RE = re.compile(r"^[a-z][a-z0-9_]{0,31}$")


def _encode_base32(value: int, length: int) -> str:
    chars = ["0"] * length
    for index in range(length - 1, -1, -1):
        chars[index] = _CROCKFORD[value & 31]
        value >>= 5
    return "".join(chars)


def new_id(prefix: str = "node") -> str:
    """Return a prefixed ULID-compatible identifier.

    The timestamp makes IDs roughly sortable while 80 random bits make collisions
    impractical. The prefix remains readable in serialized specs and diagnostics.
    """

    if not _PREFIX_RE.fullmatch(prefix):
        raise ValueError("ID prefix must start with a lowercase letter and contain [a-z0-9_].")
    timestamp_ms = int(time.time_ns() // 1_000_000) & ((1 << 48) - 1)
    randomness = secrets.randbits(80)
    encoded = _encode_base32((timestamp_ms << 80) | randomness, 26)
    return f"{prefix}_{encoded}"
