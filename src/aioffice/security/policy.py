"""Resource and feature limits applied before parsing Office packages."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class SecurityPolicy:
    allow_macros: bool = False
    allow_external_relationships: bool = True
    allow_network_assets: bool = False
    max_file_size_mb: int = 100
    max_uncompressed_size_mb: int = 500
    max_package_parts: int = 10_000
    max_compression_ratio: float = 200.0
    max_xml_part_size_mb: int = 32

    @property
    def max_file_size_bytes(self) -> int:
        return self.max_file_size_mb * 1024 * 1024

    @property
    def max_uncompressed_size_bytes(self) -> int:
        return self.max_uncompressed_size_mb * 1024 * 1024

    @property
    def max_xml_part_size_bytes(self) -> int:
        return self.max_xml_part_size_mb * 1024 * 1024
