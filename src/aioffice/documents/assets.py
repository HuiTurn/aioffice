"""Verified binary assets returned from a native document."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class ImageAsset:
    """One extracted image occurrence plus its verified native bytes."""

    image_id: str
    asset_id: str
    media_type: str
    filename: str
    sha256: str
    data: bytes

    @property
    def size_bytes(self) -> int:
        return len(self.data)

    def write(
        self,
        target: str | Path,
        *,
        overwrite: bool = False,
    ) -> Path:
        path = Path(target)
        if path.exists() and not overwrite:
            raise FileExistsError(
                f"Refusing to overwrite existing image asset {path}."
            )
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self.data)
        return path


__all__ = ["ImageAsset"]
