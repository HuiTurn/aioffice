"""Verified binary assets returned from a native document."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from aioffice.core.errors import NativePackageError, SecurityError
from aioffice.security import SecurityPolicy
from aioffice.spec.models import AssetRef

_IMAGE_FORMATS = {
    "image/png": ("png",),
    "image/jpeg": ("jpg", "jpeg"),
    "image/gif": ("gif",),
    "image/bmp": ("bmp",),
    "image/tiff": ("tif", "tiff"),
}


def _detected_media_type(payload: bytes) -> str | None:
    if payload.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if payload.startswith(b"\xff\xd8\xff") and payload.endswith(b"\xff\xd9"):
        return "image/jpeg"
    if payload.startswith((b"GIF87a", b"GIF89a")) and payload.endswith(b";"):
        return "image/gif"
    if payload.startswith(b"BM") and len(payload) >= 26:
        return "image/bmp"
    if payload.startswith((b"II*\x00", b"MM\x00*")):
        return "image/tiff"
    return None


@dataclass(frozen=True, slots=True)
class PreparedImageAsset:
    """Validated, content-addressed image input kept outside the JSON Spec."""

    asset: AssetRef
    data: bytes


def prepare_image_asset(
    source: bytes | bytearray | memoryview | str | Path | "ImageAsset",
    *,
    media_type: str | None,
    security_policy: SecurityPolicy,
) -> PreparedImageAsset:
    """Read and signature-check one bounded raster image input."""

    if media_type is not None and not isinstance(media_type, str):
        raise NativePackageError(
            "Replacement image media_type must be a string when provided."
        )
    if isinstance(source, ImageAsset):
        payload = source.data
    elif isinstance(source, (bytes, bytearray, memoryview)):
        payload = bytes(source)
    else:
        path = Path(source)
        size = path.stat().st_size
        if size > security_policy.max_file_size_bytes:
            raise SecurityError(
                f"Replacement image size {size} exceeds "
                f"{security_policy.max_file_size_mb} MB."
            )
        payload = path.read_bytes()
    if not payload:
        raise NativePackageError("Replacement image cannot be empty.")
    if len(payload) > security_policy.max_file_size_bytes:
        raise SecurityError(
            f"Replacement image size {len(payload)} exceeds "
            f"{security_policy.max_file_size_mb} MB."
        )

    detected = _detected_media_type(payload)
    if detected is None:
        raise NativePackageError(
            "Replacement image must be a signature-valid PNG, JPEG, GIF, BMP, or TIFF."
        )
    declared = media_type.casefold() if media_type is not None else detected
    if declared not in _IMAGE_FORMATS:
        raise NativePackageError(
            "Replacement image media_type must be image/png, image/jpeg, image/gif, "
            "image/bmp, or image/tiff."
        )
    if declared != detected:
        raise NativePackageError(
            f"Declared replacement media type {declared!r} does not match "
            f"detected type {detected!r}."
        )

    digest = hashlib.sha256(payload).hexdigest()
    extension = _IMAGE_FORMATS[detected][0]
    filename = f"aioffice-{digest}.{extension}"
    asset = AssetRef(
        id=f"asset_{digest}",
        sha256=digest,
        media_type=detected,
        filename=filename,
        size_bytes=len(payload),
    )
    return PreparedImageAsset(asset=asset, data=payload)


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


__all__ = [
    "ImageAsset",
    "PreparedImageAsset",
    "prepare_image_asset",
]
