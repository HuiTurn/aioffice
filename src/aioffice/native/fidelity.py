"""Round-trip fidelity contracts and reports."""

from __future__ import annotations

from enum import StrEnum

from pydantic import BaseModel, ConfigDict, Field


class FidelityPolicy(StrEnum):
    STRICT = "strict"
    PRESERVE_UNKNOWN = "preserve_unknown"
    REBUILD = "rebuild"


class FidelityLevel(StrEnum):
    EXACT_PACKAGE = "exact_package"
    EXACT_PARTS = "exact_parts"
    STRUCTURAL = "structural"
    VISUAL = "visual"
    SEMANTIC = "semantic"
    LOSSY = "lossy"


class FidelityReport(BaseModel):
    model_config = ConfigDict(extra="forbid")

    policy: FidelityPolicy
    level: FidelityLevel
    source_sha256: str
    result_sha256: str
    affected_parts: list[str] = Field(default_factory=list)
    untouched_parts: int = 0
    opaque_features_preserved: list[str] = Field(default_factory=list)
    lossy_features: list[str] = Field(default_factory=list)
    visual_verification_required: bool = False

    @property
    def lossless(self) -> bool:
        return not self.lossy_features and self.level != FidelityLevel.LOSSY
