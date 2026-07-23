"""Strict, AI-facing operation selector models."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class OperationModel(BaseModel):
    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)


class TextRange(OperationModel):
    """A half-open range measured in Unicode code points."""

    start: int = Field(ge=0, strict=True)
    end: int = Field(ge=0, strict=True)
    unit: Literal["unicode_codepoint"] = "unicode_codepoint"

    @model_validator(mode="after")
    def validate_non_empty(self) -> "TextRange":
        if self.end <= self.start:
            raise ValueError("Text range end must be greater than start.")
        return self


class TextMatch(OperationModel):
    """An exact, non-overlapping text occurrence selector."""

    text: str = Field(min_length=1)
    occurrence: int = Field(default=1, ge=1, strict=True)


__all__ = ["OperationModel", "TextMatch", "TextRange"]
