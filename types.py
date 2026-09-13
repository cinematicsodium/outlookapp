from pathlib import Path
from typing import TYPE_CHECKING, TypeAlias, TypeVar

if TYPE_CHECKING:
    from .models.base import BaseModel

LowerStr: TypeAlias = str
StrPath: TypeAlias = str | Path
T = TypeVar("T", covariant=True)  # ruff: ignore[PLC0105]
ModelT = TypeVar("ModelT", bound="ItemModel")
RawT = TypeVar("RawT")
