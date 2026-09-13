from pathlib import Path
from typing import TypeAlias, TypeVar

LowerStr: TypeAlias = str
StrPath: TypeAlias = str | Path
T = TypeVar("T", covariant=True)  # ruff: ignore[PLC0105]
ModelT = TypeVar("ModelT", bound="ItemModel")
RawT = TypeVar("RawT")
