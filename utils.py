from __future__ import annotations

from collections.abc import Callable, Iterable
from enum import IntEnum
from typing import cast, overload

from .constants import _UNSET, SMTP_ADDRESS_SCHEMA
from .exceptions import COM_ERRORS
from .protocols import OlAddressEntry, OlCollection, OlObject
from .types import LowerStr, ModelT, RawT, T


def is_accessible_ol_item(
    item: OlObject, target_type: IntEnum, properties: Iterable[str] | None = None
) -> bool:
    """Check whether an Outlook COM object is accessible and has the expected type.

    Parameters
    ----------
    item : OlObject
        Outlook COM object to inspect.
    target_type : IntEnum
        Expected Outlook object class.
    properties : iterable of str, optional
        Properties used as a fallback accessibility check when ``Class`` is
        unavailable.

    Returns
    -------
    bool
        ``True`` when the object is accessible and matches the expected type.
    """
    if item is None:
        return False
    try:
        item_type = getattr(item, "Class", _UNSET)
        if item_type is not _UNSET:
            return item_type == target_type
        properties = properties or ()
        for name in properties:
            getattr(item, name)
        return True
    except COM_ERRORS:
        return False


@overload
def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: None = None,
    limit: int | None = None,
    predicate: Callable[[T], bool] | None = None,
) -> list[T]:
    """Type an unpacked collection without model transformation.

    Parameters
    ----------
    collection : OlCollection[T], optional
        Collection to unpack.
    transformer : None, default=None
        Disable model transformation.
    limit : int, optional
        Maximum number of matching items.
    predicate : callable, optional
        Filter applied to each item.

    Returns
    -------
    list of T
        Unpacked collection items.
    """
    ...  # noqa: PIE790


@overload
def unpack_collection(
    collection: OlCollection[RawT] | None,
    *,
    transformer: type[ModelT],
    limit: int | None = None,
    predicate: Callable[[RawT], bool] | None = None,
) -> list[ModelT]:
    """Type an unpacked collection with model transformation.

    Parameters
    ----------
    collection : OlCollection[RawT], optional
        Collection to unpack.
    transformer : type[ModelT]
        Model class used to wrap raw items.
    limit : int, optional
        Maximum number of matching items.
    predicate : callable, optional
        Filter applied to each raw item.

    Returns
    -------
    list of ModelT
        Wrapped collection items.
    """
    ...  # noqa: PIE790


def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: type[ModelT] | None = None,
    limit: int | None = None,
    predicate: Callable[[T], bool] | None = None,
) -> list[T] | list[ModelT]:
    """Convert an Outlook collection to a list.

    Parameters
    ----------
    collection : OlCollection or None
        One-indexed Outlook collection to unpack.
    transformer : type, optional
        Model class used to wrap each raw item.
    limit : int, optional
        Maximum number of matching items to return.
    predicate : callable, optional
        Filter applied to each raw item before transformation.

    Returns
    -------
    list
        Accessible items in collection order.
    """
    if collection is None or limit is not None and limit <= 0:
        return []
    result = []
    for idx in range(1, collection.Count + 1):
        item = collection.Item(idx)
        if item is None or predicate is not None and not predicate(item):
            continue
        value = item if transformer is None else transformer.from_outlook_item(item)
        if value is not None:
            result.append(value)
            if limit is not None and len(result) >= limit:
                break
    return cast(list[T] | list[ModelT], result)


def get_smtp_address(user: OlAddressEntry) -> LowerStr:
    """Return the SMTP address for an Outlook address entry.

    Parameters
    ----------
    user : OlAddressEntry
        Address entry to inspect.

    Returns
    -------
    str
        The lowercase SMTP address, or an empty string when it cannot be read.

    Notes
    -----
    Exchange entries use their primary SMTP address. Other entries use the
    MAPI SMTP property before falling back to the standard address.
    """
    try:
        if not user:
            return ""
        try:
            exch_user = user.GetExchangeUser()
        except COM_ERRORS:
            exch_user = None
        if exch_user:
            try:
                primary_address = exch_user.PrimarySmtpAddress
            except COM_ERRORS:
                primary_address = ""
            if primary_address:
                return str(primary_address).lower()
        try:
            address = user.PropertyAccessor.GetProperty(SMTP_ADDRESS_SCHEMA)
            return str(address).lower()
        except COM_ERRORS:
            return str(user.Address).lower()
    except COM_ERRORS:
        return ""
