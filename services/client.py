from typing import Any

from ..enums import FolderEnum
from ..exceptions import COM_ERRORS, OutlookError
from ..models.account import Account
from ..validation import validate_email


def _load_dispatch() -> Any:
    """Load the pywin32 COM client module.

    Parameters
    ----------
    None

    Returns
    -------
    Any
        pywin32 COM client module.
    """
    try:
        from win32com import client  # type: ignore

        return client
    except ImportError as exc:
        raise OutlookError(
            "win32com.client is required. Install pywin32 on Windows."
        ) from exc


def _connect() -> Any:
    """Connect to the Outlook application.

    Parameters
    ----------
    None

    Returns
    -------
    Any
        Outlook application COM object.
    """
    try:
        return _load_dispatch().Dispatch("Outlook.Application")
    except OutlookError:
        raise
    except COM_ERRORS as exc:
        raise OutlookError("Unable to connect to Outlook.") from exc


def _open_mapi(app: Any) -> Any:
    """Open the MAPI namespace for an Outlook application.

    Parameters
    ----------
    app : Any
        Outlook application COM object.

    Returns
    -------
    Any
        MAPI namespace COM object.
    """
    try:
        return app.GetNamespace("MAPI")
    except COM_ERRORS as exc:
        raise OutlookError("Unable to open the Outlook MAPI namespace.") from exc


def _verify_mailbox(mapi: Any, address: str, account: Account | None) -> str:
    """Resolve a sending address and verify access to its Inbox.

    Parameters
    ----------
    mapi : Any
        Active Outlook namespace.
    address : str
        Shared mailbox SMTP address or configured account name/address.
    account : Account or None
        Matching configured account, if present.

    Returns
    -------
    str
        Normalized sending address.

    Raises
    ------
    OutlookError
        If the address is invalid, unresolved, or its Inbox is inaccessible.

    Notes
    -----
    Inbox access does not verify Exchange Send As permission.
    """
    address = validate_email(account.email_address if account else address)
    if not address or ";" in address:
        raise OutlookError("Provide exactly one sending mailbox address.")
    try:
        if account is not None:
            inbox = account.store.GetDefaultFolder(FolderEnum.INBOX)
        else:
            recipient = mapi.CreateRecipient(address)
            if not recipient.Resolve():
                raise OutlookError(
                    f"Outlook mailbox could not be resolved: {address!r}."
                )
            inbox = mapi.GetSharedDefaultFolder(recipient, FolderEnum.INBOX)
        # Force a read; resolving a directory entry alone does not prove access.
        _ = inbox.Items.Count
    except COM_ERRORS as exc:
        raise OutlookError(f"Outlook mailbox is not accessible: {address!r}.") from exc
    return address


def _select_account(
    accounts: list[Account], value: str | None = None
) -> Account | None:
    """Select an Outlook account by name or SMTP address.

    Parameters
    ----------
    accounts : list of Account
        Available accounts.
    value : str, optional
        Case-insensitive account name or SMTP address.

    Returns
    -------
    Account or None
        Matching account, the sole account, or ``None`` when selection is ambiguous.
    """
    if not accounts:
        raise OutlookError("Outlook has no configured email accounts.")
    if value:
        account = next((item for item in accounts if item.matches(value)), None)
        if account is None:
            available = ", ".join(sorted(item.email_address for item in accounts))
            raise OutlookError(
                f"Outlook account not found: {value!r}. Available accounts: {available}."
            )
        return account
    return accounts[0] if len(accounts) == 1 else None
