from __future__ import annotations

import logging
from functools import cached_property
from types import TracebackType

from ..exceptions import COM_ERRORS, OutlookError
from ..protocols import OlApplication, OlMailItem, OlNamespace
from ..services.client import _connect, _open_mapi, _select_account, _verify_mailbox
from ..utils import unpack_collection
from .account import Account
from .mail_item import MailItem

logger = logging.getLogger(__name__)


class Outlook:
    """Connect to Outlook and select a sending mailbox.

    Parameters
    ----------
    address : str, optional
        Shared mailbox SMTP address or configured account name/address. When
        supplied, its Inbox must be accessible. Shared mailboxes require Exchange
        Send As permission for recipients to see only the mailbox as sender.
        With ``None``, select the sole configured account, if unambiguous.
    app : OlApplication, optional
        Existing Outlook application object, primarily for dependency injection.
    mapi : OlNamespace, optional
        Existing MAPI namespace, primarily for dependency injection.

    Raises
    ------
    OutlookError
        If Outlook cannot be opened, has no configured accounts, or ``address``
        cannot be resolved or its Inbox cannot be accessed.
    """

    def __init__(
        self,
        address: str | None = None,
        app: OlApplication | None = None,
        mapi: OlNamespace | None = None,
    ) -> None:
        """Connect to Outlook and optionally verify a sending mailbox.

        Parameters
        ----------
        address : str, optional
            Shared mailbox SMTP address or configured account name/address.
        app : OlApplication, optional
            Existing Outlook application COM object.
        mapi : OlNamespace, optional
            Existing Outlook MAPI namespace.

        Returns
        -------
        None
        """
        if address is not None and (
            not isinstance(address, str) or not address.strip()
        ):
            raise OutlookError("Provide a nonempty mailbox address or None.")
        self._app = _connect() if app is None else app
        self._mapi = _open_mapi(self._app) if mapi is None else mapi
        self.account = _select_account(self.accounts)
        if address is None:
            self.address = self.account.email_address if self.account else None
        else:
            self.account = self.find_account(address.strip())
            self.address = _verify_mailbox(self._mapi, address, self.account)

    def __repr__(self) -> str:
        """Return a developer representation of the Outlook connection.

        Parameters
        ----------
        None

        Returns
        -------
        str
            Connection representation.
        """
        return f"Outlook(address={self.address!r}, connected={self._app is not None})"

    __str__ = __repr__

    def __enter__(self) -> Outlook:  # ruff: ignore[PYI034]
        """Enter the Outlook connection context.

        Parameters
        ----------
        None

        Returns
        -------
        Outlook
            Active Outlook connection.
        """
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        """Close the Outlook connection when leaving a context.

        Parameters
        ----------
        exc_type : type of BaseException, optional
            Exception type raised in the context.
        exc : BaseException, optional
            Exception raised in the context.
        traceback : TracebackType, optional
            Traceback for the exception.

        Returns
        -------
        None
        """
        self.close()

    @cached_property
    def accounts(self) -> list[Account]:
        """Return all accounts in the active Outlook profile."""
        try:
            return unpack_collection(
                self._require_mapi().Accounts,
                transformer=Account,
            )
        except COM_ERRORS as exc:
            raise OutlookError("Unable to read Outlook accounts.") from exc

    def _require_app(self) -> OlApplication:
        """Return the active Outlook application.

        Parameters
        ----------
        None

        Returns
        -------
        OlApplication
            Active application COM object.
        """
        if self._app is None:
            raise OutlookError("Outlook connection is closed or unavailable.")
        return self._app

    def _require_mapi(self) -> OlNamespace:
        """Return the active Outlook MAPI namespace.

        Parameters
        ----------
        None

        Returns
        -------
        OlNamespace
            Active MAPI namespace COM object.
        """
        if self._mapi is None:
            raise OutlookError("Outlook namespace is closed or unavailable.")
        return self._mapi

    def _require_account(self) -> Account:
        """Return the selected Outlook account.

        Parameters
        ----------
        None

        Returns
        -------
        Account
            Selected Outlook account.
        """
        if self.account is None:
            raise OutlookError("No Outlook account is selected.")
        return self.account

    def find_account(self, value: str) -> Account | None:
        """Find an account by display name or SMTP address.

        Parameters
        ----------
        value : str
            Case-insensitive display name or SMTP address.

        Returns
        -------
        Account or None
            The matching account, if present.
        """
        return next(
            (account for account in self.accounts if account.matches(value)),
            None,
        )

    def new_email(self) -> MailItem:
        """Create a new email with the selected mailbox in From.

        Returns
        -------
        MailItem
            A new unsaved email message.

        Raises
        ------
        OutlookError
            If the connection is closed, no sending mailbox is selected, or Outlook
            does not return an accessible mail item.
        """
        app = self._require_app()
        address = self.address or self._require_account().email_address
        try:
            item: OlMailItem = app.CreateItem(0)
            if self.account is not None:
                item.SendUsingAccount = self.account.ol_item
            item.SentOnBehalfOfName = address
        except COM_ERRORS as exc:
            raise OutlookError(f"Unable to create an email from {address!r}.") from exc
        mail = MailItem.from_outlook_item(item)
        if mail is None:
            raise OutlookError("Unable to create an Outlook email.")
        return mail

    def close(self) -> None:
        """Release held COM references without closing the Outlook process."""
        self.__dict__.pop("accounts", None)
        self.account = None
        self.address = None
        self._mapi = None
        self._app = None
