from __future__ import annotations

from ..enums import ItemType
from ..protocols import OlAddressEntry
from ..utils import get_smtp_address
from .base import BaseModel


class AddressEntry(BaseModel):
    """Represent an Outlook address entry.

    Parameters
    ----------
    address_entry : OlAddressEntry
        Outlook address entry COM object to wrap.
    """

    item_name = "AddressEntry"
    item_type = ItemType.ADDRESS_ENTRY
    required_properties = (
        "Address",
        "Name",
        "Type",
        "PropertyAccessor",
        "AddressEntryUserType",
    )

    def __init__(self, address_entry: OlAddressEntry):
        """Initialize an address-entry wrapper.

        Parameters
        ----------
        address_entry : OlAddressEntry
            Outlook address-entry COM object.

        Returns
        -------
        None
        """
        super().__init__(address_entry)
        self._protocol = address_entry

    @property
    def name(self) -> str:
        """Returns the display name of the address entry."""
        return str(self._protocol.Name)

    @property
    def email_address(self) -> str:
        """Returns the email address of the address entry."""
        if address := get_smtp_address(self._protocol):
            return address
        return self._protocol.Address or ""

    @property
    def property_accessor(self):
        """Returns the PropertyAccessor object of the address entry."""
        return self._protocol.PropertyAccessor

    @property
    def user_type(self) -> int:
        """Returns the user type of the address entry."""
        return self._protocol.AddressEntryUserType
