from __future__ import annotations

from types import SimpleNamespace

from outlook import Folder, utils
from outlook.enums import ItemType
from outlook.tests.test_model_navigation import FakeCollection, FakeFolder, FakeMailItem


class OrderedCollection(FakeCollection):
    """Provide a fake collection that supports Outlook restrictions."""

    def Restrict(self, query: str) -> OrderedCollection:
        """Return unread messages for Outlook's unread restriction."""
        assert query == "[UnRead] = True"
        self.restricted = OrderedCollection(
            [item for item in self.items if getattr(item, "UnRead", False)]
        )
        return self.restricted

    def Sort(self, field: str, descending: bool) -> None:
        """Sort fake mail by received time."""
        super().Sort(field, descending)
        self.items.sort(key=lambda item: item.ReceivedTime, reverse=descending)


class FlakyCollection(FakeCollection):
    """Fail while resolving one collection entry."""

    def Item(self, index: int) -> object:
        """Raise an accessibility error for the third entry."""
        if index == 3:
            raise AttributeError
        return super().Item(index)


def test_folder_iteration_reads_only_requested_message() -> None:
    """Verify direct iteration does not materialize the entire folder."""
    items = OrderedCollection([FakeMailItem(str(index)) for index in range(10_000)])
    for index, item in enumerate(items.items):
        item.ReceivedTime = index

    first = next(iter(Folder(FakeFolder("Inbox", items=items))))

    assert first.subject == "9999"
    assert items.item_calls == 1


def test_unread_restriction_is_sorted_after_filtering() -> None:
    """Verify Outlook restriction retains received-time ordering."""
    old = FakeMailItem("old")
    old.ReceivedTime = 1
    read = FakeMailItem("read", unread=False)
    read.ReceivedTime = 3
    new = FakeMailItem("new")
    new.ReceivedTime = 2
    items = OrderedCollection([old, read, new])

    messages = Folder(FakeFolder("Inbox", items=items)).list_messages(unread_only=True)

    assert [message.subject for message in messages] == ["new", "old"]
    assert items.sort_calls == []
    assert items.restricted is not None
    assert items.restricted.sort_calls == [("[ReceivedTime]", True)]


def test_iteration_skips_inaccessible_collection_entries() -> None:
    """Verify non-mail and inaccessible COM entries do not stop enumeration."""
    first = FakeMailItem("first")
    second = FakeMailItem("second")
    items = FlakyCollection(
        [first, SimpleNamespace(Class=ItemType.FOLDER), object(), second]
    )

    messages = Folder(FakeFolder("Inbox", items=items)).list_messages()

    assert [message.subject for message in messages] == ["first", "second"]


def test_inaccessible_class_property_is_rejected(monkeypatch) -> None:
    """Verify a COM failure while reading Class is treated as inaccessible."""

    class ComFailure(Exception):
        """Represent a COM property-access error."""

    class BrokenItem:
        @property
        def Class(self) -> ItemType:
            """Raise while resolving the COM Class property."""
            raise ComFailure

    monkeypatch.setattr(utils, "COM_ERRORS", (ComFailure,))

    assert not utils.is_accessible_ol_item(BrokenItem(), ItemType.MAIL_ITEM)
