import datetime as dt
from typing import Iterable, List

from asyncexchange.models.email import EmailMessage
from asyncexchange.services.exchange.base import AsyncExchangeBaseService
from asyncexchange.services.xml.email import EwsXmlHelper


class EmailService(AsyncExchangeBaseService):
    """
    Async email service that uses the EWS SOAP API to interact with the Exchange server.
    """

    def __init__(
        self,
        username: str,
        password: str,
        server_url: str,
        tz: dt.tzinfo | None = None,
    ) -> None:
        super().__init__(
            username=username,
            password=password,
            server_url=server_url,
            tz=tz,
        )

    async def get_messages(
        self,
        *,
        end: dt.datetime | None = None,
        start: dt.datetime | None = None,
        is_read: bool | None = None,
    ) -> List[EmailMessage]:
        """
        Fetch messages from the Exchange server.
        """
        # First, use FindItem to get message IDs and basic metadata.
        body = EwsXmlHelper.build_finditem_body(start=start, end=end, is_read=is_read)
        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
            body=body,
        )
        basic_items = EwsXmlHelper.parse_finditem_response(root)

        # If nothing was found, return early.
        if not basic_items:
            return []

        # Then, use GetItem to fetch full details (recipients, body, etc.)
        getitem_body = EwsXmlHelper.build_getitem_body(basic_items)
        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/GetItem",
            body=getitem_body,
        )
        result = EwsXmlHelper.parse_getitem_response(root)
        for item in result:
            if item.author and item.author.email_address:
                resolved = await self.resolve_email_address(item.author.email_address)
                if resolved:
                    item.author.email_address = resolved

            for recipient in item.to_recipients:
                if recipient.email_address:
                    resolved = await self.resolve_email_address(recipient.email_address)
                    if resolved:
                        recipient.email_address = resolved
        return result
    async def mark_as_read(self, messages: Iterable[EmailMessage]) -> None:
        """
        Mark messages as read on the Exchange server.
        """
        message_list = [m for m in messages if m.id and m.change_key]
        if not message_list:
            return

        body = EwsXmlHelper.build_updateitem_body(message_list)

        await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem",
            body=body,
        )
