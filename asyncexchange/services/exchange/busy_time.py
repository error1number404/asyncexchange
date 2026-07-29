import datetime as dt

from asyncexchange.models.calendar import CalendarMeeting
from asyncexchange.services.exchange.base import AsyncExchangeBaseService
from asyncexchange.services.xml.calendar import CalendarXmlHelper


class BusyTimeService(AsyncExchangeBaseService):
    """
    Async calendar service that uses the EWS SOAP API to fetch meetings
    (busy time) from the authenticated user's calendar.
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

    async def get_meetings(
        self,
        *,
        start: dt.datetime,
        end: dt.datetime,
        max_entries: int = 1000,
    ) -> list[CalendarMeeting]:
        """
        Fetch calendar meetings that fall within ``[start, end)``.

        Uses EWS ``FindItem`` with a ``CalendarView``, then ``GetItem``
        for full details (attendees, body, etc.).
        """
        def _to_service_tz(value: dt.datetime) -> dt.datetime:
            if value.tzinfo is None:
                value = value.replace(tzinfo=self.tz)
            return value.astimezone(self.tz)

        start = _to_service_tz(start)
        end = _to_service_tz(end)

        body = CalendarXmlHelper.build_finditem_body(
            start=start,
            end=end,
            max_entries=max_entries,
        )
        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
            body=body,
        )
        basic_items = CalendarXmlHelper.parse_finditem_response(root)

        if not basic_items:
            return []

        getitem_body = CalendarXmlHelper.build_getitem_body(basic_items)
        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/GetItem",
            body=getitem_body,
        )
        result = CalendarXmlHelper.parse_getitem_response(root)

        for item in result:
            item.start = item.start.astimezone(self.tz)
            item.end = item.end.astimezone(self.tz)

            if item.organizer and item.organizer.email_address:
                resolved = await self.resolve_email_address(item.organizer.email_address)
                if resolved:
                    item.organizer.email_address = resolved

            for recipient in item.required_attendees + item.optional_attendees:
                if recipient.email_address:
                    resolved = await self.resolve_email_address(recipient.email_address)
                    if resolved:
                        recipient.email_address = resolved

        return result
