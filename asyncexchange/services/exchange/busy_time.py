import datetime as dt
import xml.etree.ElementTree as ET

import httpx

from asyncexchange.models.calendar import CalendarMeeting
from asyncexchange.services.exchange.base import AsyncExchangeBaseService
from asyncexchange.services.xml.calendar import CalendarXmlHelper


class BusyTimeService(AsyncExchangeBaseService):
    """
    Async calendar service that uses the EWS SOAP API to fetch meetings
    (busy time) from the authenticated user's calendar, or another
    mailbox's calendar / free-busy view.
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
        email: str | None = None,
    ) -> list[CalendarMeeting]:
        """
        Fetch calendar meetings that fall within ``[start, end)``.

        Without ``email``, reads the authenticated user's calendar via
        ``FindItem`` + ``GetItem`` (full details).

        With ``email``, first tries that mailbox's calendar folder for
        full meeting details. If calendar access is denied, falls back to
        ``GetUserAvailability`` (free/busy): busy blocks are returned, but
        subject/body/location/attendees are empty unless free/busy
        permissions grant limited details.
        """
        def _to_service_tz(value: dt.datetime) -> dt.datetime:
            if value.tzinfo is None:
                value = value.replace(tzinfo=self.tz)
            return value.astimezone(self.tz)

        start = _to_service_tz(start)
        end = _to_service_tz(end)

        if email:
            detailed = await self._try_get_calendar_meetings(
                start=start,
                end=end,
                max_entries=max_entries,
                email=email,
            )
            if detailed is not None:
                return detailed
            return await self._get_free_busy_meetings(
                start=start,
                end=end,
                email=email,
            )

        return await self._get_calendar_meetings(
            start=start,
            end=end,
            max_entries=max_entries,
            email=None,
        )

    async def _try_get_calendar_meetings(
        self,
        *,
        start: dt.datetime,
        end: dt.datetime,
        max_entries: int,
        email: str,
    ) -> list[CalendarMeeting] | None:
        """
        Attempt a calendar-folder read. Returns ``None`` when the caller
        lacks permission so the free/busy fallback can run.
        """
        try:
            body = CalendarXmlHelper.build_finditem_body(
                start=start,
                end=end,
                max_entries=max_entries,
                email=email,
            )
            root = await self._post_ews(
                soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
                body=body,
            )
        except httpx.HTTPStatusError:
            return None

        if CalendarXmlHelper.finditem_access_denied(root):
            return None

        return await self._finish_calendar_meetings(root)

    async def _get_calendar_meetings(
        self,
        *,
        start: dt.datetime,
        end: dt.datetime,
        max_entries: int,
        email: str | None,
    ) -> list[CalendarMeeting]:
        body = CalendarXmlHelper.build_finditem_body(
            start=start,
            end=end,
            max_entries=max_entries,
            email=email,
        )
        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
            body=body,
        )
        return await self._finish_calendar_meetings(root)

    async def _finish_calendar_meetings(
        self,
        finditem_root: ET.Element,
    ) -> list[CalendarMeeting]:
        basic_items = CalendarXmlHelper.parse_finditem_response(finditem_root)

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

    async def _get_free_busy_meetings(
        self,
        *,
        start: dt.datetime,
        end: dt.datetime,
        email: str,
    ) -> list[CalendarMeeting]:
        body = CalendarXmlHelper.build_getuseravailability_body(
            email=email,
            start=start,
            end=end,
        )
        root = await self._post_ews(
            soap_action=(
                "http://schemas.microsoft.com/exchange/services/2006/messages/"
                "GetUserAvailability"
            ),
            body=body,
        )
        result = CalendarXmlHelper.parse_getuseravailability_response(root)

        for item in result:
            item.start = item.start.astimezone(self.tz)
            item.end = item.end.astimezone(self.tz)

        return result
