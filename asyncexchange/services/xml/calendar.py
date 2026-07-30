import datetime as dt
import html
import xml.etree.ElementTree as ET
from collections.abc import Iterable

from asyncexchange.models.calendar import CalendarMeeting
from asyncexchange.models.email import Mailbox
from asyncexchange.services.xml.email import EWS_NS, EwsXmlHelper

# Response codes that mean the caller cannot open the target calendar folder.
_CALENDAR_ACCESS_ERRORS = frozenset(
    {
        "ErrorAccessDenied",
        "ErrorFolderNotFound",
        "ErrorItemNotFound",
        "ErrorNonExistentMailbox",
        "ErrorImpersonateUserDenied",
        "ErrorImpersonationFailed",
        "ErrorMailboxStoreUnavailable",
    }
)


class CalendarXmlHelper:
    """
    Helper for building and parsing EWS SOAP XML payloads for calendar items.
    """

    @staticmethod
    def build_finditem_body(
        *,
        start: dt.datetime,
        end: dt.datetime,
        max_entries: int = 1000,
        email: str | None = None,
    ) -> str:
        """
        Build the EWS ``FindItem`` request body for the Calendar folder
        using a ``CalendarView`` over ``[start, end)``.

        When ``email`` is set, the request targets that mailbox's calendar
        (requires appropriate permissions); otherwise the authenticated
        user's calendar is used.
        """
        start_utc = EwsXmlHelper.ews_datetime_utc(start)
        end_utc = EwsXmlHelper.ews_datetime_utc(end)

        if email:
            folder_id = f"""
        <t:DistinguishedFolderId Id="calendar">
          <t:Mailbox>
            <t:EmailAddress>{html.escape(email, quote=True)}</t:EmailAddress>
          </t:Mailbox>
        </t:DistinguishedFolderId>"""
        else:
            folder_id = """
        <t:DistinguishedFolderId Id="calendar" />"""

        return f"""
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="calendar:Start" />
          <t:FieldURI FieldURI="calendar:End" />
          <t:FieldURI FieldURI="calendar:Location" />
          <t:FieldURI FieldURI="calendar:IsAllDayEvent" />
          <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus" />
          <t:FieldURI FieldURI="calendar:Organizer" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView MaxEntriesReturned="{max_entries}" StartDate="{start_utc}" EndDate="{end_utc}" />
      <m:ParentFolderIds>
        {folder_id}
      </m:ParentFolderIds>
    </m:FindItem>
        """

    @staticmethod
    def finditem_access_denied(root: ET.Element) -> bool:
        """
        Return True when a ``FindItem`` response indicates the caller
        cannot open the target calendar folder.
        """
        for msg in root.findall(".//m:FindItemResponseMessage", EWS_NS):
            if msg.attrib.get("ResponseClass") != "Error":
                continue
            code_el = msg.find("m:ResponseCode", EWS_NS)
            code = code_el.text if code_el is not None and code_el.text else ""
            if code in _CALENDAR_ACCESS_ERRORS:
                return True
        return False

    @staticmethod
    def build_getuseravailability_body(
        *,
        email: str,
        start: dt.datetime,
        end: dt.datetime,
        merged_free_busy_interval: int = 30,
    ) -> str:
        """
        Build an EWS ``GetUserAvailability`` request for free/busy
        (and limited details when the free/busy ACL allows it).
        """
        start_utc = EwsXmlHelper.ews_datetime_utc(start)
        end_utc = EwsXmlHelper.ews_datetime_utc(end)
        # GetUserAvailability TimeWindow is interpreted in the request TimeZone;
        # use naive UTC wall times (no Z) with Bias=0.
        start_local = start_utc.rstrip("Z")
        end_local = end_utc.rstrip("Z")
        safe_email = html.escape(email, quote=True)

        # Bias is minutes west of UTC. StandardTime/DaylightTime must use
        # different months — identical transitions make Exchange reject the TZ.
        return f"""
    <m:GetUserAvailabilityRequest>
      <t:TimeZone>
        <t:Bias>0</t:Bias>
        <t:StandardTime>
          <t:Bias>0</t:Bias>
          <t:Time>00:00:00</t:Time>
          <t:DayOrder>1</t:DayOrder>
          <t:Month>11</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:StandardTime>
        <t:DaylightTime>
          <t:Bias>0</t:Bias>
          <t:Time>00:00:00</t:Time>
          <t:DayOrder>1</t:DayOrder>
          <t:Month>3</t:Month>
          <t:DayOfWeek>Sunday</t:DayOfWeek>
        </t:DaylightTime>
      </t:TimeZone>
      <m:MailboxDataArray>
        <t:MailboxData>
          <t:Email>
            <t:Address>{safe_email}</t:Address>
          </t:Email>
          <t:AttendeeType>Required</t:AttendeeType>
          <t:ExcludeConflicts>false</t:ExcludeConflicts>
        </t:MailboxData>
      </m:MailboxDataArray>
      <t:FreeBusyViewOptions>
        <t:TimeWindow>
          <t:StartTime>{start_local}</t:StartTime>
          <t:EndTime>{end_local}</t:EndTime>
        </t:TimeWindow>
        <t:MergedFreeBusyIntervalInMinutes>{merged_free_busy_interval}</t:MergedFreeBusyIntervalInMinutes>
        <t:RequestedView>DetailedMerged</t:RequestedView>
      </t:FreeBusyViewOptions>
    </m:GetUserAvailabilityRequest>
        """

    @staticmethod
    def parse_getuseravailability_response(root: ET.Element) -> list[CalendarMeeting]:
        """
        Parse ``GetUserAvailability`` into ``CalendarMeeting`` objects.

        With free/busy-only permissions, only start/end/busy status are
        populated. With limited details, subject/location may also appear;
        body and attendees are never returned by this operation.
        """
        meetings: list[CalendarMeeting] = []
        types_ns = EWS_NS["t"]
        messages_ns = EWS_NS["m"]

        def _child(parent: ET.Element, name: str) -> ET.Element | None:
            el = parent.find(f"{{{types_ns}}}{name}")
            if el is not None:
                return el
            return parent.find(f"{{{messages_ns}}}{name}")

        events = root.findall(f".//{{{types_ns}}}CalendarEvent")
        if not events:
            events = root.findall(f".//{{{messages_ns}}}CalendarEvent")

        for event in events:
            start_el = _child(event, "StartTime")
            end_el = _child(event, "EndTime")
            if start_el is None or end_el is None:
                continue

            start = CalendarXmlHelper._parse_datetime(start_el.text)
            end = CalendarXmlHelper._parse_datetime(end_el.text)
            if start is None or end is None:
                continue

            # Request TimeZone uses Bias=0 (UTC). Naive wall times from the
            # response are UTC — attach UTC so later astimezone() converts.
            if start.tzinfo is None:
                start = start.replace(tzinfo=dt.UTC)
            if end.tzinfo is None:
                end = end.replace(tzinfo=dt.UTC)

            busy_el = _child(event, "BusyType")
            busy_status = (
                busy_el.text if busy_el is not None and busy_el.text else "Busy"
            )
            if busy_status == "Free":
                continue

            details = _child(event, "CalendarEventDetails")
            event_id = ""
            subject = ""
            location = ""
            if details is not None:
                id_el = _child(details, "ID")
                subject_el = _child(details, "Subject")
                location_el = _child(details, "Location")
                event_id = id_el.text if id_el is not None and id_el.text else ""
                subject = (
                    subject_el.text if subject_el is not None and subject_el.text else ""
                )
                location = (
                    location_el.text
                    if location_el is not None and location_el.text
                    else ""
                )

            meetings.append(
                CalendarMeeting(
                    id=event_id,
                    change_key="",
                    subject=subject,
                    start=start,
                    end=end,
                    location=location,
                    busy_status=busy_status,
                )
            )

        return meetings

    @staticmethod
    def build_getitem_body(meetings: Iterable[CalendarMeeting]) -> str:
        """
        Build the EWS ``GetItem`` request body to fetch full meeting
        details (attendees, body, etc.) for the given calendar items.
        """
        item_ids_xml = ""
        for meeting in meetings:
            if not meeting.id:
                continue
            change_key_attr = (
                f' ChangeKey="{meeting.change_key}"' if meeting.change_key else ""
            )
            item_ids_xml += f"""
      <t:ItemId Id="{meeting.id}"{change_key_attr} />"""

        return f"""
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:Body" />
          <t:FieldURI FieldURI="calendar:Start" />
          <t:FieldURI FieldURI="calendar:End" />
          <t:FieldURI FieldURI="calendar:Location" />
          <t:FieldURI FieldURI="calendar:IsAllDayEvent" />
          <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus" />
          <t:FieldURI FieldURI="calendar:Organizer" />
          <t:FieldURI FieldURI="calendar:RequiredAttendees" />
          <t:FieldURI FieldURI="calendar:OptionalAttendees" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        {item_ids_xml}
      </m:ItemIds>
    </m:GetItem>
        """

    @staticmethod
    def _parse_datetime(value: str | None) -> dt.datetime | None:
        if not value:
            return None
        raw = value
        if raw.endswith("Z"):
            raw = raw.replace("Z", "+00:00")
        return dt.datetime.fromisoformat(raw)

    @staticmethod
    def _parse_mailbox(mailbox_el: ET.Element | None) -> Mailbox | None:
        if mailbox_el is None:
            return None
        email_el = mailbox_el.find("t:EmailAddress", EWS_NS)
        if email_el is None or not email_el.text:
            return None
        return Mailbox(email_address=email_el.text)

    @staticmethod
    def _parse_attendees(item: ET.Element, path: str) -> list[Mailbox]:
        attendees: list[Mailbox] = []
        for attendee_el in item.findall(path, EWS_NS):
            mailbox = CalendarXmlHelper._parse_mailbox(
                attendee_el.find("t:Mailbox", EWS_NS)
            )
            if mailbox is not None:
                attendees.append(mailbox)
        return attendees

    @staticmethod
    def _parse_calendar_items(root: ET.Element) -> list[CalendarMeeting]:
        meetings: list[CalendarMeeting] = []

        for item in root.findall(".//t:CalendarItem", EWS_NS):
            item_id_el = item.find("t:ItemId", EWS_NS)
            start_el = item.find("t:Start", EWS_NS)
            end_el = item.find("t:End", EWS_NS)
            if item_id_el is None or start_el is None or end_el is None:
                continue

            start = CalendarXmlHelper._parse_datetime(start_el.text)
            end = CalendarXmlHelper._parse_datetime(end_el.text)
            if start is None or end is None:
                continue

            subject_el = item.find("t:Subject", EWS_NS)
            location_el = item.find("t:Location", EWS_NS)
            is_all_day_el = item.find("t:IsAllDayEvent", EWS_NS)
            busy_el = item.find("t:LegacyFreeBusyStatus", EWS_NS)
            body_el = item.find("t:Body", EWS_NS)
            organizer_el = item.find("t:Organizer/t:Mailbox", EWS_NS)

            body_text = body_el.text if body_el is not None and body_el.text else ""
            body_type = (
                body_el.attrib.get("BodyType", "").lower() if body_el is not None else ""
            )
            if body_type == "html":
                html_body = body_text
                text_body = EwsXmlHelper._html_to_text(body_text)
            else:
                html_body = ""
                text_body = body_text

            meetings.append(
                CalendarMeeting(
                    id=item_id_el.attrib.get("Id", ""),
                    change_key=item_id_el.attrib.get("ChangeKey", ""),
                    subject=(
                        subject_el.text
                        if subject_el is not None and subject_el.text
                        else ""
                    ),
                    text_body=text_body,
                    html_body=html_body,
                    start=start,
                    end=end,
                    location=(
                        location_el.text
                        if location_el is not None and location_el.text
                        else ""
                    ),
                    is_all_day=(
                        is_all_day_el.text.lower() == "true"
                        if is_all_day_el is not None and is_all_day_el.text
                        else False
                    ),
                    busy_status=(
                        busy_el.text
                        if busy_el is not None and busy_el.text
                        else "Busy"
                    ),
                    organizer=CalendarXmlHelper._parse_mailbox(organizer_el),
                    required_attendees=CalendarXmlHelper._parse_attendees(
                        item,
                        "t:RequiredAttendees/t:Attendee",
                    ),
                    optional_attendees=CalendarXmlHelper._parse_attendees(
                        item,
                        "t:OptionalAttendees/t:Attendee",
                    ),
                )
            )

        return meetings

    @staticmethod
    def parse_finditem_response(root: ET.Element) -> list[CalendarMeeting]:
        """
        Parse a ``FindItem`` SOAP response into a list of ``CalendarMeeting`` objects.
        """
        return CalendarXmlHelper._parse_calendar_items(root)

    @staticmethod
    def parse_getitem_response(root: ET.Element) -> list[CalendarMeeting]:
        """
        Parse a ``GetItem`` SOAP response into a list of ``CalendarMeeting`` objects.
        """
        return CalendarXmlHelper._parse_calendar_items(root)
