import datetime as dt

from pydantic import BaseModel

from asyncexchange.models.email import Mailbox


class CalendarMeeting(BaseModel):
    id: str
    change_key: str
    subject: str = ""
    text_body: str = ""
    html_body: str = ""
    start: dt.datetime
    end: dt.datetime
    location: str = ""
    is_all_day: bool = False
    busy_status: str = "Busy"
    organizer: Mailbox | None = None
    required_attendees: list[Mailbox] = []
    optional_attendees: list[Mailbox] = []
