import datetime as dt

from pydantic import BaseModel, Field


class Mailbox(BaseModel):
    email_address: str


class EmailMessage(BaseModel):
    id: str
    change_key: str
    subject: str = ""
    body: str = ""
    datetime_sent: dt.datetime
    is_read: bool = False
    from_: str | None = Field(default=None, alias="from")
    to: list[str] | None = None
    author: Mailbox | None = None
    to_recipients: list[Mailbox] = []
    text_body: str = ""

    class Config:
        populate_by_name = True


class MarkAsReadPayload(BaseModel):
    ids: list[str]
    