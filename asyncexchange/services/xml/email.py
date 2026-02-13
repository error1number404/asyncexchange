import datetime as dt
import xml.etree.ElementTree as ET
from typing import Iterable, List

from asyncexchange.models.email import EmailMessage, Mailbox

EWS_NS = {
    "s": "http://schemas.xmlsoap.org/soap/envelope/",
    "m": "http://schemas.microsoft.com/exchange/services/2006/messages",
    "t": "http://schemas.microsoft.com/exchange/services/2006/types",
}


class EwsXmlHelper:
    """
    Helper for building and parsing EWS SOAP XML payloads.
    """

    @staticmethod
    def build_soap_envelope(body: str) -> str:
        """
        Wrap a raw EWS body fragment into a full SOAP envelope.
        """
        return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    {body}
  </soap:Body>
</soap:Envelope>"""

    @staticmethod
    def build_finditem_body(
        *,
        end: dt.datetime | None = None,
        start: dt.datetime | None = None,
        is_read: bool | None = None,
    ) -> str:
        """
        Build the EWS ``FindItem`` request body for the Inbox with
        """
        conditions: list[str] = []

        # EWS "True" / "False" are capitalised strings.
        if is_read is not None:
            conditions.append(
                f"""
              <t:IsEqualTo>
                <t:FieldURI FieldURI="message:IsRead" />
                <t:FieldURIOrConstant>
                  <t:Constant Value="{str(is_read).lower().capitalize()}" />
                </t:FieldURIOrConstant>
              </t:IsEqualTo>
            """
            )

        if start is not None and end is not None:
            conditions.append(
                f"""
              <t:And>
                <t:IsGreaterThanOrEqualTo>
                  <t:FieldURI FieldURI="item:DateTimeSent" />
                  <t:FieldURIOrConstant>
                    <t:Constant Value="{start.isoformat()}" />
                  </t:FieldURIOrConstant>
                </t:IsGreaterThanOrEqualTo>
                <t:IsLessThanOrEqualTo>
                  <t:FieldURI FieldURI="item:DateTimeSent" />
                  <t:FieldURIOrConstant>
                    <t:Constant Value="{end.isoformat()}" />
                  </t:FieldURIOrConstant>
                </t:IsLessThanOrEqualTo>
              </t:And>
            """
            )

        # If there is more than one condition, wrap them in a single <t:And>.
        restriction_inner = ""
        if conditions:
            if len(conditions) == 1:
                restriction_inner = conditions[0]
            else:
                restriction_inner = f"""
              <t:And>
                {''.join(conditions)}
              </t:And>
            """

        restriction_block = ""
        if restriction_inner:
            restriction_block = f"""
      <m:Restriction>
        {restriction_inner}
      </m:Restriction>
      """

        return f"""
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="message:From" />
        </t:AdditionalProperties>
      </m:ItemShape>
      {restriction_block}
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox" />
      </m:ParentFolderIds>
    </m:FindItem>
        """

    @staticmethod
    def build_getitem_body(messages: Iterable[EmailMessage]) -> str:
        """
        Build the EWS ``GetItem`` request body to fetch full message
        details (including recipients and body) for the given messages.
        """
        item_ids_xml = ""
        for msg in messages:
            if not msg.id:
                continue
            change_key_attr = f' ChangeKey="{msg.change_key}"' if msg.change_key else ""
            item_ids_xml += f"""
      <t:ItemId Id="{msg.id}"{change_key_attr} />"""

        return f"""
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:CcRecipients" />
          <t:FieldURI FieldURI="message:BccRecipients" />
          <t:FieldURI FieldURI="item:Body" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        {item_ids_xml}
      </m:ItemIds>
    </m:GetItem>
        """

    @staticmethod
    def build_resolvenames_body(unresolved_entry: str) -> str:
        """
        Build the EWS ``ResolveNames`` request body to resolve a legacy
        distinguished name (legacyDN/X.500) or other ambiguous value to
        a directory object (typically yielding an SMTP address).
        """
        return f"""
    <m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">
      <m:UnresolvedEntry>{unresolved_entry}</m:UnresolvedEntry>
    </m:ResolveNames>
        """

    @staticmethod
    def _parse_messages_common(root: ET.Element) -> List[EmailMessage]:
        """
        Internal helper to parse SOAP responses (``FindItem`` / ``GetItem``)
        into a list of ``EmailMessage`` objects.
        """
        messages: List[EmailMessage] = []

        for item in root.findall(".//t:Message", EWS_NS):
            item_id_el = item.find("t:ItemId", EWS_NS)
            subject_el = item.find("t:Subject", EWS_NS)
            body_el = item.find("t:Body", EWS_NS)
            is_read_el = item.find("t:IsRead", EWS_NS)
            dt_sent_el = item.find("t:DateTimeSent", EWS_NS)
            from_el = item.find("t:From/t:Mailbox/t:EmailAddress", EWS_NS)
            to_recips = [
                e.text or ""
                for e in item.findall(
                    "t:ToRecipients/t:Mailbox/t:EmailAddress",
                    EWS_NS,
                )
            ]

            if item_id_el is None or dt_sent_el is None:
                continue

            sent_raw = dt_sent_el.text or ""
            if sent_raw.endswith("Z"):
                sent_raw = sent_raw.replace("Z", "+00:00")

            author_email = from_el.text if from_el is not None and from_el.text else ""

            msg = EmailMessage(
                id=item_id_el.attrib.get("Id", ""),
                change_key=item_id_el.attrib.get("ChangeKey", ""),
                subject=subject_el.text if subject_el is not None and subject_el.text else "",
                body=body_el.text if body_el is not None and body_el.text else "",
                datetime_sent=dt.datetime.fromisoformat(sent_raw),
                is_read=is_read_el.text.lower() == "true" if is_read_el is not None and is_read_el.text else False,
                from_=author_email or None,
                to=to_recips or None,
                author=Mailbox(email_address=author_email) if author_email else None,
                to_recipients=[Mailbox(email_address=e) for e in to_recips],
                text_body=body_el.text if body_el is not None and body_el.text else "",
            )
            messages.append(msg)

        return messages

    @staticmethod
    def parse_finditem_response(root: ET.Element) -> List[EmailMessage]:
        """
        Parse a ``FindItem`` SOAP response into a list of ``EmailMessage`` objects.
        """
        return EwsXmlHelper._parse_messages_common(root)

    @staticmethod
    def parse_getitem_response(root: ET.Element) -> List[EmailMessage]:
        """
        Parse a ``GetItem`` SOAP response into a list of ``EmailMessage`` objects.
        """
        return EwsXmlHelper._parse_messages_common(root)

    @staticmethod
    def parse_resolvenames_response(root: ET.Element) -> str | None:
        """
        Parse a ``ResolveNames`` SOAP response and return the first SMTP
        email address found, if any.
        """
        print(root.text)
        # Typical structure:
        # <m:ResolveNamesResponseMessage>
        #   <m:ResolutionSet>
        #     <t:Resolution>
        #       <t:Mailbox>
        #         <t:Name>Display Name</t:Name>
        #         <t:EmailAddress>user@example.com</t:EmailAddress>
        #         <t:RoutingType>SMTP</t:RoutingType>
        #       </t:Mailbox>
        #     </t:Resolution>
        #   </m:ResolutionSet>
        # </m:ResolveNamesResponseMessage>
        for resolution in root.findall(".//t:Resolution", EWS_NS):
            mailbox_el = resolution.find("t:Mailbox", EWS_NS)
            if mailbox_el is None:
                continue

            email_el = mailbox_el.find("t:EmailAddress", EWS_NS)
            routing_el = mailbox_el.find("t:RoutingType", EWS_NS)

            if email_el is None or not email_el.text:
                continue

            routing_type = (routing_el.text or "").upper() if routing_el is not None and routing_el.text else ""

            # Prefer SMTP, but if RoutingType is missing assume it's fine.
            if not routing_type or routing_type == "SMTP":
                return email_el.text

        return None

    @staticmethod
    def build_updateitem_body(messages: Iterable[EmailMessage]) -> str:
        """
        Build the EWS ``UpdateItem`` request body that marks the given
        messages as read.
        """
        updates_xml = ""
        for msg in messages:
            updates_xml += f"""
        <m:ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="{msg.id}" ChangeKey="{msg.change_key}" />
            <t:Updates>
              <t:SetItemField>
                <t:FieldURI FieldURI="message:IsRead" />
                <t:Message>
                  <t:IsRead>true</t:IsRead>
                </t:Message>
              </t:SetItemField>
            </t:Updates>
          </t:ItemChange>
        </m:ItemChanges>
            """

        return f"""
    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AutoResolve">
      {updates_xml}
    </m:UpdateItem>
        """
