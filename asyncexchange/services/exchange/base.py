import datetime as dt
import xml.etree.ElementTree as ET

import httpx

from asyncexchange.services.xml.email import EwsXmlHelper


class AsyncExchangeBaseService:
    """
    Base class for async services that talk to Exchange via EWS over HTTP.
    Provides HTTP client setup, timezone handling and a low-level EWS caller.
    """

    def __init__(
        self,
        username: str,
        password: str,
        server_url: str,
        tz: dt.tzinfo | None = None,
    ) -> None:
        self.username = username
        self.password = password

        self.client = httpx.AsyncClient(
            base_url=server_url,
            auth=(self.username, self.password),
            timeout=10.0,
        )
        self.tz: dt.tzinfo = tz or dt.timezone(
            dt.timedelta(hours=0),
            "UTC",
        )

    async def aclose(self) -> None:
        await self.client.aclose()

    async def _post_ews(self, soap_action: str, body: str) -> ET.Element:
        """
        Perform a raw EWS SOAP call against the standard EWS endpoint:
        https://<server_url>/EWS/Exchange.asmx
        """
        envelope = EwsXmlHelper.build_soap_envelope(body)

        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": soap_action,
        }

        response = await self.client.post(
            "/EWS/Exchange.asmx",
            content=envelope.encode("utf-8"),
            headers=headers,
        )
        response.raise_for_status()
        return ET.fromstring(response.text)

    async def resolve_email_address(self, unresolved_entry: str) -> str | None:
        """
        Resolve a legacy distinguished name (legacyDN/X.500) or other
        ambiguous identifier to an SMTP email address using the EWS
        ``ResolveNames`` operation.

        Returns the resolved SMTP address, or ``None`` if resolution
        fails or no SMTP address is found.
        """
        body = EwsXmlHelper.build_resolvenames_body(unresolved_entry)

        root = await self._post_ews(
            soap_action="http://schemas.microsoft.com/exchange/services/2006/messages/ResolveNames",
            body=body,
        )

        return EwsXmlHelper.parse_resolvenames_response(root)
