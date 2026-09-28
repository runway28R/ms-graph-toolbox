import requests
import base64
import os
import mimetypes


def send_email(gph_object, 
               subject, 
               content_type, 
               body, 
               sender, 
               to_field, 
               cc_field=None, 
               bcc_field=None, 
               priority="Normal", 
               attachments=None):
    """
    Send an email using Microsoft Graph on behalf of `sender`.

    Parameters:
        logger: logger instance to record actions (debug/error).
        subject: email subject string.
        content_type: one of 'text', 'html', 'text/plain', 'text/html', etc.
        body: email body string.
        to_field, cc_field, bcc_field: comma-separated recipient strings (e.g. "a@x.com, b@y.com").
        priority: message importance, e.g. "Low", "Normal", "High".
        attachments: optional list of attachment descriptors. Each descriptor may be:
            - {'path': 'C:\\full\\path\\file.pdf', 'name': 'file.pdf', 'content_type': 'application/pdf', 'inline': False}
            - {'content_bytes': b'...', 'name': 'image.png', 'content_type': 'image/png', 'inline': True, 'content_id': 'img1'}
            For inline images set 'inline': True and reference them in an HTML body as <img src="cid:content_id">.
            If content_type not provided it will be guessed from the filename.
            The code will base64-encode file contents as required by Graph.

    Returns:
        0 on success (202 response), 1 on exception, 2 if no access token, 3 on non-202 HTTP response.
    """
    try:
        # Ensure we have an access token before attempting to send
        if not gph_object.access_token:
            gph_object.logger.error("Invalid Access Token, email cannot be sent!")
            return 2

        # Parse recipient fields into Graph-friendly lists
        to_recipients = parse_recipients(to_field)
        cc_recipients = parse_recipients(cc_field)
        bcc_recipients = parse_recipients(bcc_field)

        # Normalize content type to one of Graph's expected values ("Text" or "HTML"), use "Text" as default
        valid_types = {"text": "Text", "plain": "Text", "html": "HTML", "text/plain": "Text", "text/html": "HTML"}
        content_type_normalized = valid_types.get(content_type.strip().lower(), "Text")

        # Process attachments and if any contain inline items but the body is not HTML, log a warning
        inline_found = False
        attachments_payload = []
        if attachments:
            # Process all attachments in one pass
            attachments_payload = [
                att for desc in attachments 
                if (att := build_attachment(descriptor=desc, logger=gph_object.logger)) is not None
            ]
            
            # Check for inline attachments
            inline_found = any(att.get("isInline") for att in attachments_payload)

            # Log warning if inline attachments found with non-HTML content
            if inline_found and content_type_normalized != "HTML":
                gph_object.logger.warning("Inline attachments present but body is not HTML. Inline images require HTML body and <img src=\"cid:contentId\"> references.")

        # Build sendMail endpoint for the configured sender user
        endpoint = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"

        # Compose message payload according to Graph sendMail schema
        email_msg = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": content_type_normalized,
                    "content": body
                },
                "toRecipients": to_recipients,
                "importance": priority
            }
        }

        # Attach cc/bcc only if provided
        if cc_recipients:
            email_msg["message"]["ccRecipients"] = cc_recipients
        if bcc_recipients:
            email_msg["message"]["bccRecipients"] = bcc_recipients

        # Attach attachments if any
        if attachments_payload:
            email_msg["message"]["attachments"] = attachments_payload

        # Set authorization header with Bearer token
        headers = {"Authorization": f"Bearer {gph_object.access_token}", "Content-Type": "application/json"}
        gph_object.logger.debug(f"Sending email using MS Graph from {sender}")
        response = requests.post(endpoint, headers=headers, json=email_msg)

        # 202 Accepted indicates Graph accepted the send request
        if response.status_code == 202:
            gph_object.logger.debug("Sending successful!")
            return 0
        else:
            # Log status and response body for debugging failures
            gph_object.logger.error(f"Sending Failed: {response.status_code}, {response.text}")
            return 3
    except Exception as e:
        # Any unexpected exception during send is logged
        gph_object.logger.error(f"Sending Failed: {e}")
        return 1


# Convert a comma-separated string of addresses into the Graph recipient JSON format.
def parse_recipients(field):
    if not field:
        return []
    return [{"emailAddress": {"address": addr.strip()}} for addr in field.split(",") if addr.strip()]


# Build Graph attachment payloads from descriptors
def build_attachment(descriptor, logger):
    """
    Accepts either:
        - {'path': 'C:\\file', ...}
        - {'content_bytes': b'...', ...}
    Returns a dict suitable for Graph message attachments.
    """
    try:
        # Basic check
        if "path" not in descriptor and "content_bytes" not in descriptor:
            logger.error(f"build_attachment failed: Descriptor missing 'path' or 'content_bytes': {descriptor}")
            return None
        # Determine name
        name = descriptor.get("name")
        content_type_guess = None

        # Load bytes from file path if provided
        if "path" in descriptor and descriptor["path"]:
            path = descriptor["path"]
            with open(path, "rb") as f:
                data = f.read()
            if not name:
                name = os.path.basename(path)
            content_type_guess = mimetypes.guess_type(path)[0]
        elif "content_bytes" in descriptor and descriptor["content_bytes"] is not None:
            data = descriptor["content_bytes"]
            if isinstance(data, str):
                data = data.encode("utf-8")

        # Determine contentType
        content_type = descriptor.get("content_type") or content_type_guess or "application/octet-stream"

        # Base64 encode content
        content_bytes_b64 = base64.b64encode(data).decode("utf-8")

        # Build attachment object (fileAttachment)
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name or "attachment",
            "contentType": content_type,
            "contentBytes": content_bytes_b64
        }

        # Inline settings for images
        if descriptor.get("inline"):
            attachment["isInline"] = True
            # contentId is used as cid reference in HTML body: <img src="cid:contentId">
            attachment["contentId"] = descriptor.get("content_id") or (name or "inline")
        return attachment
    except Exception as e:
        logger.error(f"build_attachment: failed for descriptor {descriptor}: {e}")
        return None