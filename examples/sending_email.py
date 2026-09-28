from ms_graph_toolbox.ms_graph_toolbox import ms_graph_toolbox
from ms_graph_toolbox.graph_email import send_email
from examples.logger import create_logger
import argparse

def test_send():
    """
    Parse client_id, client_secret and tenant_id from CLI and send a test email.
    Also, specify TO recipients as a comma-separated list.
    Additional optional arguments allow overriding sender/subject/body.
    """
    parser = argparse.ArgumentParser(description="Send test email via Microsoft Graph.")
    parser.add_argument("--client-id", required=True, help="Azure AD application (client) ID")
    parser.add_argument("--client-secret", required=True, help="Azure AD application client secret")
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID")
    parser.add_argument("--sender", required=True, help="Mailbox to send from")
    parser.add_argument("--to", required=True, dest="to_field", help="Comma-separated TO recipients")
    parser.add_argument("--subject", default="Test Email", help="Email subject")
    parser.add_argument("--content-type", default="Text", help="Email content type (e.g. HTML or Text)")
    parser.add_argument("--body", default="<h1>Test Body</h1>", help="Email body")
    
    test_attachment = {
        'content_bytes': b'test text attachment',  # Binary content
        'name': 'test.txt',                   # Filename
        'content_type': 'text/plain',         # MIME type
        'inline': False                       # Not an inline attachment
    }

    args = parser.parse_args()

    logger = create_logger()

    # Create Graph client object
    gph_object = ms_graph_toolbox(
        client_id = args.client_id,
        client_secret = args.client_secret,
        tenant_id = args.tenant_id,
        logger = logger
    )

    # Send test email (returns status code per implementation)
    result = send_email(
        gph_object = gph_object,
        subject = args.subject,
        content_type = args.content_type,
        body = args.body,
        sender = args.sender,
        to_field = args.to_field,
        cc_field = None,
        bcc_field = None,
        priority = "normal",
        attachments = [test_attachment]
    )

    if result == 0:
        logger.debug("Email send request accepted (202)")
    else:
        logger.error(f"Email send failed with code: {result}")

if __name__ == "__main__":
    test_send()