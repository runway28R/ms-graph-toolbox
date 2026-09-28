"""
Basic code to obtain an application token via MSAL to be used with Microsoft Graph.

"""
import msal


class ms_graph_toolbox:
    """
    Wrapper class to handle app-only authentication and sending email via Microsoft Graph.

    Attributes:
        logger: Logger with .debug/.info/.warning/.error methods for logging.
        sender: The user (email) that will be used as the "From" / mailbox to send from.
        client_id: Azure AD Application (client) ID used for the OAuth2 client credentials flow.
        client_secret: Confidential value (application secret) for the Azure AD app used to authenticate the app. 
        tenant_id: Azure AD tenant identifier (GUID) or tenant domain used to build the authority URL
    """

    def __init__(self, 
                 client_id, 
                 client_secret, 
                 tenant_id, 
                 logger):
        
        # Set logger and access token placeholder
        self.logger = logger
        self.access_token = None

        try:
            # Build authority URL for tenant
            authority = f"https://login.microsoftonline.com/{tenant_id}"
            # Use .default scope for client credentials to get app-level permissions
            scopes = ["https://graph.microsoft.com/.default"]

            # Create MSAL confidential client app using client credentials
            app = msal.ConfidentialClientApplication(
                client_id,
                authority=authority,
                client_credential=client_secret
            )

            # Acquire token for client (app-only)
            result = app.acquire_token_for_client(scopes)

            # If token obtained, store it. Otherwise log the error.
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.debug("Successfully obtained Graph API token.")
            else:
                # error_description may contain helpful details about why token request failed
                error_msg = result.get("error_description", str(result))
                logger.error(f"Failed to get token: {error_msg}")

        except Exception as e:
            # Catch-all to ensure initialization failure is logged
            logger.error(f"graph_emailer initialization failed: {e}")
