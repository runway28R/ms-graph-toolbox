from ms_graph_toolbox.ms_graph_toolbox import ms_graph_toolbox
from ms_graph_toolbox.graph_users import get_users
from examples.logger import create_logger
import argparse


def test_users():
    """
    Parse client_id, client_secret and tenant_id from CLI and get users.
    Additional optional arguments allow filtering and selecting specific parameters.
    """
    parser = argparse.ArgumentParser(description="Get users' details via Microsoft Graph.")
    parser.add_argument("--client-id", required=True, help="Azure AD application (client) ID")
    parser.add_argument("--client-secret", required=True, help="Azure AD application client secret")
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID")

    parser.add_argument("--select_data", help="Comma-separated user properties to return (e.g. displayName,mail,jobTitle)")
    parser.add_argument("--search_name", help="Filter users by display name (startswith)")
    parser.add_argument("--search_title", help="Filter users by job title (startswith)")
    parser.add_argument("--search_email", help="Filter users by email address (startswith)")
    parser.add_argument("--search_alias", help="Filter users by alias (startswith)")
    parser.add_argument("--search_company", help="Filter users by company name (startswith)")

    args = parser.parse_args()

    logger = create_logger()

    # Create Graph client object
    gph_object = ms_graph_toolbox(
        client_id = args.client_id,
        client_secret = args.client_secret,
        tenant_id = args.tenant_id,
        logger = logger
    )

    if gph_object.access_token is None:
        logger.error("Cannot proceed without a valid access token.")
        return

    users = get_users(
        gph_object,
        select_data=args.select_data,
        search_name=args.search_name,
        search_title=args.search_title,
        search_email=args.search_email,
        search_alias=args.search_alias,
        search_company=args.search_company
    )

    if users is not None:
        logger.info(f"Found {len(users)} user(s) matching criteria.")
        for user in users:
            logger.debug(f"{user}")
    else:
        logger.error("Failed to retrieve users.")


if __name__ == "__main__":
    test_users()