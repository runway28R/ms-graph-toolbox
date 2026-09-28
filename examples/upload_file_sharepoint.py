from ms_graph_toolbox.ms_graph_toolbox import ms_graph_toolbox
from ms_graph_toolbox.graph_sharepoint import graph_sharepoint
from examples.logger import create_logger
import argparse


def upload_file():
    """
    Parse client_id, client_secret and tenant_id from CLI and get users.
    Additional optional arguments allow filtering and selecting specific parameters.
    """
    parser = argparse.ArgumentParser(description="Upload a file to SharePoint using Microsoft Graph.")
    parser.add_argument("--client-id", required=True, help="Azure AD application (client) ID", type=str)
    parser.add_argument("--client-secret", required=True, help="Azure AD application client secret", type=str)
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID", type=str)

    parser.add_argument("--site_url", required=True, help="Sharepoint Site URL", type=str)
    parser.add_argument("--local_file_path", required=True, help="Local path of the file to be uploaded", type=str)
    parser.add_argument("--document_library", required=False, help="Root document library on the Sharepoint site", default="Documents", type=str)
    parser.add_argument("--folder_path", required=False, help="Path inside the document library", default="", type=str)

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

    graph_sharepoint_obj = graph_sharepoint(access_token=gph_object.access_token,
                                            logger=logger)

    siteid = graph_sharepoint_obj.get_site_id(site_url=args.site_url)
    if not siteid:
        logger.error(f"Site ID could not be retrieved for site URL: {args.site_url}")
        return

    all_drives = graph_sharepoint_obj.get_document_libraries(siteid)

    if all_drives:
        files_ok = 0
        files_nok = 0

        # Looking for the desired document library
        if args.document_library not in [d[1] for d in drive_id]:
            logger.error(f"Document library '{args.document_library}' was not found in document libraries on site '{args.site_url}'.")
            logger.debug(f"The following libraries were found:")
            for did, name in drive_id:
                logger.debug(f"{name}; Drive ID: {did}")
            return
        else:
            drive_ids = [d for d in all_drives if d[1] == args.document_library]
            drive_id = drive_ids[0][0]

            # Optional, print folder content
            content = graph_sharepoint_obj.get_folder_content(siteid, drive_id)
            graph_sharepoint_obj.print_folder_content(content)

            # Uploading the file to Sharepoint
            file_url, success = graph_sharepoint_obj.upload_file_graph(
                site_id=siteid,
                drive_id=drive_id,
                folder_path=args.folder_path,
                local_file_path=args.local_file_path
            )
            if success:
                files_ok += 1
                logger.info(f"File successfully uploaded to: {file_url}")
            else:
                files_nok += 1
                logger.error(f"File upload failed: {file_url}")
        logger.info(f"Upload summary: {files_ok} files uploaded successfully, {files_nok} files failed.")
    else:
        logger.error("No drive ID found; file upload failed.")


if __name__ == "__main__":
    upload_file()