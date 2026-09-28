from urllib.parse import quote
import pathlib as pl
import requests


class graph_sharepoint:
    def __init__(self, access_token:str, logger):
        self.access_token = access_token
        self.logger = logger


    def get_site_id(self, site_url:str):
        # Request site ID
        try:
            full_url = f'https://graph.microsoft.com/v1.0/sites/{site_url}'
            response = requests.get(full_url, 
                                    headers={'Authorization': f'Bearer {self.access_token}'})
            self.logger.debug(f"get_site_id response: {response.status_code} - {response.text}")
            return response.json().get('id')  # Return the site ID
        except Exception as e:
            self.logger.error(f"get_site_id failed: {e}")
            return None


    def get_document_libraries(self, site_id:str):
        # Retrieve drive IDs and names associated with a site
        try:
            drives_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
            response = requests.get(drives_url, headers={'Authorization': f'Bearer {self.access_token}'})
            drives = response.json().get('value', [])
            return [(drive['id'], drive['name']) for drive in drives]
        except Exception as e:
            self.logger.error(f"get_document_libraries failed: {e}")
            return None


    def get_folder_content(self, site_id:str, drive_id:str):
        # Get the contents of a folder
        try:
            folder_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children'
            response = requests.get(folder_url, headers={'Authorization': f'Bearer {self.access_token}'})
            return response.json().get('value', [])
        except Exception as e:
            self.logger.error(f"get_folder_content failed: {e}")
            return None

    def print_folder_content(self, folder_content):
        # Display the contents of a SharePoint folder
        folders = []
        files = []
        
        for item in folder_content:
            if "folder" in item:
                folders.append(item["name"])
            elif "file" in item:
                files.append(item["name"])

        print(f"Folders: {len(folders)}")
        for f in sorted(folders):
            print(f)

        print(f"Files: {len(files)}")
        for f in sorted(files):
            print(f)


    def upload_file_graph(self, site_id, drive_id:str, folder_path:str, local_file_path:str):
        """
        Uploads a file to a SharePoint folder using Microsoft Graph API.

        :param site_id: The SharePoint site ID
        :param drive_id: The drive ID for the desired folder
        :param folder_path: Path inside taht folder
        :param local_file_path: Local path of the file to be uploaded
        :return: (file_url, success_file_count)
        """
        try:
            file_path = pl.Path(local_file_path)
            if not file_path.is_file():
                raise FileNotFoundError(f"File not found: {file_path}")

            with open(file_path, "rb") as file:
                file_content = file.read()

            # Build the upload URL
            folder_path_encoded = quote(folder_path)
            upload_url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path_encoded}/{file_path.name}:/content"
            )

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/octet-stream"
            }

            response = requests.put(upload_url, headers=headers, data=file_content)
            if response.status_code in [200, 201]:
                file_info = response.json()
                file_url = file_info.get("webUrl", "")
                self.logger.debug(f"File uploaded to: {file_url}")
                return file_url, 1
            else:
                self.logger.error(f"Upload to {upload_url} failed: {response.text}")
                return response.text, 0
        except FileNotFoundError as e:
            return str(e), 0
        except Exception as e:
            self.logger.error(f"upload_file_graph failed: {e}")
            return str(e), 0
