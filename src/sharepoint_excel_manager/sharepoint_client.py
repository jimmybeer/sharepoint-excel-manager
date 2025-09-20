"""
SharePoint client for interacting with Teams SharePoint sites
"""

import asyncio
from typing import List, Dict, Optional
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
import getpass


class SharePointClient:
    def __init__(self):
        self.context = None
        self.authenticated = False

    async def authenticate(self, site_url: str) -> bool:
        """Authenticate with SharePoint using user credentials"""
        try:
            # For now, we'll use interactive authentication
            # In production, you might want to use Azure AD app registration
            username = input("Enter your SharePoint username (email): ")
            password = getpass.getpass("Enter your password: ")

            # Create authentication context
            auth_context = AuthenticationContext(site_url)
            user_credentials = UserCredential(username, password)

            # Authenticate
            if auth_context.acquire_token_for_user(username, password):
                self.context = ClientContext(site_url, auth_context)
                self.authenticated = True
                return True
            else:
                return False

        except Exception as e:
            print(f"Authentication error: {e}")
            return False

    async def test_connection(self, team_url: str, folder_path: str = "") -> bool:
        """Test connection to SharePoint site"""
        try:
            if not self.authenticated:
                success = await self.authenticate(team_url)
                if not success:
                    return False

            # Test by getting site information
            web = self.context.web
            self.context.load(web)
            self.context.execute_query()

            return True

        except Exception as e:
            print(f"Connection test failed: {e}")
            return False

    async def get_excel_files(self, team_url: str, folder_path: str = "") -> List[Dict]:
        """Get list of Excel files from SharePoint folder"""
        try:
            if not self.authenticated:
                success = await self.authenticate(team_url)
                if not success:
                    raise Exception("Authentication failed")

            # Get document library
            if folder_path:
                folder = self.context.web.get_folder_by_server_relative_url(folder_path)
            else:
                # Default to "Shared Documents" if no path specified
                folder = self.context.web.default_document_library().root_folder

            # Load files
            files = folder.files
            self.context.load(files)
            self.context.execute_query()

            excel_files = []
            for file in files:
                # Filter for Excel files
                if file.name.endswith((".xlsx", ".xlsm", ".xls")):
                    excel_files.append(
                        {
                            "name": file.name,
                            "url": file.serverRelativeUrl,
                            "modified": str(file.time_last_modified),
                            "size": file.length,
                        }
                    )

            return excel_files

        except Exception as e:
            print(f"Error getting files: {e}")
            raise

    async def download_file(self, file_url: str, local_path: str) -> bool:
        """Download a file from SharePoint"""
        try:
            if not self.authenticated:
                raise Exception("Not authenticated")

            # Get file
            file = self.context.web.get_file_by_server_relative_url(file_url)

            # Download file content
            file_content = file.read()
            self.context.execute_query()

            # Save to local file
            with open(local_path, "wb") as local_file:
                local_file.write(file_content)

            return True

        except Exception as e:
            print(f"Error downloading file: {e}")
            return False

    async def upload_file(
        self, local_path: str, sharepoint_folder: str, filename: str
    ) -> bool:
        """Upload a file to SharePoint"""
        try:
            if not self.authenticated:
                raise Exception("Not authenticated")

            # Read local file
            with open(local_path, "rb") as local_file:
                file_content = local_file.read()

            # Get target folder
            folder = self.context.web.get_folder_by_server_relative_url(
                sharepoint_folder
            )

            # Upload file
            file = folder.upload_file(filename, file_content)
            self.context.execute_query()

            return True

        except Exception as e:
            print(f"Error uploading file: {e}")
            return False
