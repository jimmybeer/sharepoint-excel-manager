"""
SharePoint client for interacting with Teams SharePoint sites
Uses modern authentication methods compatible with Conditional Access policies
"""
import asyncio
import json
import logging
import os
import urllib.parse
import webbrowser
from pathlib import Path
from typing import Dict, List, Optional

import requests
from msal import PublicClientApplication

logger = logging.getLogger(__name__)


class SharePointClient:
    def __init__(self):
        self.access_token = None
        self.authenticated = False
        self.site_url = None
        
        # MSAL configuration for SharePoint
        self.client_id = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Office 365 Management Shell
        self.authority = "https://login.microsoftonline.com/common"
        self.scope = ["https://graph.microsoft.com/.default"]
        
        # Initialize MSAL app
        self.app = PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority
        )
    
    async def authenticate(self, site_url: str) -> bool:
        """Authenticate using MSAL with device code flow or interactive login"""
        try:
            self.site_url = site_url
            
            # Try to get token silently first (from cache)
            accounts = self.app.get_accounts()
            if accounts:
                logger.info("Found cached account, attempting silent authentication...")
                result = self.app.acquire_token_silent(
                    scopes=self.scope,
                    account=accounts[0]
                )
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    self.authenticated = True
                    logger.info("Silent authentication successful")
                    return True
            
            # If silent auth fails, try interactive authentication
            logger.info("Attempting interactive authentication...")
            result = self.app.acquire_token_interactive(
                scopes=self.scope,
                prompt="select_account"  # Allow user to select account
            )
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                self.authenticated = True
                logger.info("Interactive authentication successful")
                return True
            else:
                error_msg = result.get("error_description", "Unknown authentication error")
                logger.error(f"Authentication failed: {error_msg}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {e}")
            return False
    
    async def authenticate_device_code(self, site_url: str) -> bool:
        """Alternative authentication using device code flow (for headless environments)"""
        try:
            self.site_url = site_url
            
            # Initiate device code flow
            flow = self.app.initiate_device_flow(scopes=self.scope)
            
            if "user_code" not in flow:
                raise Exception("Failed to create device flow")
            
            print("\n" + "="*60)
            print("DEVICE CODE AUTHENTICATION")
            print("="*60)
            print(flow["message"])
            print("="*60)
            print("Please complete the authentication in your browser.")
            print("This window will wait for you to complete the process...")
            print("="*60 + "\n")
            
            # Complete the device code flow
            result = self.app.acquire_token_by_device_flow(flow)
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                self.authenticated = True
                logger.info("Device code authentication successful")
                return True
            else:
                error_msg = result.get("error_description", "Unknown authentication error")
                logger.error(f"Device code authentication failed: {error_msg}")
                return False
                
        except Exception as e:
            logger.error(f"Device code authentication error: {e}")
            return False
    
    def _get_site_id_from_url(self, site_url: str) -> Optional[str]:
        """Extract site ID from SharePoint URL using Microsoft Graph"""
        try:
            # Parse the site URL to get hostname and site path
            parsed_url = urllib.parse.urlparse(site_url)
            hostname = parsed_url.netloc
            site_path = parsed_url.path.strip('/')
            
            # Use Graph API to get site information
            graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json"
            }
            
            response = requests.get(graph_url, headers=headers)
            if response.status_code == 200:
                site_info = response.json()
                return site_info.get("id")
            else:
                logger.error(f"Failed to get site ID: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"Error getting site ID: {e}")
            return None
    
    async def test_connection(self, team_url: str, folder_path: str = "") -> bool:
        """Test connection to SharePoint site using Microsoft Graph"""
        try:
            if not self.authenticated:
                success = await self.authenticate(team_url)
                if not success:
                    # Try device code as fallback
                    success = await self.authenticate_device_code(team_url)
                    if not success:
                        return False
            
            # Test connection by getting site information
            site_id = self._get_site_id_from_url(team_url)
            if site_id:
                logger.info("Connection test successful")
                return True
            else:
                logger.error("Failed to connect to SharePoint site")
                return False
            
        except Exception as e:
            logger.error(f"Connection test failed: {e}")
            return False
    
    async def get_excel_files(self, team_url: str, folder_path: str = "") -> List[Dict]:
        """Get list of Excel files from SharePoint folder using Microsoft Graph"""
        try:
            if not self.authenticated:
                success = await self.authenticate(team_url)
                if not success:
                    raise Exception("Authentication failed")
            
            # Get site ID
            site_id = self._get_site_id_from_url(team_url)
            if not site_id:
                raise Exception("Could not get site information")
            
            # Construct Graph API URL for drive items
            if folder_path and folder_path.strip():
                # If specific folder path provided, try to find that folder
                folder_path = folder_path.strip('/')
                graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
            else:
                # Default to root of default document library
                graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json"
            }
            
            response = requests.get(graph_url, headers=headers)
            
            if response.status_code != 200:
                logger.error(f"Failed to get files: {response.status_code} - {response.text}")
                raise Exception(f"Failed to retrieve files: {response.status_code}")
            
            files_data = response.json()
            excel_files = []
            
            # Filter for Excel files
            for item in files_data.get("value", []):
                if item.get("file") and item.get("name", "").lower().endswith(('.xlsx', '.xlsm', '.xls')):
                    excel_files.append({
                        "name": item["name"],
                        "url": item["webUrl"],
                        "download_url": item.get("@microsoft.graph.downloadUrl", ""),
                        "modified": item.get("lastModifiedDateTime", "Unknown"),
                        "size": item.get("size", 0),
                        "id": item["id"]
                    })
            
            logger.info(f"Found {len(excel_files)} Excel files")
            return excel_files
            
        except Exception as e:
            logger.error(f"Error getting files: {e}")
            raise
    
    async def download_file(self, file_info: Dict, local_path: str) -> bool:
        """Download a file from SharePoint using Microsoft Graph"""
        try:
            if not self.authenticated:
                raise Exception("Not authenticated")
            
            download_url = file_info.get("download_url")
            if not download_url:
                raise Exception("No download URL available for file")
            
            headers = {
                "Authorization": f"Bearer {self.access_token}"
            }
            
            response = requests.get(download_url, headers=headers)
            
            if response.status_code == 200:
                with open(local_path, 'wb') as local_file:
                    local_file.write(response.content)
                logger.info(f"File downloaded successfully to {local_path}")
                return True
            else:
                logger.error(f"Failed to download file: {response.status_code}")
                return False
            
        except Exception as e:
            logger.error(f"Error downloading file: {e}")
            return False
    
    async def upload_file(self, local_path: str, team_url: str, folder_path: str, filename: str) -> bool:
        """Upload a file to SharePoint using Microsoft Graph"""
        try:
            if not self.authenticated:
                raise Exception("Not authenticated")
            
            # Get site ID
            site_id = self._get_site_id_from_url(team_url)
            if not site_id:
                raise Exception("Could not get site information")
            
            # Read local file
            with open(local_path, 'rb') as local_file:
                file_content = local_file.read()
            
            # Construct upload URL
            if folder_path and folder_path.strip():
                folder_path = folder_path.strip('/')
                upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}/{filename}:/content"
            else:
                upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{filename}:/content"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/octet-stream"
            }
            
            response = requests.put(upload_url, headers=headers, data=file_content)
            
            if response.status_code in [200, 201]:
                logger.info(f"File uploaded successfully: {filename}")
                return True
            else:
                logger.error(f"Failed to upload file: {response.status_code} - {response.text}")
                return False
            
        except Exception as e:
            logger.error(f"Error uploading file: {e}")
            return False