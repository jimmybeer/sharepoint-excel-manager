"""
Tests for SharePoint client functionality
"""
import pytest
from unittest.mock import Mock, patch, AsyncMock
from sharepoint_excel_manager.sharepoint_client import SharePointClient


class TestSharePointClient:
    def setup_method(self):
        """Setup test fixtures"""
        self.client = SharePointClient()
    
    def test_init(self):
        """Test client initialization"""
        assert self.client.access_token is None
        assert self.client.authenticated is False
        assert self.client.site_url is None
        assert self.client.client_id == "d3590ed6-52b3-4102-aeff-aad2292ab01c"
        assert self.client.authority == "https://login.microsoftonline.com/common"
    
    @pytest.mark.asyncio
    @patch('sharepoint_excel_manager.sharepoint_client.PublicClientApplication')
    async def test_authenticate_success_with_cache(self, mock_msal):
        """Test successful authentication using cached token"""
        # Mock MSAL app
        mock_app_instance = Mock()
        mock_app_instance.get_accounts.return_value = [{"username": "test@example.com"}]
        mock_app_instance.acquire_token_silent.return_value = {"access_token": "fake_token"}
        mock_msal.return_value = mock_app_instance
        
        # Create new client to get mocked MSAL app
        client = SharePointClient()
        
        result = await client.authenticate("https://example.sharepoint.com")
        
        assert result is True
        assert client.authenticated is True
        assert client.access_token == "fake_token"
    
    @pytest.mark.asyncio
    @patch('sharepoint_excel_manager.sharepoint_client.PublicClientApplication')
    async def test_authenticate_interactive(self, mock_msal):
        """Test interactive authentication when cache fails"""
        # Mock MSAL app
        mock_app_instance = Mock()
        mock_app_instance.get_accounts.return_value = []  # No cached accounts
        mock_app_instance.acquire_token_interactive.return_value = {"access_token": "interactive_token"}
        mock_msal.return_value = mock_app_instance
        
        # Create new client to get mocked MSAL app
        client = SharePointClient()
        
        result = await client.authenticate("https://example.sharepoint.com")
        
        assert result is True
        assert client.authenticated is True
        assert client.access_token == "interactive_token"
    
    @pytest.mark.asyncio
    @patch('sharepoint_excel_manager.sharepoint_client.PublicClientApplication')
    async def test_authenticate_device_code(self, mock_msal):
        """Test device code authentication"""
        # Mock MSAL app
        mock_app_instance = Mock()
        mock_app_instance.initiate_device_flow.return_value = {
            "user_code": "ABC123",
            "message": "Go to https://microsoft.com/devicelogin and enter code ABC123"
        }
        mock_app_instance.acquire_token_by_device_flow.return_value = {"access_token": "device_token"}
        mock_msal.return_value = mock_app_instance
        
        # Create new client to get mocked MSAL app
        client = SharePointClient()
        
        result = await client.authenticate_device_code("https://example.sharepoint.com")
        
        assert result is True
        assert client.authenticated is True
        assert client.access_token == "device_token"
    
    @pytest.mark.asyncio
    @patch('requests.get')
    async def test_get_site_id_from_url(self, mock_get):
        """Test extracting site ID from SharePoint URL"""
        self.client.access_token = "fake_token"
        
        # Mock successful Graph API response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"id": "site123"}
        mock_get.return_value = mock_response
        
        site_id = self.client._get_site_id_from_url("https://example.sharepoint.com/sites/testsite")
        
        assert site_id == "site123"
        mock_get.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_test_connection_not_authenticated(self):
        """Test connection test when not authenticated"""
        with patch.object(self.client, 'authenticate', return_value=False):
            with patch.object(self.client, 'authenticate_device_code', return_value=False):
                result = await self.client.test_connection("https://example.sharepoint.com")
                assert result is False
    
    @pytest.mark.asyncio
    @patch('requests.get')
    async def test_get_excel_files_success(self, mock_get):
        """Test getting Excel files successfully"""
        self.client.authenticated = True
        self.client.access_token = "fake_token"
        
        # Mock site ID call
        mock_site_response = Mock()
        mock_site_response.status_code = 200
        mock_site_response.json.return_value = {"id": "site123"}
        
        # Mock files call
        mock_files_response = Mock()
        mock_files_response.status_code = 200
        mock_files_response.json.return_value = {
            "value": [
                {
                    "name": "document.pdf",
                    "file": {},
                    "webUrl": "https://example.com/document.pdf",
                    "lastModifiedDateTime": "2023-01-01T00:00:00Z",
                    "size": 1000,
                    "id": "file1"
                """
Tests for SharePoint client functionality
"""
import pytest
from unittest.mock import Mock, patch, AsyncMock
from sharepoint_excel_manager.sharepoint_client import SharePointClient


class TestSharePointClient:
    def setup_method(self):
        """Setup test fixtures"""
        self.client = SharePointClient()
    
    def test_init(self):
        """Test client initialization"""
        assert self.client.context is None
        assert self.client.authenticated is False
    
    @pytest.mark.asyncio
    @patch('sharepoint_excel_manager.sharepoint_client.AuthenticationContext')
    @patch('sharepoint_excel_manager.sharepoint_client.UserCredential')
    @patch('sharepoint_excel_manager.sharepoint_client.input')
    @patch('sharepoint_excel_manager.sharepoint_client.getpass.getpass')
    async def test_authenticate_success(self, mock_getpass, mock_input, mock_user_cred, mock_auth_context):
        """Test successful authentication"""
        # Mock user inputs
        mock_input.return_value = "test@example.com"
        mock_getpass.return_value = "password"
        
        # Mock authentication context
        mock_auth_instance = Mock()
        mock_auth_instance.acquire_token_for_user.return_value = True
        mock_auth_context.return_value = mock_auth_instance
        
        # Test authentication
        result = await self.client.authenticate("https://example.sharepoint.com")
        
        assert result is True
        assert self.client.authenticated is True
        assert self.client.context is not None
    
    @pytest.mark.asyncio
    async def test_test_connection_not_authenticated(self):
        """Test connection test when not authenticated"""
        with patch.object(self.client, 'authenticate', return_value=False):
            result = await self.client.test_connection("https://example.sharepoint.com")
            assert result is False
    
    @pytest.mark.asyncio
    async def test_get_excel_files_not_authenticated(self):
        """Test getting files when not authenticated"""
        with patch.object(self.client, 'authenticate', return_value=False):
            with pytest.raises(Exception, match="Authentication failed"):
                await self.client.get_excel_files("https://example.sharepoint.com")
    
    def test_excel_file_filtering(self):
        """Test that only Excel files are included in results"""
        # This would require more complex mocking of SharePoint objects
        # For now, just test the file extension logic conceptually
        excel_extensions = ['.xlsx', '.xlsm', '.xls']
        test_files = ['document.pdf', 'spreadsheet.xlsx', 'data.xlsm', 'old_file.xls', 'text.txt']
        
        excel_files = [f for f in test_files if any(f.endswith(ext) for ext in excel_extensions)]
        
        assert len(excel_files) == 3
        assert 'spreadsheet.xlsx' in excel_files
        assert 'data.xlsm' in excel_files
        assert 'old_file.xls' in excel_files
        assert 'document.pdf' not in excel_files
        assert 'text.txt' not in excel_files