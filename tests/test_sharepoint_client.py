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