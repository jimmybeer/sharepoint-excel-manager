"""
Tests for settings management functionality
"""
import json
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from sharepoint_excel_manager.settings import AppSettings, SettingsManager


class TestAppSettings:
    def test_default_values(self):
        """Test default settings values"""
        settings = AppSettings()
        assert settings.team_url == ""
        assert settings.document_folder == ""
        assert settings.window_width == 800
        assert settings.window_height == 600
        assert settings.window_x is None
        assert settings.window_y is None
        assert settings.remember_credentials is False
        assert settings.auto_connect is False
        assert settings.last_username == ""
        assert settings.theme == "system"
    
    def test_to_dict(self):
        """Test converting settings to dictionary"""
        settings = AppSettings(
            team_url="https://example.sharepoint.com",
            document_folder="/Documents",
            window_width=1024,
            window_height=768
        )
        
        result = settings.to_dict()
        
        assert result["team_url"] == "https://example.sharepoint.com"
        assert result["document_folder"] == "/Documents"
        assert result["window_width"] == 1024
        assert result["window_height"] == 768
    
    def test_from_dict(self):
        """Test creating settings from dictionary"""
        data = {
            "team_url": "https://test.sharepoint.com",
            "document_folder": "/TestDocs",
            "window_width": 1200,
            "window_height": 900,
            "unknown_field": "should_be_ignored"
        }
        
        settings = AppSettings.from_dict(data)
        
        assert settings.team_url == "https://test.sharepoint.com"
        assert settings.document_folder == "/TestDocs"
        assert settings.window_width == 1200
        assert settings.window_height == 900
        # Unknown fields should be ignored
        assert not hasattr(settings, "unknown_field")


class TestSettingsManager:
    def test_init_creates_config_dir(self):
        """Test that initialization creates config directory"""
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(SettingsManager, '_get_config_directory', return_value=Path(temp_dir) / "test_app"):
                manager = SettingsManager("TestApp")
                assert manager._config_dir.exists()
    
    def test_get_set_settings(self):
        """Test getting and setting individual settings"""
        manager = SettingsManager("TestApp")
        
        # Test setting and getting
        manager.set("team_url", "https://test.sharepoint.com")
        assert manager.get("team_url") == "https://test.sharepoint.com"
        
        # Test default value
        assert manager.get("nonexistent_setting", "default") == "default"
        
        # Test setting invalid key
        with pytest.raises(AttributeError):
            manager.set("invalid_key", "value")
    
    def test_update_multiple_settings(self):
        """Test updating multiple settings at once"""
        manager = SettingsManager("TestApp")
        
        manager.update(
            team_url="https://example.com",
            document_folder="/docs",
            window_width=1024
        )
        
        assert manager.get("team_url") == "https://example.com"
        assert manager.get("document_folder") == "/docs"
        assert manager.get("window_width") == 1024
    
    def test_save_and_load(self):
        """Test saving and loading settings"""
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir) / "test_app"
            
            with patch.object(SettingsManager, '_get_config_directory', return_value=config_dir):
                # Create manager and set some values
                manager1 = SettingsManager("TestApp")
                manager1.update(
                    team_url="https://test.com",
                    document_folder="/test",
                    window_width=1200
                )
                
                # Save settings
                assert manager1.save() is True
                
                # Create new manager and verify settings loaded
                manager2 = SettingsManager("TestApp")
                assert manager2.get("team_url") == "https://test.com"
                assert manager2.get("document_folder") == "/test"
                assert manager2.get("window_width") == 1200
    
    def test_load_nonexistent_file(self):
        """Test loading when no settings file exists"""
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir) / "nonexistent_app"
            
            with patch.object(SettingsManager, '_get_config_directory', return_value=config_dir):
                manager = SettingsManager("NonexistentApp")
                
                # Should use default values
                assert manager.get("team_url") == ""
                assert manager.get("window_width") == 800
    
    def test_load_invalid_json(self):
        """Test loading corrupted settings file"""
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir) / "test_app"
            config_dir.mkdir()
            
            # Create invalid JSON file
            config_file = config_dir / "settings.json"
            config_file.write_text("invalid json content")
            
            with patch.object(SettingsManager, '_get_config_directory', return_value=config_dir):
                manager = SettingsManager("TestApp")
                
                # Should fall back to defaults
                assert manager.get("team_url") == ""
                assert manager.get("window_width") == 800
    
    def test_reset_to_defaults(self):
        """Test resetting settings to defaults"""
        manager = SettingsManager("TestApp")
        
        # Set some custom values
        manager.update(
            team_url="https://custom.com",
            window_width=1600
        )
        
        # Reset and verify defaults
        manager.reset_to_defaults()
        assert manager.get("team_url") == ""
        assert manager.get("window_width") == 800
    
    def test_recent_connections(self):
        """Test recent connections functionality"""
        manager = SettingsManager("TestApp")
        
        # Initially no recent connections
        recent = manager.get_recent_connections()
        assert len(recent) == 0
        
        # Add a connection
        manager.add_recent_connection("https://example.com", "/docs")
        
        recent = manager.get_recent_connections()
        assert len(recent) == 1
        assert recent[0]["url"] == "https://example.com"
        assert recent[0]["folder"] == "/docs"
        assert recent[0]["last_used"] == "current"
    
    def test_export_import_settings(self):
        """Test exporting and importing settings"""
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir) / "test_app"
            export_file = Path(temp_dir) / "exported_settings.json"
            
            with patch.object(SettingsManager, '_get_config_directory', return_value=config_dir):
                # Create manager with custom settings
                manager1 = SettingsManager("TestApp")
                manager1.update(
                    team_url="https://export.test.com",
                    document_folder="/export_test",
                    window_width=1400,
                    theme="dark"
                )
                
                # Export settings
                assert manager1.export_settings(export_file) is True
                assert export_file.exists()
                
                # Create new manager and import settings
                manager2 = SettingsManager("TestApp2")
                assert manager2.import_settings(export_file) is True
                
                # Verify imported settings
                assert manager2.get("team_url") == "https://export.test.com"
                assert manager2.get("document_folder") == "/export_test"
                assert manager2.get("window_width") == 1400
                assert manager2.get("theme") == "dark"
    
    def test_context_manager(self):
        """Test using SettingsManager as context manager"""
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir) / "test_app"
            
            with patch.object(SettingsManager, '_get_config_directory', return_value=config_dir):
                # Use as context manager
                with SettingsManager("TestApp") as manager:
                    manager.set("team_url", "https://context.test.com")
                    # Settings should be automatically saved on exit
                
                # Verify settings were saved
                manager2 = SettingsManager("TestApp")
                assert manager2.get("team_url") == "https://context.test.com"
    
    @patch('os.name', 'nt')
    @patch.dict('os.environ', {'APPDATA': '/fake/appdata'})
    def test_get_config_directory_windows(self):
        """Test config directory detection on Windows"""
        manager = SettingsManager("TestApp")
        expected = Path("/fake/appdata/SharePointExcelManager")
        assert manager._config_dir == expected
    
    @patch('os.name', 'posix')
    @patch('os.sys.platform', 'darwin')
    def test_get_config_directory_macos(self):
        """Test config directory detection on macOS"""
        with patch('pathlib.Path.home') as mock_home:
            mock_home.return_value = Path('/Users/testuser')
            manager = SettingsManager("TestApp")
            expected = Path("/Users/testuser/Library/Application Support/SharePointExcelManager")
            assert manager._config_dir == expected
    
    @patch('os.name', 'posix')
    @patch('os.sys.platform', 'linux')
    def test_get_config_directory_linux(self):
        """Test config directory detection on Linux"""
        with patch('pathlib.Path.home') as mock_home:
            mock_home.return_value = Path('/home/testuser')
            manager = SettingsManager("TestApp")
            expected = Path("/home/testuser/.config/SharePointExcelManager")
            assert manager._config_dir == expected