"""
Settings management for SharePoint Excel Manager
Handles persistent storage of user preferences and configuration
"""
import json
import logging
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)


@dataclass
class AppSettings:
    """Data class to hold application settings"""
    team_url: str = ""
    document_folder: str = ""
    window_width: int = 800
    window_height: int = 600
    window_x: Optional[int] = None
    window_y: Optional[int] = None
    remember_credentials: bool = False
    auto_connect: bool = False
    last_username: str = ""
    theme: str = "system"  # system, light, dark
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert settings to dictionary"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'AppSettings':
        """Create settings from dictionary"""
        # Filter out any keys that don't match our dataclass fields
        valid_keys = {field.name for field in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in valid_keys}
        return cls(**filtered_data)


class SettingsManager:
    """Manages application settings with persistent storage"""
    
    def __init__(self, app_name: str = "SharePointExcelManager"):
        self.app_name = app_name
        self._settings = AppSettings()
        self._config_dir = self._get_config_directory()
        self._config_file = self._config_dir / "settings.json"
        
        # Ensure config directory exists
        self._config_dir.mkdir(parents=True, exist_ok=True)
        
        # Load settings on initialization
        self.load()
    
    def _get_config_directory(self) -> Path:
        """Get the appropriate config directory for the platform"""
        if os.name == 'nt':  # Windows
            config_base = Path(os.environ.get('APPDATA', Path.home()))
        elif os.name == 'posix':  # macOS, Linux
            if 'darwin' in os.sys.platform.lower():  # macOS
                config_base = Path.home() / "Library" / "Application Support"
            else:  # Linux
                config_base = Path.home() / ".config"
        else:
            config_base = Path.home()
        
        return config_base / self.app_name
    
    @property
    def settings(self) -> AppSettings:
        """Get current settings"""
        return self._settings
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a specific setting value"""
        return getattr(self._settings, key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set a specific setting value"""
        if hasattr(self._settings, key):
            setattr(self._settings, key, value)
        else:
            raise AttributeError(f"Unknown setting: {key}")
    
    def update(self, **kwargs) -> None:
        """Update multiple settings at once"""
        for key, value in kwargs.items():
            self.set(key, value)
    
    def load(self) -> bool:
        """Load settings from file"""
        try:
            if self._config_file.exists():
                with open(self._config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Validate and load settings
                self._settings = AppSettings.from_dict(data)
                logger.info(f"Settings loaded from {self._config_file}")
                return True
            else:
                logger.info("No settings file found, using defaults")
                return False
                
        except (json.JSONDecodeError, IOError, TypeError) as e:
            logger.warning(f"Error loading settings from {self._config_file}: {e}")
            logger.info("Using default settings")
            self._settings = AppSettings()  # Reset to defaults
            return False
    
    def save(self) -> bool:
        """Save current settings to file"""
        try:
            # Create backup of existing file
            if self._config_file.exists():
                backup_file = self._config_file.with_suffix('.json.backup')
                # Remove existing backup if it exists
                if backup_file.exists():
                    backup_file.unlink()
                self._config_file.rename(backup_file)
            
            # Write new settings
            with open(self._config_file, 'w', encoding='utf-8') as f:
                json.dump(self._settings.to_dict(), f, indent=2, ensure_ascii=False)
            
            logger.info(f"Settings saved to {self._config_file}")
            return True
            
        except (IOError, TypeError) as e:
            logger.error(f"Error saving settings to {self._config_file}: {e}")
            
            # Try to restore backup if save failed
            backup_file = self._config_file.with_suffix('.json.backup')
            if backup_file.exists():
                backup_file.rename(self._config_file)
                logger.info("Restored settings from backup")
            
            return False
    
    def reset_to_defaults(self) -> None:
        """Reset all settings to default values"""
        self._settings = AppSettings()
        logger.info("Settings reset to defaults")
    
    def get_recent_connections(self) -> list:
        """Get list of recent SharePoint connections"""
        # This could be extended to maintain a list of recent connections
        recent = []
        if self._settings.team_url:
            recent.append({
                'url': self._settings.team_url,
                'folder': self._settings.document_folder,
                'last_used': 'current'
            })
        return recent
    
    def add_recent_connection(self, url: str, folder: str = "") -> None:
        """Add a connection to recent list and update current settings"""
        self.update(team_url=url, document_folder=folder)
    
    def export_settings(self, file_path: Path) -> bool:
        """Export settings to a specified file"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self._settings.to_dict(), f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            logger.error(f"Error exporting settings: {e}")
            return False
    
    def import_settings(self, file_path: Path) -> bool:
        """Import settings from a specified file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self._settings = AppSettings.from_dict(data)
            return True
        except Exception as e:
            logger.error(f"Error importing settings: {e}")
            return False
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - automatically save settings"""
        self.save()