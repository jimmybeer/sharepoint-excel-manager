#!/usr/bin/env python3
"""
Example of using the SharePoint Excel Manager settings system
"""
import sys
from pathlib import Path

# Add src to path for import
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sharepoint_excel_manager.settings import AppSettings, SettingsManager


def main():
    """Demonstrate settings management functionality"""
    print("SharePoint Excel Manager Settings Example")
    print("=" * 50)
    
    # Create settings manager
    with SettingsManager("ExampleApp") as settings:
        print(f"Settings file location: {settings._config_file}")
        print()
        
        # Show current settings
        print("Current settings:")
        current = settings.settings
        for key, value in current.to_dict().items():
            print(f"  {key}: {value}")
        print()
        
        # Update some settings
        print("Updating settings...")
        settings.update(
            team_url="https://example.sharepoint.com/sites/myteam",
            document_folder="/Shared Documents/Projects",
            window_width=1024,
            window_height=768,
            theme="dark",
            auto_connect=True
        )
        
        print("Settings updated!")
        print()
        
        # Show updated settings
        print("Updated settings:")
        updated = settings.settings
        for key, value in updated.to_dict().items():
            print(f"  {key}: {value}")
        print()
        
        # Demonstrate individual get/set
        print("Individual setting operations:")
        old_url = settings.get("team_url")
        print(f"Old URL: {old_url}")
        
        settings.set("team_url", "https://newsite.sharepoint.com/sites/newteam")
        new_url = settings.get("team_url")
        print(f"New URL: {new_url}")
        print()
        
        # Show recent connections
        print("Recent connections:")
        recent = settings.get_recent_connections()
        for conn in recent:
            print(f"  URL: {conn['url']}")
            print(f"  Folder: {conn['folder']}")
            print(f"  Last used: {conn['last_used']}")
        print()
        
        # Export settings
        export_file = Path("example_settings_export.json")
        if settings.export_settings(export_file):
            print(f"Settings exported to: {export_file}")
            print(f"Export file contents:")
            print(export_file.read_text(encoding='utf-8'))
            
            # Clean up
            export_file.unlink()
            print("Export file cleaned up")
        
        print()
        print("Settings will be automatically saved when exiting context manager...")
    
    print("Example completed!")


if __name__ == "__main__":
    main()