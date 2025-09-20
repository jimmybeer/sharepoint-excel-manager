"""
GUI implementation using Toga for SharePoint Excel Manager
"""
import json
import os
from pathlib import Path

import toga
from toga.style.pack import COLUMN, ROW, Pack

from .sharepoint_client import SharePointClient


class SharePointExcelApp(toga.App):
    
    def __init__(self):
        super().__init__(
            formal_name="SharePoint Excel Manager",
            app_id="com.example.sharepoint_excel_manager",
            app_name="SharePoint Excel Manager",
            description="A GUI application for managing Excel files in SharePoint",
            author="Your Name",
            version="1.0.0"
        )
    
    def startup(self):
        """Initialize the application"""
        self.sharepoint_client = SharePointClient()
        self.config_file = Path.home() / ".sharepoint_excel_config.json"
        
        # Load saved configuration
        self.config = self.load_config()
        
        # Main container
        main_box = toga.Box(style=Pack(direction=COLUMN, padding=20))
        
        # Title
        title = toga.Label(
            "SharePoint Excel Manager",
            style=Pack(padding=(0, 0, 20, 0), text_align="center", font_size=18, font_weight="bold")
        )
        
        # Team URL input
        url_label = toga.Label("Team SharePoint URL:", style=Pack(padding=(0, 0, 5, 0)))
        self.url_input = toga.TextInput(
            value=self.config.get("team_url", ""),
            style=Pack(width=400, padding=(0, 0, 10, 0))
        )
        
        # Document folder input
        folder_label = toga.Label("Document Folder Path:", style=Pack(padding=(0, 0, 5, 0)))
        self.folder_input = toga.TextInput(
            value=self.config.get("document_folder", ""),
            style=Pack(width=400, padding=(0, 0, 10, 0))
        )
        
        # Buttons container
        button_box = toga.Box(style=Pack(direction=ROW, padding=(20, 0, 0, 0)))
        
        # Test connection button
        test_button = toga.Button(
            "Test Connection",
            on_press=self.test_connection,
            style=Pack(padding=(0, 10, 0, 0), width=120)
        )
        
        # Save config button
        save_button = toga.Button(
            "Save Config",
            on_press=self.save_config,
            style=Pack(padding=(0, 10, 0, 0), width=120)
        )
        
        # Browse files button
        browse_button = toga.Button(
            "Browse Files",
            on_press=self.browse_files,
            style=Pack(padding=(0, 10, 0, 0), width=120)
        )
        
        # Status label
        self.status_label = toga.Label(
            "Ready",
            style=Pack(padding=(20, 0, 0, 0), color="green")
        )
        
        # File list
        self.file_list = toga.DetailedList(
            style=Pack(height=300, padding=(20, 0, 0, 0))
        )
        
        # Add components to containers
        button_box.add(test_button)
        button_box.add(save_button)
        button_box.add(browse_button)
        
        main_box.add(title)
        main_box.add(url_label)
        main_box.add(self.url_input)
        main_box.add(folder_label)
        main_box.add(self.folder_input)
        main_box.add(button_box)
        main_box.add(self.status_label)
        main_box.add(self.file_list)
        
        # Create main window
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = main_box
        self.main_window.show()
    
    def load_config(self):
        """Load configuration from file"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
        return {}
    
    async def save_config(self, widget):
        """Save current configuration"""
        config = {
            "team_url": self.url_input.value,
            "document_folder": self.folder_input.value
        }
        
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
            self.status_label.text = "Configuration saved successfully"
            self.status_label.style.color = "green"
        except Exception as e:
            self.status_label.text = f"Error saving config: {str(e)}"
            self.status_label.style.color = "red"
    
    async def test_connection(self, widget):
        """Test connection to SharePoint"""
        team_url = self.url_input.value.strip()
        folder_path = self.folder_input.value.strip()
        
        if not team_url:
            self.status_label.text = "Please enter a team URL"
            self.status_label.style.color = "red"
            return
        
        self.status_label.text = "Testing connection..."
        self.status_label.style.color = "orange"
        
        try:
            success = await self.sharepoint_client.test_connection(team_url, folder_path)
            if success:
                self.status_label.text = "Connection successful!"
                self.status_label.style.color = "green"
            else:
                self.status_label.text = "Connection failed - check credentials"
                self.status_label.style.color = "red"
        except Exception as e:
            self.status_label.text = f"Connection error: {str(e)}"
            self.status_label.style.color = "red"
    
    async def browse_files(self, widget):
        """Browse Excel files in SharePoint"""
        team_url = self.url_input.value.strip()
        folder_path = self.folder_input.value.strip()
        
        if not team_url:
            self.status_label.text = "Please enter a team URL and test connection first"
            self.status_label.style.color = "red"
            return
        
        self.status_label.text = "Loading files..."
        self.status_label.style.color = "orange"
        
        try:
            files = await self.sharepoint_client.get_excel_files(team_url, folder_path)
            
            # Clear existing items
            self.file_list.data.clear()
            
            # Add files to list
            for file_info in files:
                self.file_list.data.append({
                    "title": file_info["name"],
                    "subtitle": f"Modified: {file_info.get('modified', 'Unknown')}",
                    "icon": None
                })
            
            self.status_label.text = f"Found {len(files)} Excel files"
            self.status_label.style.color = "green"
            
        except Exception as e:
            self.status_label.text = f"Error browsing files: {str(e)}"
            self.status_label.style.color = "red"