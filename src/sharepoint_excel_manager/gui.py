"""
GUI implementation using Toga for SharePoint Excel Manager
"""
import toga
from toga.style.pack import COLUMN, ROW, Pack

from .settings import SettingsManager
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
        
        # Initialize settings manager
        self.settings_manager = SettingsManager()
        
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
            value=self.settings_manager.get("team_url", ""),
            style=Pack(width=400, padding=(0, 0, 10, 0)),
            on_change=self.on_url_change
        )
        
        # Document folder input
        folder_label = toga.Label("Document Folder Path:", style=Pack(padding=(0, 0, 5, 0)))
        self.folder_input = toga.TextInput(
            value=self.settings_manager.get("document_folder", ""),
            style=Pack(width=400, padding=(0, 0, 10, 0)),
            on_change=self.on_folder_change
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
        
        # Settings button
        settings_button = toga.Button(
            "Settings",
            on_press=self.show_settings,
            style=Pack(padding=(0, 10, 0, 0), width=120)
        )
        
        # Device auth button (alternative for strict environments)
        device_auth_button = toga.Button(
            "Device Auth",
            on_press=self.device_auth_connection,
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
        button_box.add(settings_button)
        button_box.add(device_auth_button)
        
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
        
        # Load window size and position from settings
        self._restore_window_state()
    
    def _restore_window_state(self):
        """Restore window size and position from settings"""
        settings = self.settings_manager.settings
        
        # Set window size
        if settings.window_width and settings.window_height:
            try:
                self.main_window.size = (settings.window_width, settings.window_height)
            except Exception:
                pass  # Ignore if setting size fails
        
        # Set window position (if available and valid)
        if settings.window_x is not None and settings.window_y is not None:
            try:
                self.main_window.position = (settings.window_x, settings.window_y)
            except Exception:
                pass  # Ignore if setting position fails
    
    def _save_window_state(self):
        """Save current window state to settings"""
        try:
            size = self.main_window.size
            position = self.main_window.position
            
            self.settings_manager.update(
                window_width=size[0],
                window_height=size[1],
                window_x=position[0],
                window_y=position[1]
            )
        except Exception:
            pass  # Ignore if getting window state fails
    
    def on_exit(self):
        """Called when the application is closing"""
        # Save window state
        self._save_window_state()
        
        # Save current settings
        self.settings_manager.save()
        
        return True
    
    def on_url_change(self, widget):
        """Handle URL input changes"""
        self.settings_manager.set("team_url", widget.value.strip())
    
    def on_folder_change(self, widget):
        """Handle folder input changes"""
        self.settings_manager.set("document_folder", widget.value.strip())
    
    async def show_settings(self, widget):
        """Show settings dialog"""
        settings = self.settings_manager.settings
        
        # Create a simple info dialog for now
        # In a full implementation, this could be a proper settings window
        info_text = f"""Current Settings:
        
Team URL: {settings.team_url or 'Not set'}
Document Folder: {settings.document_folder or 'Not set'}
Window Size: {settings.window_width}x{settings.window_height}
Auto Connect: {'Yes' if settings.auto_connect else 'No'}
Remember Credentials: {'Yes' if settings.remember_credentials else 'No'}
Theme: {settings.theme}

Settings are automatically saved when changed.
Configuration file location: {self.settings_manager._config_file}"""
        
        await self.main_window.info_dialog("Settings", info_text)
    
    async def device_auth_connection(self, widget):
        """Test connection using device code authentication (for strict environments)"""
        team_url = self.url_input.value.strip()
        folder_path = self.folder_input.value.strip()
        
        if not team_url:
            self.status_label.text = "Please enter a team URL"
            self.status_label.style.color = "red"
            return
        
        # Show instructions to user
        await self.main_window.info_dialog(
            "Device Code Authentication",
            "This method will display a code that you need to enter on a separate device or browser.\n\n" +
            "This is useful in environments with strict security policies.\n\n" +
            "Click OK to continue, then check the console for instructions."
        )
        
        self.status_label.text = "Starting device authentication - check console for code..."
        self.status_label.style.color = "orange"
        
        try:
            success = await self.sharepoint_client.authenticate_device_code(team_url)
            if success:
                # Test the connection after authentication
                connection_success = await self.sharepoint_client.test_connection(team_url, folder_path)
                if connection_success:
                    self.status_label.text = "Device authentication and connection successful!"
                    self.status_label.style.color = "green"
                    
                    # Auto-save successful connection settings
                    self.settings_manager.update(
                        team_url=team_url,
                        document_folder=folder_path
                    )
                else:
                    self.status_label.text = "Authentication succeeded but connection test failed"
                    self.status_label.style.color = "orange"
            else:
                self.status_label.text = "Device authentication failed"
                self.status_label.style.color = "red"
        except Exception as e:
            self.status_label.text = f"Device auth error: {str(e)[:50]}..."
            self.status_label.style.color = "red"
    
    async def save_config(self, widget):
        """Save current configuration"""
        team_url = self.url_input.value.strip()
        document_folder = self.folder_input.value.strip()
        
        try:
            # Update settings
            self.settings_manager.update(
                team_url=team_url,
                document_folder=document_folder
            )
            
            # Save to file
            if self.settings_manager.save():
                self.status_label.text = "Configuration saved successfully"
                self.status_label.style.color = "green"
            else:
                self.status_label.text = "Error saving configuration"
                self.status_label.style.color = "red"
                
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
        
        self.status_label.text = "Testing connection - authentication may open browser..."
        self.status_label.style.color = "orange"
        
        try:
            success = await self.sharepoint_client.test_connection(team_url, folder_path)
            if success:
                self.status_label.text = "Connection successful!"
                self.status_label.style.color = "green"
                
                # Auto-save successful connection settings
                self.settings_manager.update(
                    team_url=team_url,
                    document_folder=folder_path
                )
            else:
                self.status_label.text = "Connection failed - check URL and try again"
                self.status_label.style.color = "red"
        except Exception as e:
            error_msg = str(e)
            if "AADSTS53003" in error_msg:
                self.status_label.text = "Connection blocked by Conditional Access - contact IT admin"
            elif "AADSTS50058" in error_msg:
                self.status_label.text = "Silent sign-in failed - please try again"
            else:
                self.status_label.text = f"Connection error: {error_msg[:50]}..."
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