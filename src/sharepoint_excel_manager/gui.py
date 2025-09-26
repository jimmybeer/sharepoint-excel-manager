"""
GUI implementation using Toga for SharePoint Excel Manager
"""
import sys
import threading
import webbrowser

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
    
    def print_to_console(self, message):
        """Add message to console text area"""
        current_text = self.console_text.value
        self.console_text.value = current_text + message + "\n"
    
    def startup(self):
        """Initialize the application"""
        self.sharepoint_client = SharePointClient()
        
        # Initialize settings manager
        self.settings_manager = SettingsManager()
        
        # Main container
        main_box = toga.Box(style=Pack(direction=COLUMN, margin=20))
        
        # Title
        title = toga.Label(
            "SharePoint Excel Manager",
            style=Pack(margin=(0, 0, 20, 0), text_align="center", font_size=18, font_weight="bold")
        )
        
        # Team URL input
        url_label = toga.Label("Team SharePoint URL:", style=Pack(margin=(0, 0, 5, 0)))
        self.url_input = toga.TextInput(
            value=self.settings_manager.get("team_url", ""),
            style=Pack(width=400, margin=(0, 0, 10, 0)),
            on_change=self.on_url_change
        )
        
        # Document folder input
        folder_label = toga.Label("Document Folder Path:", style=Pack(margin=(0, 0, 5, 0)))
        self.folder_input = toga.TextInput(
            value=self.settings_manager.get("document_folder", ""),
            style=Pack(width=400, margin=(0, 0, 10, 0)),
            on_change=self.on_folder_change
        )
        
        # Buttons container
        button_box = toga.Box(style=Pack(direction=ROW, margin=(20, 0, 0, 0)))
        
        # Test connection button
        test_button = toga.Button(
            "Test Connection",
            on_press=self.test_connection,
            style=Pack(margin=(0, 10, 0, 0), width=120)
        )
        
        # Save config button
        save_button = toga.Button(
            "Save Config",
            on_press=self.save_config,
            style=Pack(margin=(0, 10, 0, 0), width=120)
        )
        
        # Browse files button
        browse_button = toga.Button(
            "Browse Files",
            on_press=self.browse_files,
            style=Pack(margin=(0, 10, 0, 0), width=120)
        )
        
        # Settings button
        settings_button = toga.Button(
            "Settings",
            on_press=self.show_settings,
            style=Pack(margin=(0, 10, 0, 0), width=120)
        )
        
        # Device auth button (alternative for strict environments)
        # Note: Device auth may hang - only use if Test Connection fails
        device_auth_button = toga.Button(
            "Device Auth (if needed)",
            on_press=self.device_auth_connection,
            style=Pack(margin=(0, 10, 0, 0), width=140)
        )
        
        # Clear console button
        clear_button = toga.Button(
            "Clear Console",
            on_press=self.clear_console,
            style=Pack(margin=(0, 10, 0, 0), width=120)
        )
        
        # Status label
        self.status_label = toga.Label(
            "Ready",
            style=Pack(margin=(20, 0, 0, 0), color="green")
        )
        
        # Files text area
        files_label = toga.Label("Excel Files:", style=Pack(margin=(20, 0, 5, 0)))
        self.files_text = toga.MultilineTextInput(
            readonly=True,
            style=Pack(height=150, margin=(0, 0, 10, 0))
        )
        
        # Console text area
        console_label = toga.Label("Console Output:", style=Pack(margin=(0, 0, 5, 0)))
        self.console_text = toga.MultilineTextInput(
            readonly=False,  # Allow editing so URLs can be clicked/selected
            style=Pack(height=150, margin=(0, 0, 0, 0))
        )
        
        # Add components to containers
        button_box.add(test_button)
        button_box.add(save_button)
        button_box.add(browse_button)
        button_box.add(settings_button)
        button_box.add(device_auth_button)
        button_box.add(clear_button)
        
        main_box.add(title)
        main_box.add(url_label)
        main_box.add(self.url_input)
        main_box.add(folder_label)
        main_box.add(self.folder_input)
        main_box.add(button_box)
        main_box.add(self.status_label)
        main_box.add(files_label)
        main_box.add(self.files_text)
        main_box.add(console_label)
        main_box.add(self.console_text)
        
        # Create main window
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = main_box
        self.main_window.show()
        
        # Load window size and position from settings
        self._restore_window_state()
        
        # Redirect print statements to console
        self._redirect_console()
        
        # Welcome message
        self.print_to_console("SharePoint Excel Manager started")
        self.print_to_console("Use 'Test Connection' for most authentication scenarios")
        self.print_to_console("'Device Auth' is available but may hang - only use if Test Connection fails")
    
    def _redirect_console(self):
        """Redirect print statements to the console text area"""
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        
        class ConsoleRedirect:
            def __init__(self, text_widget, original_stream):
                self.text_widget = text_widget
                self.original_stream = original_stream
            
            def write(self, text):
                if text.strip():
                    current_text = self.text_widget.value
                    self.text_widget.value = current_text + text
                self.original_stream.write(text)
            
            def flush(self):
                self.original_stream.flush()
        
        sys.stdout = ConsoleRedirect(self.console_text, original_stdout)
        sys.stderr = ConsoleRedirect(self.console_text, original_stderr)
    
    def copy_to_clipboard(self, text):
        """Copy text to clipboard - platform independent"""
        try:
            # Try using pyperclip if available
            import pyperclip
            pyperclip.copy(text)
            return True
        except ImportError:
            try:
                # Fallback for Windows
                import subprocess
                subprocess.run(['clip'], input=text.encode(), check=True)
                return True
            except:
                try:
                    # Fallback for macOS
                    subprocess.run(['pbcopy'], input=text.encode(), check=True)
                    return True
                except:
                    try:
                        # Fallback for Linux
                        subprocess.run(['xclip', '-selection', 'clipboard'], input=text.encode(), check=True)
                        return True
                    except:
                        return False
    
    def clear_console(self, widget):
        """Clear the console text area"""
        self.console_text.value = ""
        self.print_to_console("Console cleared")
    
    def _restore_window_state(self):
        """Restore window size and position from settings"""
        settings = self.settings_manager.settings
        
        # Set window size
        if settings.window_width and settings.window_height:
            try:
                self.main_window.size = (settings.window_width, max(settings.window_height, 700))
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
        
        self.status_label.text = "Preparing device authentication..."
        self.status_label.style.color = "orange"
        
        try:
            # Get device code and URL immediately
            flow = self.sharepoint_client.app.initiate_device_flow(scopes=self.sharepoint_client.scope)
            
            if "user_code" not in flow:
                raise Exception("Failed to create device flow")
            
            verification_uri = flow.get("verification_uri", "")
            user_code = flow.get("user_code", "")
            
            # Show dialog with code and instructions
            dialog_message = f"""Device Authentication Setup:

URL: {verification_uri}
Code: {user_code}

When you click OK:
1. A browser will open to the authentication URL
2. The code will be copied to your clipboard
3. Paste the code (Ctrl+V) in the browser
4. Complete the authentication

The application will wait for you to complete the process.
This may take a few moments after you authenticate."""
            
            await self.main_window.info_dialog("Device Code Authentication", dialog_message)
            
            # Copy code to clipboard
            if self.copy_to_clipboard(user_code):
                self.print_to_console(f"Device code copied to clipboard: {user_code}")
            else:
                self.print_to_console(f"Could not copy to clipboard. Code: {user_code}")
            
            # Open browser
            def open_browser():
                try:
                    webbrowser.open(verification_uri)
                    self.print_to_console(f"Browser opened to: {verification_uri}")
                except Exception as e:
                    self.print_to_console(f"Could not open browser: {e}")
            
            threading.Thread(target=open_browser, daemon=True).start()
            
            self.status_label.text = "Complete authentication in browser - app will freeze temporarily"
            self.status_label.style.color = "orange"
            
            # Use regular print for debugging (will appear in terminal)
            print("DEBUG: Starting device authentication flow...")
            
            # Complete device flow (this will block)
            try:
                print("DEBUG: About to call acquire_token_by_device_flow...")
                print("DEBUG: If this hangs for more than 2 minutes after you complete browser auth,")
                print("DEBUG: close this app and restart it, then try 'Test Connection' instead.")
                
                result = self.sharepoint_client.app.acquire_token_by_device_flow(flow)
                print("DEBUG: acquire_token_by_device_flow returned")
                print(f"DEBUG: Result: {result}")
                
            except Exception as flow_error:
                print(f"DEBUG: Exception caught: {flow_error}")
                self.status_label.text = f"Device flow error: {str(flow_error)[:50]}..."
                self.status_label.style.color = "red"
                return
            
            print("DEBUG: Checking result...")
            if result and "access_token" in result:
                self.sharepoint_client.access_token = result["access_token"]
                self.sharepoint_client.authenticated = True
                
                # Test the connection after authentication
                connection_success = await self.sharepoint_client.test_connection(team_url, folder_path)
                if connection_success:
                    self.status_label.text = "Device authentication and connection successful!"
                    self.status_label.style.color = "green"
                    self.print_to_console("Device authentication successful!")
                    
                    # Auto-save successful connection settings
                    self.settings_manager.update(
                        team_url=team_url,
                        document_folder=folder_path
                    )
                else:
                    self.status_label.text = "Authentication succeeded but connection test failed"
                    self.status_label.style.color = "orange"
                    self.print_to_console("Authentication succeeded but connection test failed")
            else:
                error_msg = result.get("error_description", "Authentication failed")
                self.status_label.text = "Device authentication failed"
                self.status_label.style.color = "red"
                self.print_to_console(f"Device authentication failed: {error_msg}")
                
        except Exception as e:
            self.status_label.text = f"Device auth error: {str(e)[:50]}..."
            self.status_label.style.color = "red"
            self.print_to_console(f"Device authentication error: {str(e)}")
    
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
                self.print_to_console("Configuration saved successfully")
            else:
                self.status_label.text = "Error saving configuration"
                self.status_label.style.color = "red"
                self.print_to_console("Error saving configuration")
                
        except Exception as e:
            self.status_label.text = f"Error saving config: {str(e)}"
            self.status_label.style.color = "red"
            self.print_to_console(f"Error saving config: {str(e)}")
    
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
        self.print_to_console(f"Testing connection to: {team_url}")
        
        try:
            success = await self.sharepoint_client.test_connection(team_url, folder_path)
            if success:
                self.status_label.text = "Connection successful!"
                self.status_label.style.color = "green"
                self.print_to_console("Connection successful!")
                
                # Auto-save successful connection settings
                self.settings_manager.update(
                    team_url=team_url,
                    document_folder=folder_path
                )
            else:
                self.status_label.text = "Connection failed - check URL and try again"
                self.status_label.style.color = "red"
                self.print_to_console("Connection failed - check URL and try again")
        except Exception as e:
            error_msg = str(e)
            if "AADSTS53003" in error_msg:
                self.status_label.text = "Connection blocked by Conditional Access - try Device Auth"
                self.print_to_console("Connection blocked by Conditional Access - try Device Auth")
            elif "AADSTS50058" in error_msg:
                self.status_label.text = "Silent sign-in failed - try Device Auth"
                self.print_to_console("Silent sign-in failed - try Device Auth")
            else:
                self.status_label.text = f"Connection error: {error_msg[:50]}..."
                self.print_to_console(f"Connection error: {error_msg}")
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
        self.print_to_console("Loading Excel files...")
        
        try:
            files = await self.sharepoint_client.get_excel_files(team_url, folder_path)
            
            # Update files display
            if files:
                file_text = f"Found {len(files)} Excel files:\n\n"
                for i, file_info in enumerate(files, 1):
                    file_text += f"{i}. {file_info['name']}\n"
                    file_text += f"   Modified: {file_info.get('modified', 'Unknown')}\n"
                    file_text += f"   Size: {file_info.get('size', 0)} bytes\n\n"
                self.files_text.value = file_text
            else:
                self.files_text.value = "No Excel files found"
            
            self.status_label.text = f"Found {len(files)} Excel files"
            self.status_label.style.color = "green"
            self.print_to_console(f"Found {len(files)} Excel files")
            
        except Exception as e:
            self.status_label.text = f"Error browsing files: {str(e)}"
            self.status_label.style.color = "red"
            self.print_to_console(f"Error browsing files: {str(e)}")