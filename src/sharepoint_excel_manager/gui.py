"""
GUI implementation using Toga for SharePoint Excel Manager
"""
import asyncio
import sys
import threading
import webbrowser

import toga
from toga.style.pack import COLUMN, ROW, Pack

from .excel_manager import ExcelManager
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
        """Add message to console text area and auto-scroll to bottom"""
        current_text = self.console_text.value
        self.console_text.value = current_text + message + "\n"
        
        # Auto-scroll to bottom
        try:
            # Try to scroll to the end of the text
            # This may not work on all platforms/versions of Toga
            if hasattr(self.console_text, 'scroll_to_bottom'):
                self.console_text.scroll_to_bottom()
            elif hasattr(self.console_text, 'set_cursor_position'):
                # Alternative: move cursor to end which may trigger scroll
                text_length = len(self.console_text.value)
                self.console_text.set_cursor_position(text_length)
        except Exception:
            # If scrolling fails, just continue without it
            pass
    
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
        files_label = toga.Label("Files:", style=Pack(margin=(20, 0, 5, 0)))
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
        
        await self.main_window.dialog(toga.InfoDialog("Settings", info_text))
    
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
            
            await self.main_window.dialog(toga.InfoDialog("Device Code Authentication", dialog_message))
            
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
        """Browse files in SharePoint"""
        team_url = self.url_input.value.strip()
        folder_path = self.folder_input.value.strip()
        
        if not team_url:
            self.status_label.text = "Please enter a team URL and test connection first"
            self.status_label.style.color = "red"
            return
        
        self.status_label.text = "Loading files..."
        self.status_label.style.color = "orange"
        self.print_to_console("Loading files...")
        
        try:
            # Clear the files text area first
            self.files_text.value = ""
            
            # Get all files and folders
            files = await self.sharepoint_client.get_all_files(team_url, folder_path)
            
            if not files:
                self.files_text.value = "No items found"
                self.status_label.text = "No items found"
                self.status_label.style.color = "orange"
                self.print_to_console("No items found")
                return
            
            # Format files as a table
            header = f"{'Name':<40} {'Type':<10} {'Size':<12} {'Modified':<20}\n"
            header += "-" * 82 + "\n"
            file_text = header
            
            excel_files = []
            folder_count = 0
            file_count = 0
            
            for file_info in files:
                name = file_info['name'][:37] + "..." if len(file_info['name']) > 40 else file_info['name']
                
                if file_info['type'] == 'folder':
                    file_type = "Folder"
                    size = f"{file_info.get('size', 0)} items"
                    folder_count += 1
                else:
                    file_count += 1
                    if file_info['name'].lower().endswith(('.xlsx', '.xlsm', '.xls')):
                        file_type = "Excel"
                        excel_files.append(file_info)
                    else:
                        file_type = "File"
                    size = self.format_file_size(file_info.get('size', 0))
                
                modified = self.format_date(file_info.get('modified', 'Unknown'))
                
                file_text += f"{name:<40} {file_type:<10} {size:<12} {modified:<20}\n"
            
            self.files_text.value = file_text
            self.status_label.text = f"Found {folder_count} folders, {file_count} files ({len(excel_files)} Excel)"
            self.status_label.style.color = "green"
            self.print_to_console(f"Found {folder_count} folders, {file_count} files ({len(excel_files)} Excel)")
            
            # Show Excel file selection dialog if any Excel files found
            if excel_files:
                await self.show_excel_selection_dialog(excel_files)
            
        except Exception as e:
            self.status_label.text = f"Error browsing files: {str(e)}"
            self.status_label.style.color = "red"
            self.print_to_console(f"Error browsing files: {str(e)}")
    
    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        elif size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes // 1024} KB"
        else:
            return f"{size_bytes // (1024 * 1024)} MB"
    
    def format_date(self, date_str):
        """Format date string to readable format"""
        if date_str == 'Unknown':
            return date_str
        try:
            from datetime import datetime
            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return dt.strftime('%Y-%m-%d %H:%M')
        except:
            return date_str[:16] if len(date_str) > 16 else date_str
    
    async def show_excel_selection_dialog(self, excel_files):
        """Show selection dialog for Excel files with proper list and buttons"""
        try:
            # Create a new window for file selection
            selection_window = toga.Window(title="Select Excel File to Update")
            selection_window.size = (300, 400)
            
            # Center the dialog relative to main window
            try:
                main_pos = self.main_window.position
                main_size = self.main_window.size
                
                # Calculate center position
                center_x = main_pos[0] + (main_size[0] - 300) // 2
                center_y = main_pos[1] + (main_size[1] - 400) // 2
                
                selection_window.position = (center_x, center_y)
            except Exception:
                # If centering fails, just use default position
                pass
            
            # Main container
            main_box = toga.Box(style=Pack(direction=COLUMN, margin=10))
            
            # Title
            title_label = toga.Label(
                f"Select Excel file to update:",
                style=Pack(margin=(0, 0, 10, 0), font_weight="bold", text_align="center")
            )
            
            count_label = toga.Label(
                f"({len(excel_files)} files found)",
                style=Pack(margin=(0, 0, 15, 0), text_align="center")
            )
            
            # Create file list using DetailedList
            self.file_list_selection = toga.DetailedList(
                style=Pack(height=250, margin=(0, 0, 15, 0))
            )
            
            # Populate the list with Excel files
            for i, file_info in enumerate(excel_files):
                modified_date = self.format_date(file_info.get('modified', 'Unknown'))
                size = self.format_file_size(file_info.get('size', 0))
                
                self.file_list_selection.data.append({
                    "title": file_info['name'],
                    "subtitle": f"{size} - {modified_date}",
                    "icon": None
                })
            
            # Buttons container
            button_box = toga.Box(style=Pack(direction=ROW, margin=(5, 0, 0, 0)))
            
            # Cancel button
            cancel_button = toga.Button(
                "Cancel",
                on_press=lambda widget: self.close_selection_dialog(selection_window, None),
                style=Pack(margin=(0, 5, 0, 0), width=100)
            )
            
            # Update button (initially disabled)
            self.update_button = toga.Button(
                "Update",
                on_press=lambda widget: self.close_selection_dialog(selection_window, excel_files),
                style=Pack(margin=(0, 0, 0, 0), width=100)
            )
            
            # Initially disable update button
            self.update_button.enabled = False
            
            # Add selection change handler to enable/disable update button
            self.file_list_selection.on_select = self.on_file_list_selection_change
            
            # Add components
            button_box.add(cancel_button)
            button_box.add(self.update_button)
            
            main_box.add(title_label)
            main_box.add(count_label)
            main_box.add(self.file_list_selection)
            main_box.add(button_box)
            
            # Set content and show
            selection_window.content = main_box
            selection_window.show()
            
            # Store reference for later use
            self.selection_window = selection_window
            self.selected_excel_files = excel_files
            
        except Exception as e:
            self.print_to_console(f"Error creating selection dialog: {e}")
            # Fallback to simple console-based selection
            await self.show_simple_excel_selection(excel_files)
    
    def on_file_list_selection_change(self, widget):
        """Handle file list selection change to enable/disable update button"""
        try:
            # Enable update button when a file is selected
            if self.file_list_selection.selection is not None:
                self.update_button.enabled = True
            else:
                self.update_button.enabled = False
        except Exception as e:
            self.print_to_console(f"Error handling selection change: {e}")
    
    def close_selection_dialog(self, window, excel_files):
        """Close the selection dialog and handle the result"""
        try:
            if excel_files is None:
                # Cancel was pressed
                self.print_to_console("File selection cancelled")
                window.close()
                return
            
            # Get selected file - need to find the index manually since selection returns Row object
            if hasattr(self, 'file_list_selection') and self.file_list_selection.selection is not None:
                # Find the selected item's index in the data
                selected_row = self.file_list_selection.selection
                selected_index = None
                
                # Find the index by comparing with the data
                for i, data_item in enumerate(self.file_list_selection.data):
                    if (data_item.title == selected_row.title and 
                        data_item.subtitle == selected_row.subtitle):
                        selected_index = i
                        break
                
                if selected_index is not None:
                    selected_file = excel_files[selected_index]
                    
                    # Close window first
                    window.close()
                    
                    # Process the selection
                    asyncio.create_task(self.update_selected_excel_file(selected_file))
                else:
                    self.print_to_console("Could not determine selected file")
                    window.close()
            else:
                self.print_to_console("No file selected")
                window.close()
                
        except Exception as e:
            self.print_to_console(f"Error processing file selection: {e}")
            window.close()
    
    async def show_simple_excel_selection(self, excel_files):
        """Fallback simple selection method if custom dialog fails"""
        self.print_to_console(f"\nFound {len(excel_files)} Excel files:")
        for i, file_info in enumerate(excel_files, 1):
            self.print_to_console(f"{i}. {file_info['name']}")
        
        dialog_message = f"Found {len(excel_files)} Excel files. Please check the console output for the list."
        await self.main_window.dialog(toga.InfoDialog("Excel Files Found", dialog_message))
    
    async def update_selected_excel_file(self, selected_file):
        """Process selected Excel file - download, open, and extract tables"""
        self.print_to_console(f"Processing Excel file: {selected_file['name']}")
        self.status_label.text = f"Processing: {selected_file['name']}"
        self.status_label.style.color = "orange"
        
        try:
            # Use Excel manager with context manager for automatic cleanup
            with ExcelManager(self.sharepoint_client) as excel_manager:
                # Download and open the file
                self.print_to_console("Downloading and opening Excel file...")
                success = await excel_manager.download_and_open_excel_file(selected_file)
                
                if not success:
                    self.print_to_console("Failed to download or open Excel file")
                    self.status_label.text = "Error: Failed to open Excel file"
                    self.status_label.style.color = "red"
                    return
                
                self.print_to_console("Excel file opened successfully")
                
                # Extract available tables
                self.print_to_console("Extracting available tables and worksheets...")
                tables = excel_manager.get_available_tables()
                
                if not tables:
                    self.print_to_console("No tables or data found in Excel file")
                    self.status_label.text = "No tables found in Excel file"
                    self.status_label.style.color = "orange"
                    return
                
                # Display table information
                self.print_to_console(f"\nFound {len(tables)} tables/worksheets:")
                self.print_to_console("-" * 50)
                
                for i, table in enumerate(tables, 1):
                    self.print_to_console(f"{i}. {table['description']}")
                    if table['sample_headers']:
                        headers = ", ".join(table['sample_headers'])
                        self.print_to_console(f"   Sample columns: {headers}")
                    self.print_to_console("")
                
                self.status_label.text = f"Found {len(tables)} tables in {selected_file['name']}"
                self.status_label.style.color = "green"
                
                # TODO: Here you could add functionality to:
                # 1. Let user select which table to update
                # 2. Show table preview
                # 3. Update table data
                # 4. Save changes back to SharePoint
                
                self.print_to_console("Table extraction completed. Ready for next steps.")
                
        except Exception as e:
            error_msg = str(e)
            self.print_to_console(f"Error processing Excel file: {error_msg}")
            self.status_label.text = f"Error: {error_msg[:50]}..."
            self.status_label.style.color = "red"