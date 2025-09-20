SharePoint Excel Manager
A cross-platform GUI application built with Python and Toga for managing Excel files stored in Microsoft Teams SharePoint sites.

Features
🖥️ Cross-platform GUI (Windows, macOS, Linux)
🔗 Connect to Teams SharePoint sites
📊 Browse and manage Excel files
💾 Download and upload Excel files
⚙️ Persistent settings storage with auto-save
🪟 Window state preservation (size and position)
🔒 Secure authentication
📁 Recent connections management
Prerequisites
Python 3.8 or higher
Microsoft 365 account with access to Teams SharePoint
Git (for development)
Quick Start
Windows
Clone the repository:
cmd
   git clone <your-repo-url>
   cd sharepoint-excel-manager
Run the installation script:
cmd
   scripts\install_windows.bat
Start the application:
cmd
   scripts\run_windows.bat
macOS
Clone the repository:
bash
   git clone <your-repo-url>
   cd sharepoint-excel-manager
Make the installation script executable and run it:
bash
   chmod +x scripts/install_macos.sh
   ./scripts/install_macos.sh
Start the application:
bash
   ./scripts/run_macos.sh
Manual Installation
If you prefer to install manually:

Create a virtual environment:
bash
   python -m venv venv
Activate the virtual environment:
Windows: venv\Scripts\activate.bat
macOS/Linux: source venv/bin/activate
Install the project:
bash
   pip install -e .
Run the application:
bash
   python -m sharepoint_excel_manager.main
Configuration
The application automatically saves and restores your settings between sessions:

Automatic Settings
Team SharePoint URL and Document Folder: Saved automatically as you type
Window size and position: Restored when you restart the app
User preferences: Theme, auto-connect options, etc.
Settings Storage Locations
Windows: %APPDATA%\SharePointExcelManager\settings.json
macOS: ~/Library/Application Support/SharePointExcelManager/settings.json
Linux: ~/.config/SharePointExcelManager/settings.json
First-Time Setup
On first run, configure:

Team SharePoint URL: The URL of your Teams SharePoint site
Format: https://yourorganization.sharepoint.com/sites/yourteam
Document Folder Path: The path to the document folder (optional)
Leave empty for the default "Shared Documents" folder
Format: /sites/yourteam/Shared Documents/YourFolder
Authentication: You'll be prompted for your Microsoft 365 credentials
Settings Management
Click the Settings button to view current configuration
Use Save Config to manually save settings (though auto-save is enabled)
Settings are automatically exported/imported when needed
Project Structure
sharepoint-excel-manager/
├── src/
│   └── sharepoint_excel_manager/
│       ├── __init__.py
│       ├── main.py              # Application entry point
│       ├── gui.py               # Toga GUI implementation
│       ├── sharepoint_client.py # SharePoint integration
│       └── settings.py          # Settings management and persistence
├── scripts/
│   ├── install_windows.bat      # Windows installation script
│   ├── run_windows.bat          # Windows run script
│   ├── install_macos.sh         # macOS installation script
│   └── run_macos.sh             # macOS run script
├── tests/                       # Test files
├── docs/                        # Documentation
├── pyproject.toml              # Project configuration
├── requirements.txt            # Python dependencies
└── README.md                   # This file
Development
Setting up Development Environment
Install development dependencies:
bash
   pip install -e ".[dev]"
Run tests:
bash
   pytest
Format code:
bash
   black src/
Lint code:
bash
   flake8 src/
Adding Features
The application is structured with clear separation of concerns:

gui.py: Handle all UI components and user interactions
sharepoint_client.py: Manage SharePoint API calls and authentication
main.py: Application entry point and initialization
Building for Distribution
To create standalone executables:

bash
# Install briefcase for packaging
pip install briefcase

# Create platform-specific packages
briefcase create
briefcase build
briefcase package
Troubleshooting
Common Issues
Authentication Failed

Ensure you're using your full Microsoft 365 email address
Check if your organization requires multi-factor authentication
Verify your SharePoint site URL is correct
Connection Timeout

Check your internet connection
Verify the SharePoint site is accessible via browser
Some corporate networks may block certain connections
Python Version Issues

Ensure Python 3.8 or higher is installed
Check that python command points to correct version
On some systems, use python3 instead of python
Module Not Found Errors

Make sure you've activated the virtual environment
Re-run the installation script if dependencies are missing
Getting Help
Check the Issues section for known problems
Create a new issue with:
Your operating system
Python version
Complete error message
Steps to reproduce
Security Considerations
Credentials are not stored permanently
Configuration files contain only URLs and folder paths
All SharePoint communications use HTTPS
Consider using Azure AD app registration for production deployments
Contributing
Fork the repository
Create a feature branch: git checkout -b feature-name
Make your changes and add tests
Run the test suite: pytest
Format your code: black src/
Commit your changes: git commit -am 'Add feature'
Push to the branch: git push origin feature-name
Create a Pull Request
Dependencies
Core Dependencies
Toga: Cross-platform GUI framework
Office365-REST-Python-Client: SharePoint API client
openpyxl: Excel file manipulation
requests: HTTP requests
msal: Microsoft Authentication Library
Development Dependencies
pytest: Testing framework
black: Code formatter
flake8: Code linter
mypy: Type checker
License
This project is licensed under the MIT License - see the LICENSE file for details.

Changelog
Version 1.0.0
Initial release
Basic SharePoint connection and file browsing
Cross-platform GUI with Toga
Windows and macOS installation scripts
Roadmap
 Excel file editing capabilities
 Batch file operations
 Advanced filtering and search
 Integration with other Office 365 services
 Enhanced error handling and logging
 Automated testing on CI/CD platforms
