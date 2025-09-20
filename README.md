# SharePoint Excel Manager

A cross-platform GUI application built with Python and Toga for managing Excel files stored in Microsoft Teams SharePoint sites.

## Features

- 🖥️ Cross-platform GUI (Windows, macOS, Linux)
- 🔗 Connect to Teams SharePoint sites
- 📊 Browse and manage Excel files
- 💾 Download and upload Excel files
- ⚙️ Save connection settings
- 🔒 Secure authentication

## Prerequisites

- Python 3.8 or higher
- Microsoft 365 account with access to Teams SharePoint
- Git (for development)

## Quick Start

### Windows

1. Clone the repository:
   ```cmd
   git clone <your-repo-url>
   cd sharepoint-excel-manager
   ```

2. Run the installation script:
   ```cmd
   scripts\install_windows.bat
   ```

3. Start the application:
   ```cmd
   scripts\run_windows.bat
   ```

### macOS

1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd sharepoint-excel-manager
   ```

2. Make the installation script executable and run it:
   ```bash
   chmod +x scripts/install_macos.sh
   ./scripts/install_macos.sh
   ```

3. Start the application:
   ```bash
   ./scripts/run_macos.sh
   ```

## Manual Installation

If you prefer to install manually:

1. Create a virtual environment:
   ```bash
   python -m venv venv
   ```

2. Activate the virtual environment:
   - Windows: `venv\Scripts\activate.bat`
   - macOS/Linux: `source venv/bin/activate`

3. Install the project:
   ```bash
   pip install -e .
   ```

4. Run the application:
   ```bash
   python -m sharepoint_excel_manager.main
   ```

## Configuration

On first run, you'll need to configure:

1. **Team SharePoint URL**: The URL of your Teams SharePoint site
   - Format: `https://yourorganization.sharepoint.com/sites/yourteam`
   
2. **Document Folder Path**: The path to the document folder (optional)
   - Leave empty for the default "Shared Documents" folder
   - Format: `/sites/yourteam/Shared Documents/YourFolder`

3. **Authentication**: You'll be prompted for your Microsoft 365 credentials

## Project Structure

```
sharepoint-excel-manager/
├── src/
│   └── sharepoint_excel_manager/
│       ├── __init__.py
│       ├── main.py              # Application entry point
│       ├── gui.py               # Toga GUI implementation
│       └── sharepoint_client.py # SharePoint integration
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
```

## Development

### Setting up Development Environment

1. Install development dependencies:
   ```bash
   pip install -e ".[dev]"
   ```

2. Run tests:
   ```bash
   pytest
   ```

3. Format code:
   ```bash
   black src/
   ```

4. Lint code:
   ```bash
   flake8 src/
   ```

### Adding Features

The application is structured with clear separation of concerns:

- `gui.py`: Handle all UI components and user interactions
- `sharepoint_client.py`: Manage SharePoint API calls and authentication
- `main.py`: Application entry point and initialization

### Building for Distribution

To create standalone executables:

```bash
# Install briefcase for packaging
pip install briefcase

# Create platform-specific packages
briefcase create
briefcase build
briefcase package
```

## Troubleshooting

### Common Issues

**Authentication Failed**
- Ensure you're using your full Microsoft 365 email address
- Check if your organization requires multi-factor authentication
- Verify your SharePoint site URL is correct

**Connection Timeout**
- Check your internet connection
- Verify the SharePoint site is accessible via browser
- Some corporate networks may block certain connections

**Python Version Issues**
- Ensure Python 3.8 or higher is installed
- Check that `python` command points to correct version
- On some systems, use `python3` instead of `python`

**Module Not Found Errors**
- Make sure you've activated the virtual environment
- Re-run the installation script if dependencies are missing

### Getting Help

1. Check the [Issues](../../issues) section for known problems
2. Create a new issue with:
   - Your operating system
   - Python version
   - Complete error message
   - Steps to reproduce

## Security Considerations

- Credentials are not stored permanently
- Configuration files contain only URLs and folder paths
- All SharePoint communications use HTTPS
- Consider using Azure AD app registration for production deployments

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes and add tests
4. Run the test suite: `pytest`
5. Format your code: `black src/`
6. Commit your changes: `git commit -am 'Add feature'`
7. Push to the branch: `git push origin feature-name`
8. Create a Pull Request

## Dependencies

### Core Dependencies
- **Toga**: Cross-platform GUI framework
- **Office365-REST-Python-Client**: SharePoint API client
- **openpyxl**: Excel file manipulation
- **requests**: HTTP requests
- **msal**: Microsoft Authentication Library

### Development Dependencies
- **pytest**: Testing framework
- **black**: Code formatter
- **flake8**: Code linter
- **mypy**: Type checker

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Changelog

### Version 1.0.0
- Initial release
- Basic SharePoint connection and file browsing
- Cross-platform GUI with Toga
- Windows and macOS installation scripts

## Roadmap

- [ ] Excel file editing capabilities
- [ ] Batch file operations
- [ ] Advanced filtering and search
- [ ] Integration with other Office 365 services
- [ ] Enhanced error handling and logging
- [ ] Automated testing on CI/CD platforms