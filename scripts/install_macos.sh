#!/bin/bash

echo "SharePoint Excel Manager - macOS Installation Script"
echo "===================================================="

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_error() {
    echo -e "${RED}ERROR: $1${NC}"
}

print_success() {
    echo -e "${GREEN}$1${NC}"
}

print_warning() {
    echo -e "${YELLOW}WARNING: $1${NC}"
}

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    print_error "Python 3 is not installed"
    echo "Please install Python 3.8 or later:"
    echo "  - Using Homebrew: brew install python"
    echo "  - Download from: https://python.org"
    exit 1
fi

# Display Python version
echo "Python version:"
python3 --version

# Check Python version (require 3.8+)
PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
REQUIRED_VERSION="3.8"

if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then 
    print_error "Python $REQUIRED_VERSION or higher is required. Found: $PYTHON_VERSION"
    exit 1
fi

# Check if we're in the project directory
if [ ! -f "pyproject.toml" ]; then
    print_error "Please run this script from the project root directory"
    echo "Expected to find pyproject.toml in current directory"
    exit 1
fi

# Create virtual environment
echo
echo "Creating virtual environment..."
python3 -m venv venv
if [ $? -ne 0 ]; then
    print_error "Failed to create virtual environment"
    exit 1
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
python -m pip install --upgrade pip

# Install the project in development mode
echo
echo "Installing project dependencies..."
pip install -e .
if [ $? -ne 0 ]; then
    print_error "Failed to install project dependencies"
    exit 1
fi

# Install development dependencies
echo
echo "Installing development dependencies..."
pip install -e ".[dev]"

# Make run script executable
chmod +x scripts/run_macos.sh

echo
print_success "===================================================="
print_success "Installation completed successfully!"
echo
echo "To run the application:"
echo "  1. Activate the virtual environment: source venv/bin/activate"
echo "  2. Run the application: python -m sharepoint_excel_manager.main"
echo
echo "Or use the run script: ./scripts/run_macos.sh"
print_success "===================================================="