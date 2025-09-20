#!/bin/bash

echo "SharePoint Excel Manager - macOS Runner"
echo "======================================="

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
NC='\033[0m' # No Color

# Function to print colored output
print_error() {
    echo -e "${RED}ERROR: $1${NC}"
}

print_success() {
    echo -e "${GREEN}$1${NC}"
}

# Check if virtual environment exists
if [ ! -f "venv/bin/activate" ]; then
    print_error "Virtual environment not found"
    echo "Please run ./scripts/install_macos.sh first"
    exit 1
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Check if package is installed
python -c "import sharepoint_excel_manager" 2>/dev/null
if [ $? -ne 0 ]; then
    print_error "SharePoint Excel Manager not installed"
    echo "Please run ./scripts/install_macos.sh first"
    exit 1
fi

# Run the application
echo "Starting SharePoint Excel Manager..."
echo
python -m sharepoint_excel_manager.main

# Check exit status
if [ $? -ne 0 ]; then
    echo
    print_error "Application exited with an error"
    exit 1
else
    print_success "Application closed successfully"
fi