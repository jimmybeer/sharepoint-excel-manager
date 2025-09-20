"""
Main application entry point for SharePoint Excel Manager
"""
from .gui import SharePointExcelApp


def main():
    """Main entry point for the application"""
    app = SharePointExcelApp()
    app.main_loop()
    return app


if __name__ == "__main__":
    main()
