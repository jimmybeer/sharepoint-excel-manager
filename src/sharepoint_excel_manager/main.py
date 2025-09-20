"""
Main application entry point for SharePoint Excel Manager
"""
import asyncio
import sys
from .gui import SharePointExcelApp


def main():
    """Main entry point for the application"""
    app = SharePointExcelApp()
    return app


if __name__ == "__main__":
    main().main_loop()