"""
Excel formula calculation utilities.

This module provides functions to calculate Excel formulas using win32com.
"""

import logging
import os
import time

from ..constants import WIN32_AVAILABLE

if WIN32_AVAILABLE:
    import win32com.client


def calculate_excel_formulas(file_path: str) -> None:
    """
    Opens Excel to calculate all formulas in the workbook before proceeding with question cutting.
    This ensures that all formulas (especially column G lookups) are properly evaluated.

    Args:
        file_path (str): Path to the Excel file to calculate
    """
    if not WIN32_AVAILABLE:
        logging.warning("win32com.client not available - skipping Excel calculation. Install pywin32 for full functionality.")
        return

    try:
        logging.info("Opening Excel to calculate all formulas...")

        # Create Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # Keep Excel hidden
        excel_app.DisplayAlerts = False  # Disable alerts

        # Convert to absolute path for Excel
        abs_file_path = os.path.abspath(file_path)
        logging.info(f"Opening Excel file: {abs_file_path}")

        # Open the workbook
        workbook = excel_app.Workbooks.Open(abs_file_path)

        # Calculate all formulas with full rebuild
        excel_app.CalculateFullRebuild()
        logging.info("Triggered full formula calculation in Excel")

        # Wait for calculation to complete
        time.sleep(3)  # Give Excel time to complete all calculations
        logging.info("Calculated all formulas in Excel")

        # Save and close
        workbook.Save()
        workbook.Close()
        excel_app.Quit()

        logging.info("Excel calculation completed successfully")

    except Exception as e:
        logging.error(f"Error during Excel calculation: {e}")
        # Don't raise the error - continue with processing even if Excel calculation fails
        logging.warning("Continuing without Excel calculation - formulas may not be evaluated")
