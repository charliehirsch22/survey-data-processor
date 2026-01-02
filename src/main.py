#!/usr/bin/env python3
"""
Survey Data Processor v4

Main entry point for command-line execution.
"""

import argparse
import logging

from .pipeline import process_excel_file
from .utils.file_utils import get_next_version_filename


def main():
    """
    Main entry point for the survey processor v4.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    parser = argparse.ArgumentParser(
        description='Process survey Excel files with improved functionality'
    )
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('--output_file', help='Path for the output processed file (optional, auto-versions if not provided)')

    args = parser.parse_args()

    # Auto-generate versioned output filename if not provided
    output_file = args.output_file if args.output_file else get_next_version_filename()

    try:
        process_excel_file(args.input_file, output_file)
        print(f"Successfully processed: {args.input_file} -> {output_file}")
    except Exception as e:
        print(f"Error: {e}")
        return 1

    return 0


if __name__ == '__main__':
    exit(main())
