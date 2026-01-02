"""
File utility functions for the survey data processor.

This module provides functions for file operations like versioning.
"""

from pathlib import Path


def get_next_version_filename(base_name: str = "test_processed_pilot") -> str:
    """
    Gets the next versioned filename in the output directory.

    Args:
        base_name (str): Base name for the file (without extension)

    Returns:
        str: Full path to the next versioned file
    """
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    version = 1
    while True:
        filename = f"{base_name}v{version}.xlsx"
        filepath = output_dir / filename
        if not filepath.exists():
            return str(filepath)
        version += 1
