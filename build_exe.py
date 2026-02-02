"""
Build script to create a standalone .exe file using PyInstaller
Run this script to build the executable:
    python build_exe.py
"""

import os
import subprocess
import sys

def main():
    print("Building Excel Processor Application...")
    print("-" * 50)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",  # Create a single executable file
        "--windowed",  # No console window (GUI only)
        "--name=krec_trec",  # Name of the executable
        "--icon=NONE",  # No icon (can be added later if needed)
        "--add-data=excel_processor.py;.",  # Include the processor module
        "--clean",  # Clean PyInstaller cache
        "gui_app.py"
    ]
    
    print(f"Running: {' '.join(cmd)}")
    print("-" * 50)
    
    try:
        subprocess.check_call(cmd)
        print("\n" + "=" * 50)
        print("BUILD SUCCESSFUL!")
        print("=" * 50)
        print("\nThe executable file is located at:")
        print("  dist\\krec_trec.exe")
        print("\nYou can copy this .exe file to any Windows computer")
        print("and it will run without requiring Python installation.")
        print("=" * 50)
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
