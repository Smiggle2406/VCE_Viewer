import os
import subprocess
import shutil
from pathlib import Path
import site
import PyQt6


def find_pyqt6_resource_path(resource):
    """Find the path to a PyQt6 resource (e.g., plugins, translations)."""
    pyqt6_path = os.path.dirname(PyQt6.__file__)
    possible_paths = [
        os.path.join(pyqt6_path, "Qt6", resource),  # Standard PyQt6 structure
        os.path.join(pyqt6_path, resource),  # Alternative structure
        os.path.join(site.getsitepackages()[0], "PyQt6", "Qt6", resource),
        os.path.join(site.getsitepackages()[0], "PyQt6", resource)
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(f"Could not find PyQt6 {resource} directory")


def package_program():
    # Define paths
    script_path = "vce_viewer.py"  # Main script name
    output_dir = Path("dist")  # PyInstaller output directory
    build_dir = Path("build")  # PyInstaller build directory
    spec_dir = Path("spec")  # Directory for .spec file
    icon_path = Path.cwd() / "app_icon.icns"  # Path to .icns file in project root
    uploads_dir = Path.cwd() / "uploaded_reports"  # Absolute path for uploaded_reports

    # Ensure the main script exists
    if not Path(script_path).exists():
        print(f"Error: {script_path} not found.")
        return

    # Create uploaded_reports directory if it doesn't exist
    uploads_dir.mkdir(exist_ok=True)
    # Create a placeholder file to ensure the directory is included
    placeholder_path = uploads_dir / ".placeholder"
    placeholder_path.touch(exist_ok=True)

    # Clean previous builds
    for dir_path in [output_dir, build_dir, spec_dir]:
        if dir_path.exists():
            shutil.rmtree(dir_path)

    # Create spec directory
    spec_dir.mkdir(exist_ok=True)

    # Find PyQt6 plugins and translations paths
    try:
        plugins_path = find_pyqt6_resource_path("plugins")
        translations_path = find_pyqt6_resource_path("translations")
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    # PyInstaller command for macOS
    pyinstaller_cmd = [
        "pyinstaller",
        "--noconfirm",  # Overwrite without confirmation
        "--onedir",  # Create a directory containing the executable
        "--windowed",  # Create a macOS .app bundle
        f"--distpath={output_dir}",
        f"--workpath={build_dir}",
        f"--specpath={spec_dir}",
        # Add PyQt6 plugins and translations
        f"--add-data={plugins_path}:PyQt6/Qt6/plugins",
        f"--add-data={translations_path}:PyQt6/Qt6/translations",
        # Add uploaded_reports directory structure
        f"--add-data={uploads_dir}:uploaded_reports",
        # Optional: Icon for the .app bundle
        *(["--icon", str(icon_path)] if icon_path.exists() else []),
        # Hidden imports to ensure all dependencies are included
        "--hidden-import=PyQt6.QtPdf",
        "--hidden-import=PyQt6.QtPdfWidgets",
        "--hidden-import=bs4",
        "--hidden-import=requests",
        # Set the name of the .app bundle
        "--name=VCEViewer",
        # Main script
        script_path
    ]

    # Run PyInstaller
    print("Running PyInstaller...")
    try:
        subprocess.run(pyinstaller_cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"PyInstaller failed: {e}")
        return

    # Post-processing: Create a README for LibreOffice dependency
    readme_content = """VCE Exam Report Viewer
=====================
This application requires LibreOffice to convert .doc/.docx files to PDF.
Please install LibreOffice from https://www.libreoffice.org/download/download/
and ensure it is available in your PATH as 'soffice' or 'libreoffice'.

To run the application:
1. Double-click VCEViewer.app to launch.
2. If macOS shows a security warning, right-click the .app, select 'Open', and confirm.
3. Ensure LibreOffice is installed for .doc/.docx conversion functionality.
"""
    readme_path = output_dir / "VCEViewer" / "README.txt"
    with open(readme_path, "w") as f:
        f.write(readme_content)

    print("Packaging complete! The application bundle is located at:")
    print(f"{output_dir / 'VCEViewer.app'}")
    print("Note: Users must install LibreOffice separately for .doc/.docx conversion.")


if __name__ == "__main__":
    package_program()
