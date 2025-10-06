import argparse
import os
import sys
from pathlib import Path

from PyInstaller.__main__ import run


def create_build_config():
    """Create build configuration for all supported applications."""
    # Get the absolute path to the src directory
    src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "vba_edit")

    return {
        "word": {
            "script": os.path.join(src_dir, "word_vba.py"),
            "name": "word-vba",
            "description": "Word VBA CLI tool",
        },
        "excel": {
            "script": os.path.join(src_dir, "excel_vba.py"),
            "name": "excel-vba",
            "description": "Excel VBA CLI tool",
        },
        "access": {
            "script": os.path.join(src_dir, "access_vba.py"),
            "name": "access-vba",
            "description": "Access VBA CLI tool",
        },
        "powerpoint": {
            "script": os.path.join(src_dir, "powerpoint_vba.py"),
            "name": "powerpoint-vba",
            "description": "PowerPoint VBA CLI tool",
        },
    }


def get_version():
    """Extract version from pyproject.toml."""
    try:
        import tomli
    except ImportError:
        try:
            import tomllib as tomli
        except ImportError:
            return "0.4.0"

    pyproject_path = Path(__file__).parent / "pyproject.toml"
    if pyproject_path.exists():
        with open(pyproject_path, "rb") as f:
            data = tomli.load(f)
            return data.get("project", {}).get("version", "0.4.0")
    return "0.4.0"


def create_version_file(exe_name: str, app_description: str, output_dir: str = "."):
    """Create a version file for PyInstaller to embed Windows properties."""
    version = get_version()
    # Convert version string to numeric format (e.g., "0.4.0-rc2" -> "0.4.0.0")
    version_numeric = version.split("-")[0]  # Remove any suffix like -rc2
    parts = version_numeric.split(".")
    while len(parts) < 4:
        parts.append("0")
    file_version = ".".join(parts[:4])

    version_file_content = f"""# UTF-8
#
# For more details about fixed file info:
# https://learn.microsoft.com/en-us/windows/win32/menurc/versioninfo-resource

VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({parts[0]}, {parts[1]}, {parts[2]}, {parts[3]}),
    prodvers=({parts[0]}, {parts[1]}, {parts[2]}, {parts[3]}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'Markus Killer'),
        StringStruct(u'FileDescription', u'{app_description}'),
        StringStruct(u'FileVersion', u'{file_version}'),
        StringStruct(u'InternalName', u'{exe_name}'),
        StringStruct(u'LegalCopyright', u'Copyright (c) 2024-2025 Markus Killer (BSD-3-Clause License)'),
        StringStruct(u'OriginalFilename', u'{exe_name}.exe'),
        StringStruct(u'ProductName', u'vba-edit'),
        StringStruct(u'ProductVersion', u'{version}')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""

    version_file_path = Path(output_dir) / f"version_{exe_name}.txt"
    version_file_path.parent.mkdir(parents=True, exist_ok=True)
    with open(version_file_path, "w", encoding="utf-8") as f:
        f.write(version_file_content)

    return str(version_file_path)


def build_executable(app_name: str, config: dict, src_dir: str, additional_args: list = None):
    """Build a single executable using PyInstaller."""
    script_path = config["script"]
    exe_name = config["name"]
    description = config["description"]

    # Check if script exists
    if not os.path.exists(script_path):
        print(f"Error: Script not found: {script_path}")
        return False

    # Create version file for Windows properties
    version_file = create_version_file(exe_name, description)

    # Base PyInstaller arguments
    args = [
        script_path,
        "--onefile",
        f"--name={exe_name}",
        "--clean",
        "--paths",
        src_dir,
        f"--version-file={version_file}",
    ]

    # Add any additional arguments
    if additional_args:
        args.extend(additional_args)

    print(f"Building {exe_name}.exe...")
    print(f"  Script: {script_path}")
    print(f"  Arguments: {' '.join(args)}")

    try:
        run(args)
        print(f"✓ Successfully built {exe_name}.exe")

        # Clean up version file
        try:
            Path(version_file).unlink()
        except Exception:
            pass

        return True
    except Exception as e:
        print(f"✗ Failed to build {exe_name}.exe: {e}")
        return False


def main():
    """Main entry point for the build script."""
    parser = argparse.ArgumentParser(
        description="Build VBA-edit executables using PyInstaller",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python create_binaries.py                    # Build all executables
    python create_binaries.py --apps excel      # Build only excel-vba.exe
    python create_binaries.py --apps word excel # Build word-vba.exe and excel-vba.exe
    python create_binaries.py --list            # List available applications
    python create_binaries.py --debug           # Build with debug info (no UPX compression)
        """,
    )

    # Get build configuration
    build_config = create_build_config()
    available_apps = list(build_config.keys())

    parser.add_argument(
        "--apps",
        nargs="*",
        choices=available_apps,
        help=f"Specify which applications to build. Available: {', '.join(available_apps)}. If not specified, builds all.",
    )

    parser.add_argument("--list", action="store_true", help="List all available applications and exit")

    parser.add_argument(
        "--debug", action="store_true", help="Build with debug information (larger executables, faster build)"
    )

    parser.add_argument("--output-dir", help="Specify output directory for executables (default: dist/)")

    args = parser.parse_args()

    # Handle --list option
    if args.list:
        print("Available applications:")
        for app_name, config in build_config.items():
            print(f"  {app_name:12} -> {config['name']}.exe ({config['description']})")
        return

    # Determine which apps to build
    if args.apps:
        apps_to_build = args.apps
    else:
        apps_to_build = available_apps

    print(f"Building executables for: {', '.join(apps_to_build)}")
    print("-" * 50)

    # Get src directory
    src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "vba_edit")

    # Prepare additional PyInstaller arguments
    additional_args = []
    if args.output_dir:
        additional_args.extend(["--distpath", args.output_dir])

    if args.debug:
        additional_args.extend(["--debug", "all", "--console"])

    # Build executables
    successful_builds = []
    failed_builds = []

    for app_name in apps_to_build:
        if app_name not in build_config:
            print(f"Warning: Unknown application '{app_name}', skipping...")
            continue

        config = build_config[app_name]
        success = build_executable(app_name, config, src_dir, additional_args)

        if success:
            successful_builds.append(app_name)
        else:
            failed_builds.append(app_name)

        print()  # Add blank line between builds

    # Summary
    print("=" * 50)
    print("Build Summary:")
    if successful_builds:
        print(f"✓ Successfully built: {', '.join(successful_builds)}")
    if failed_builds:
        print(f"✗ Failed to build: {', '.join(failed_builds)}")
        sys.exit(1)

    print("All requested executables built successfully!")


if __name__ == "__main__":
    main()
