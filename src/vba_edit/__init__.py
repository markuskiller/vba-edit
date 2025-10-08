from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
import sys
import re

package_name = "vba-edit"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    # Fallback for frozen executables (PyInstaller) - read from pyproject.toml
    try:
        # Get the directory where this module is located
        if getattr(sys, "frozen", False):
            # Running as compiled executable - pyproject.toml is in the _MEIPASS directory
            base_path = Path(sys._MEIPASS)
        else:
            # Running as normal Python script - go up to project root
            base_path = Path(__file__).parent.parent.parent

        pyproject_path = base_path / "pyproject.toml"

        if pyproject_path.exists():
            with open(pyproject_path, "r", encoding="utf-8") as f:
                content = f.read()
                # Extract version from pyproject.toml (simple regex, no toml parser needed)
                match = re.search(r'^version\s*=\s*["\']([^"\']+)["\']', content, re.MULTILINE)
                if match:
                    __version__ = match.group(1)
                else:
                    __version__ = "unknown"
        else:
            __version__ = "unknown"
    except Exception:
        __version__ = "unknown"

package_version = __version__
