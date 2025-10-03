from importlib.metadata import PackageNotFoundError, version

package_name = "vba-edit"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    __version__ = "unknown"

package_version = __version__
