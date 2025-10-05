"""Pytest configuration file."""

import pytest
import win32com.client

# Track which Office apps were actually used by tests (for safe cleanup)
_office_apps_used = set()


def pytest_addoption(parser):
    parser.addoption(
        "--apps",
        action="store",
        default="all",
        help="Comma-separated list of apps to test (excel,word,access,powerpoint) or 'all' for all available apps",
    )
    parser.addoption(
        "--include-access-interactive",
        action="store_true",
        default=False,
        help="Include Access tests that require user interaction (clicking OK on Save As dialogs). "
             "Without this flag, Access tests are skipped by default.",
    )


def pytest_configure(config):
    """Register custom marks."""
    markers = [
        "excel: mark test as Excel-specific",
        "word: mark test as Word-specific",
        "access: mark test as Access-specific",
        "powerpoint: mark test as PowerPoint-specific",
        "office: mark test as general Office test",
        "com: marks tests that require COM initialization",
        "integration: mark test as integration test",
        "skip_access: skip Access variants of parameterized tests (requires user interaction)",
    ]
    for marker in markers:
        config.addinivalue_line("markers", marker)


def pytest_generate_tests(metafunc):
    """Dynamically parametrize vba_app based on command line options."""
    if "vba_app" in metafunc.fixturenames:
        # Import here to avoid circular import issues
        from tests.cli.helpers import get_installed_apps

        # Get selected apps from command line
        apps_option = metafunc.config.getoption("--apps")
        if apps_option.lower() == "all":
            selected_apps = ["excel", "word", "access", "powerpoint"]
        else:
            selected_apps = [app.strip().lower() for app in apps_option.split(",")]
            valid_apps = ["excel", "word", "access", "powerpoint"]
            invalid_apps = [app for app in selected_apps if app not in valid_apps]
            if invalid_apps:
                raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")

        apps = get_installed_apps(selected_apps=selected_apps)
        metafunc.parametrize("vba_app", apps, ids=lambda x: f"{x}-vba")


def pytest_collection_modifyitems(config, items):
    """Skip tests based on selected apps and other markers."""
    apps_option = config.getoption("--apps")
    include_access_interactive = config.getoption("--include-access-interactive")
    
    if apps_option.lower() == "all":
        selected_apps = ["excel", "word", "access", "powerpoint"]
    else:
        selected_apps = [app.strip().lower() for app in apps_option.split(",")]

    # Skip tests that don't match selected apps
    for item in items:
        # Skip Access tests if they have the skip_access marker and flag is not set
        if item.get_closest_marker("skip_access") and not include_access_interactive:
            # Check if this is an Access test variant (via parametrization)
            if hasattr(item, 'callspec') and 'vba_app' in item.callspec.params:
                if item.callspec.params['vba_app'] == 'access':
                    item.add_marker(pytest.mark.skip(
                        reason="Access requires user interaction (use --include-access-interactive to enable)"
                    ))
                    continue
        
        # Check if test has app-specific markers
        test_apps = []
        if item.get_closest_marker("excel"):
            test_apps.append("excel")
        if item.get_closest_marker("word"):
            test_apps.append("word")
        if item.get_closest_marker("access"):
            test_apps.append("access")
        if item.get_closest_marker("powerpoint"):
            test_apps.append("powerpoint")

        # If test has app-specific markers and none match selected apps, skip it
        if test_apps and not any(app in selected_apps for app in test_apps):
            item.add_marker(pytest.mark.skip(reason=f"Test requires {test_apps} but only {selected_apps} selected"))


@pytest.fixture
def vba_app():
    """VBA application fixture - will be parametrized by pytest_generate_tests."""
    # This fixture body will never execute because pytest_generate_tests
    # will parametrize it with actual values
    pass


@pytest.fixture
def selected_apps(request):
    """Get the list of apps selected for testing."""
    apps_option = request.config.getoption("--apps")
    if apps_option.lower() == "all":
        return ["excel", "word", "access", "powerpoint"]
    else:
        # Parse comma-separated list and validate
        apps = [app.strip().lower() for app in apps_option.split(",")]
        valid_apps = ["excel", "word", "access", "powerpoint"]
        invalid_apps = [app for app in apps if app not in valid_apps]
        if invalid_apps:
            raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")
        return apps


@pytest.fixture
def excel_only(request):
    """Check if running in Excel-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["excel"]


@pytest.fixture
def word_only(request):
    """Check if running in Word-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["word"]


@pytest.fixture
def access_only(request):
    """Check if running in Access-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["access"]


@pytest.fixture
def powerpoint_only(request):
    """Check if running in PowerPoint-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["powerpoint"]


@pytest.fixture(scope="session", autouse=True)
def cleanup_office_apps():
    """Clean up Office applications after all tests."""
    yield

    import warnings
    import subprocess
    import time

    for app_name in ["Word.Application", "Excel.Application", "PowerPoint.Application", "Access.Application"]:
        # Only force-kill if the app was actually used by tests
        # This prevents data loss if tests were aborted by safety checks
        if app_name not in _office_apps_used:
            continue
            
        try:
            app = win32com.client.GetObject(Class=app_name)
            # Suppress all alerts
            app.DisplayAlerts = False
            
            # Mark all open documents/workbooks as saved to prevent prompts
            try:
                if hasattr(app, 'Workbooks'):  # Excel
                    count = app.Workbooks.Count
                    for i in range(count, 0, -1):
                        try:
                            app.Workbooks(i).Saved = True
                        except:
                            pass
                elif hasattr(app, 'Documents'):  # Word
                    count = app.Documents.Count
                    for i in range(count, 0, -1):
                        try:
                            app.Documents(i).Saved = True
                        except:
                            pass
            except:
                pass
            
            # Try to quit gracefully with a timeout
            try:
                app.Quit()
                time.sleep(0.3)  # Give it a moment to close
            except:
                pass
                
        except Exception:
            pass  # Already closed - this is fine
    
    # Force kill any remaining Office processes (in case Quit() hung)
    # Only kill processes for apps that were actually used by tests
    time.sleep(0.2)
    app_to_process = {
        "Excel.Application": "EXCEL",
        "Word.Application": "WINWORD",
        "PowerPoint.Application": "POWERPNT",
        "Access.Application": "MSACCESS"
    }
    
    for app_name, process_name in app_to_process.items():
        # Only force-kill if the app was used by tests
        if app_name not in _office_apps_used:
            continue
            
        try:
            # Use /T to kill child processes too
            result = subprocess.run(
                ["taskkill", "/F", "/IM", f"{process_name}.EXE", "/T"], 
                capture_output=True, 
                timeout=1,
                text=True
            )
            # Only print if process was actually killed (not "not found")
            if result.returncode == 0:
                print(f"Force-killed {process_name}")
        except:
            pass
