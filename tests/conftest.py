"""Pytest configuration file."""

import pytest

from tests.office_processes import close_office_apps_started_since, office_pids_at_session_start

# Office applications that were running before the tests started are never closed by the cleanup.
_office_pids_before_tests = office_pids_at_session_start()


def pytest_addoption(parser):
    parser.addoption(
        "--apps",
        action="store",
        default="all",
        help="Comma-separated list of apps to test (excel,word,access,powerpoint) or 'all' for all available apps. "
        "Note: Access is excluded from 'all' by default due to interactive requirements.",
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
        "gui: mark test as requiring an interactive display (skipped in CI)",
    ]
    for marker in markers:
        config.addinivalue_line("markers", marker)


def pytest_sessionstart(session):
    """Check Office application safety before starting integration tests.

    This runs once at the very beginning of the test session, before any tests are collected.
    Only runs if integration tests are being executed (not in CI pipeline).
    """
    # Check if we're running integration tests by looking at command line args
    config = session.config

    # Skip safety check if:
    # 1. No test files specified (running all tests, likely in CI)
    # 2. Only running unit tests (no cli/ or integration markers)
    # 3. Running specific non-integration test files

    # Get the test paths being run
    test_args = config.args if config.args else []

    # Skip if no args (default test discovery) - likely CI
    if not test_args:
        return

    # Check if any integration test paths are in the arguments
    has_integration_tests = any(
        "cli" in str(arg).lower()
        or "integration" in str(arg).lower()
        or "test_excel_integration" in str(arg)
        or "test_word" in str(arg)
        or "test_powerpoint" in str(arg)
        for arg in test_args
    )

    if not has_integration_tests:
        return

    # Import the safety check function
    try:
        from tests.cli.helpers import check_office_apps_are_safe_to_use

        is_safe, message = check_office_apps_are_safe_to_use()

        if not is_safe:
            # Print warning and abort
            print(message)
            pytest.exit(
                "Integration tests aborted - Office applications have open documents. "
                "Close all Office windows or set PYTEST_ALLOW_OFFICE_FORCE_QUIT=1 to override.",
                returncode=1,
            )
    except ImportError:
        # If helpers not available, skip the check
        pass


def pytest_generate_tests(metafunc):
    """Dynamically parametrize vba_app based on command line options."""
    if "vba_app" in metafunc.fixturenames:
        # Import here to avoid circular import issues
        from tests.cli.helpers import get_installed_apps

        # Get selected apps from command line
        apps_option = metafunc.config.getoption("--apps")
        include_access_interactive = metafunc.config.getoption("--include-access-interactive")

        if apps_option.lower() == "all":
            # By default, exclude Access unless explicitly enabled via flag
            if include_access_interactive:
                selected_apps = ["excel", "word", "access", "powerpoint"]
            else:
                selected_apps = ["excel", "word", "powerpoint"]
        else:
            selected_apps = [app.strip().lower() for app in apps_option.split(",")]
            valid_apps = ["excel", "word", "access", "powerpoint"]
            invalid_apps = [app for app in selected_apps if app not in valid_apps]
            if invalid_apps:
                raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")

            # If Access is explicitly selected, require the interactive flag
            if "access" in selected_apps and not include_access_interactive:
                raise ValueError(
                    "Access tests require user interaction. "
                    "Please add --include-access-interactive flag to enable Access tests."
                )

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
            if hasattr(item, "callspec") and "vba_app" in item.callspec.params:
                if item.callspec.params["vba_app"] == "access":
                    item.add_marker(
                        pytest.mark.skip(
                            reason="Access requires user interaction (use --include-access-interactive to enable)"
                        )
                    )
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


def _close_office_apps_started_by_tests(when: str) -> None:
    """Drop cached COM references, then close the Office applications the tests started."""
    # Clear the app instance cache FIRST, before touching any processes.
    # This prevents trying to use dead COM references later.
    try:
        from tests.cli import helpers

        helpers._app_instances.clear()
    except Exception as e:
        print(f"  Warning: Could not clear app cache: {e}")

    # Skip a graceful shutdown: the applications may be in a bad state after a failed test.
    closed = close_office_apps_started_since(_office_pids_before_tests)
    if closed:
        print(f"\nClosed Office applications started by the tests ({when}): {', '.join(closed)}")


@pytest.fixture(scope="module", autouse=True)
def cleanup_office_apps_per_module():
    """Close the Office applications a test module started.

    Applications that were already running when the session began are left alone.
    """
    yield
    _close_office_apps_started_by_tests("end of test module")


@pytest.fixture(scope="session", autouse=True)
def cleanup_office_apps():
    """Close the Office applications the test session started.

    Applications that were already running when the session began are left alone.
    """
    yield
    _close_office_apps_started_by_tests("end of test session")
