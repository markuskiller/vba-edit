"""CLI tests entry point - imports all CLI test modules.

This file serves as an entry point for pytest discovery of all CLI test modules.
The actual test implementations have been moved to the tests/cli/ directory for better organization.

To run all CLI tests:
    pytest tests/test_cli.py -v

To run specific test modules:
    pytest tests/cli/test_basic.py -v
    pytest tests/cli/test_config.py -v
    pytest tests/cli/test_integration.py -v
    pytest tests/cli/test_debugging.py -v

To run debugging tests that write option files:
    pytest tests/cli/test_debugging.py -v

Test module breakdown:
- test_basic.py: Basic CLI functionality (help, commands, options)
- test_config.py: Configuration file loading and placeholder resolution
- test_integration.py: Integration tests with real Office documents
- test_debugging.py: Debugging tests that write effective option values to files
"""

import pytest

# Import all test modules to ensure they're discovered by pytest
# The order of imports doesn't matter for test discovery, but we organize them logically
from tests.cli.test_basic import *       # Basic CLI functionality tests
from tests.cli.test_config import *      # Configuration file tests
from tests.cli.test_integration import * # Integration tests with real Office documents
from tests.cli.test_debugging import *   # Debugging tests that write option values to files


if __name__ == "__main__":
    pytest.main(["-v", __file__])
