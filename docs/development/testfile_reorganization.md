## ✅ Test Reorganization

The CLI test reorganization has been **successfully completed** with the following results:

### 📁 New Test Structure:

```
tests/
├── cli/
│   ├── __init__.py          # Module marker
│   ├── helpers.py           # Common test utilities (CLITester, ReferenceDocuments, fixtures)
│   ├── test_basic.py        # Basic CLI functionality tests
│   ├── test_config.py       # Configuration file tests
│   ├── test_integration.py  # Integration tests with real Office documents
│   └── test_debugging.py    # New debugging tests with options file output
└── test_cli.py             # Entry point that imports all CLI test modules
```

### 🎯 Key Benefits:

1. **Better Organization**: Tests are now logically grouped by functionality
2. **Debugging Support**: The new `test_debugging.py` includes comprehensive tests that write effective option values to `test_options_{test_name}.txt` files for debugging configuration merging and placeholder resolution
3. **Maintainability**: Each file has a focused responsibility, making it easier to maintain and extend
4. **Scalability**: Easy to add new test categories without cluttering a single file

### 🧪 Test Discovery Verification:

- **28 tests collected** in 1.50s - All tests are being discovered correctly by pytest
- The reorganization maintains all existing functionality while adding new debugging capabilities

### 🔧 Debugging Features Added:

The new `test_debugging.py` includes tests that write detailed debugging information to files:

- **Default option values** without configuration
- **Config file and CLI argument merging** behavior
- **Placeholder resolution** with expected vs actual values
- **Relative path resolution** in different contexts
- **Command line precedence** over config files
- **Real document testing** with actual Office files
- **Multiple configuration scenarios** and their outcomes

### 📖 Usage Examples:

```bash
# Run all CLI tests
pytest tests/cli/ -v

# Run specific test categories
pytest tests/cli/test_basic.py -v
pytest tests/cli/test_config.py -v
pytest tests/cli/test_debugging.py -v

# Run debugging tests that write option files (for your original request)
pytest tests/cli/test_debugging.py -v

# Run with specific Office apps
pytest tests/cli/ --apps excel -v
```

The debugging tests will create `test_options_*.txt` files in the temporary directories that contain detailed information about:

- Effective option values after merging
- Placeholder resolution results
- Configuration file processing
- Command line argument precedence

This should help you identify exactly where issues occur in the configuration merging and placeholder resolution process!
