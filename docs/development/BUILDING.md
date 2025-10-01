create_binaries.py provides the following features:

## Key Features:

1. **Selective Building**: Use `--apps` to specify which executables to build:
   ```bash
   python create_binaries.py --apps excel          # Build only excel-vba.exe
   python create_binaries.py --apps word excel     # Build word-vba.exe and excel-vba.exe
   ```

2. **List Available Apps**: Use `--list` to see all available applications:
   ```bash
   python create_binaries.py --list
   ```

3. **Debug Mode**: Use `--debug` for faster builds with debug information:
   ```bash
   python create_binaries.py --debug --apps excel
   ```

4. **Custom Output Directory**: Specify where to place the built executables:
   ```bash
   python create_binaries.py --output-dir ./my-builds --apps excel
   ```

5. **Error Handling**: The script now handles missing files and build failures gracefully.

6. **Build Summary**: Shows which builds succeeded and which failed.

## Usage Examples:

```bash
# Build all executables (default behavior)
python create_binaries.py

# Build only Excel VBA executable
python create_binaries.py --apps excel

# Build Word and Excel executables
python create_binaries.py --apps word excel

# List available applications
python create_binaries.py --list

# Build with debug information
python create_binaries.py --apps excel --debug

# Build to custom directory
python create_binaries.py --apps excel --output-dir ./release
```

The script maintains backward compatibility - running it without arguments will build all executables just like the original version. The configuration is centralized in the `create_build_config()` function, making it easy to add new applications or modify build settings in the future.