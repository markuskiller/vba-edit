"""Debug script to test header splitting for document modules."""

from pathlib import Path
from vba_edit.office_vba import VBAComponentHandler

handler = VBAComponentHandler()

# Read the actual exported DieseArbeitsmappe.cls file from temp
test_file = Path(
    r"C:\Users\mki\AppData\Local\Temp\pytest-of-mki\pytest-58\test_import_thisworkbook_to_co0\vba\DieseArbeitsmappe.cls"
)

if test_file.exists():
    with open(test_file, "r", encoding="cp1252") as f:
        full_content = f.read()

    print("=" * 80)
    print("FULL CONTENT:")
    print("=" * 80)
    print(full_content[:500])
    print("\n...\n")

    header, code = handler.split_vba_content(full_content)

    print("=" * 80)
    print("HEADER PORTION:")
    print("=" * 80)
    print(header)
    print()

    print("=" * 80)
    print("CODE PORTION (first 500 chars):")
    print("=" * 80)
    print(code[:500])

    print(f"\n\nHeader lines: {len(header.splitlines())}")
    print(f"Code lines: {len(code.splitlines())}")
    print(f"Total lines in file: {len(full_content.splitlines())}")
else:
    print(f"File not found: {test_file}")
    print("Run the test first to create the temp file")
