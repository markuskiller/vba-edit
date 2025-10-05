"""Debug script to test ThisWorkbook import detection."""
from pathlib import Path
from vba_edit.office_vba import VBAComponentHandler, VBADocumentNames

# Test the detection logic
handler = VBAComponentHandler()

test_file = Path("DieseArbeitsmappe.cls")

print(f"Testing file: {test_file}")
print(f"Stem (name): {test_file.stem}")
print(f"Is document module: {VBADocumentNames.is_document_module(test_file.stem)}")

module_type = handler.get_module_type(test_file, in_file_headers=True, encoding='cp1252')
print(f"Detected module type: {module_type}")
print("Expected: VBAModuleType.DOCUMENT")
