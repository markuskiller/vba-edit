"""Quick test to verify split_vba_content() fix is active"""
from vba_edit.excel_vba import ExcelVBAHandler

test_content = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "TestClass"
Attribute VB_GlobalNameSpace = False

Sub TestMethod()
    MsgBox "Hello"
End Sub
"""

# Create a handler instance to access the method
handler = ExcelVBAHandler(None, None)
headers, code = handler.split_vba_content(test_content)

print("=== HEADERS ===")
print(headers)
print(f"\nHeader line count: {len(headers.splitlines())}")

print("\n=== CODE ===")
print(code)
print(f"\nCode line count: {len(code.splitlines())}")

print("\n=== VERIFICATION ===")
if "VERSION" in code or "BEGIN" in code or "MultiUse" in code:
    print("❌ FAILED: Headers are still in code!")
else:
    print("✅ SUCCESS: Headers correctly separated from code!")
