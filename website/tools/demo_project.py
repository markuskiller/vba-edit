"""Build the sample project used for the terminal recordings and screenshots of the docs.

Creates `Budget.xlsm` in a NEW directory with a hidden Excel instance and closes that instance again.
The workbook gets English code names (ThisWorkbook, Sheet1) on any Office language.

    python demo_project.py <new dir>
"""

import sys
from pathlib import Path

import pythoncom
import win32com.client

REPORTING = """\'@Folder("Reports")
Option Explicit

Public Sub BuildMonthlyReport()
    Dim total As Currency
    total = SumColumn(Sheet1.Range("B2:B13"))
    Sheet1.Range("B14").Value = total
    MsgBox "Total: " & FormatMoney(total)
End Sub
"""

HELPERS = """\'@Folder("Shared")
Option Explicit

Public Function SumColumn(ByVal cells As Range) As Currency
    Dim cell As Range
    For Each cell In cells
        If IsNumeric(cell.Value) Then SumColumn = SumColumn + cell.Value
    Next cell
End Function

Public Function FormatMoney(ByVal money As Currency) As String
    FormatMoney = Format$(money, "#,##0.00")
End Function
"""

INVOICE = """\'@Folder("Domain")
Option Explicit

Private mAmount As Currency
Public Customer As String

Public Property Get Amount() As Currency
    Amount = mAmount
End Property

Public Property Let Amount(ByVal newAmount As Currency)
    If newAmount < 0 Then Err.Raise 5, , "Amount must not be negative"
    mAmount = newAmount
End Property

Public Function IsOverdue(ByVal dueDate As Date) As Boolean
    IsOverdue = (Date > dueDate)
End Function
"""

THIS_WORKBOOK = """Option Explicit

Private Sub Workbook_Open()
    Application.StatusBar = "Budget tools ready"
End Sub
"""

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
AMOUNTS = [4120, 3980, 4410, 4275, 4630, 4890, 5010, 4760, 4540, 4820, 5130, 5390]


def set_code(component, text):
    module = component.CodeModule
    if module.CountOfLines:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(text.rstrip("\n"))


def build(directory: Path, *, with_form: bool = False) -> Path:
    directory = directory.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / "Budget.xlsm"
    if path.exists():
        raise ValueError(f"{path} exists; use a new directory")
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        book = excel.Workbooks.Add()
        while book.Worksheets.Count > 1:
            book.Worksheets(book.Worksheets.Count).Delete()
        sheet = book.Worksheets(1)
        components = book.VBProject.VBComponents
        # Document modules are renamed through _CodeName; setting .Name leaves a stray module behind.
        components(book.CodeName).Properties("_CodeName").Value = "ThisWorkbook"
        components(sheet.CodeName).Properties("_CodeName").Value = "Sheet1"
        sheet.Name = "Budget"
        sheet.Range("A1").Value = "Month"
        sheet.Range("B1").Value = "Spend"
        for row, (month, amount) in enumerate(zip(MONTHS, AMOUNTS), start=2):
            sheet.Cells(row, 1).Value = month
            sheet.Cells(row, 2).Value = amount
        sheet.Range("A14").Value = "Total"
        for name, kind, code in (("Reporting", 1, REPORTING), ("Helpers", 1, HELPERS), ("Invoice", 2, INVOICE)):
            component = components.Add(kind)
            component.Name = name
            set_code(component, code)
        set_code(components("ThisWorkbook"), THIS_WORKBOOK)
        # Microsoft Scripting Runtime, so the references commands have something to show.
        book.VBProject.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
        if with_form:
            form = components.Add(3)
            form.Name = "ExportDialog"
            form.Properties("Caption").Value = "Export report"
            set_code(
                form, 'Option Explicit\n\nPrivate Sub UserForm_Initialize()\n    Me.Caption = "Export report"\nEnd Sub'
            )
        book.SaveAs(str(path), FileFormat=52)
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    return path


if __name__ == "__main__":
    print(build(Path(sys.argv[1]), with_form="--with-form" in sys.argv))
