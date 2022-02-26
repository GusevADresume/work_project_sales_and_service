Attribute VB_Name = "Module3"
Option Explicit
Dim one As String


Sub upload()
Application.ScreenUpdating = False
one = Cells(2, 3).End(xlDown).Row
Range(Cells(1, 1), Cells(one, 27)).Copy
Workbooks.Add
   Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Columns("K:Q").Select
    Application.CutCopyMode = False
    Selection.Delete
Columns("M:R").Select
    Application.CutCopyMode = False
    Selection.Delete
End Sub

