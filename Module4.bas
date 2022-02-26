Attribute VB_Name = "Module4"
Option Explicit
Dim y, x As Integer


Public cp, od, Line, pc, line1, newname As String

Public Sub dims()
Application.ScreenUpdating = False
cp = Range("b2").End(xlDown).Offset(, 9).Row
od = Range("A1").End(xlToRight).Column
Line = Range("B1").End(xlDown).Offset(, 12).Row
line1 = Range("b1").End(xlDown).Row
pc = Range("B1").End(xlDown).Offset(, -1).Row
newname = Sheets("Команды").Range("R2")
End Sub




Sub форматирование()

Sheets("Общий реестр продаж").Activate
Call dims
    Range(Cells(1, 1), Cells(line1, od)).RemoveDuplicates Array(5, 8), xlYes
    Range("A2") = 1
        Range("A3") = 2
            Range(Cells(2, 1), Cells(3, 1)).Select
                Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(pc, 1)), Type:=xlFillDefault
    Range(Cells(2, 7), Cells(line1, 7)).Select
        Selection.Replace What:="Смартфон", Replacement:=""
            Selection.Replace What:="Планшет", Replacement:=""
                Selection.Replace What:="Часы", Replacement:=""
                    Selection.Replace What:="Телевизор", Replacement:=""
    Range(Cells(2, 9), Cells(line1, 10)).Select
        Selection.Replace What:=" ", Replacement:=""
    Range(Cells(2, 4), Cells(Line, 4)).Select
        With Selection
            Selection.Replace What:=",", Replacement:="."
            Selection.NumberFormat = "dd/mm/yyyy"
        End With
    Range(Cells(2, 9), Cells(Line, 10)).Select
        With Selection
            .Replace What:=" ", Replacement:=""
            .NumberFormat = "0.00"
        End With
    Range(Cells(1, 13), Cells(line1, 14)).NumberFormat = "0.00"
    Range(Cells(2, 5), Cells(line1, 5)).Replace What:="№", Replacement:=""
    For y = 6 To 6
    For x = 2 To line1
           If Cells(x, y).Value = ("часы") Then
                If Cells(x, y).Offset(0, 6).Value = "" Then
                    Cells(x, y).Offset(0, 6).Value = ("«Simple Protect» для телевизоров и смарт-часов")
                End If
                End If
    Next
Next
For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("телевизор") Then
            If Cells(x, y).Offset(0, 6).Value = "" Then
                Cells(x, y).Offset(0, 6).Value = ("«Simple Protect» для телевизоров и смарт-часов")
                
                
                End If
                End If
    Next
Next

 For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("смартфон") Then
            If Cells(x, y).Offset(0, 3).Value < 15000 Then
                    If Cells(x, y).Offset(0, 6).Value = "" Then
                        Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 0 до 15000 рублей")
                End If
                End If
                End If
    Next
Next
For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("смартфон") Then
            If Cells(x, y).Offset(0, 3).Value > 15001 < 35000 Then
                        If Cells(x, y).Offset(0, 6).Value = "" Then
                            Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 15001 до 35000 рублей")
                End If
                End If
                End If
                                    
    Next
Next
 For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("смартфон") Then
                If Cells(x, y).Offset(0, 3).Value > 35000 Then
                    If Cells(x, y).Offset(0, 6).Value = "" Then
                        Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов» и планшетов стоимостью от 35001 до 150000 рублей")
                End If
                End If
                End If
    Next
Next

For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("планшет") Then
            If Cells(x, y).Offset(0, 3).Value < 15000 Then
                    If Cells(x, y).Offset(0, 6).Value = "" Then
                        Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 0 до 15000 рублей")
                End If
                End If
                End If
    Next
Next
For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("планшет") Then
            If Cells(x, y).Offset(0, 3).Value > 15001 < 35000 Then
                        If Cells(x, y).Offset(0, 6).Value = "" Then
                            Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 15001 до 35000 рублей")
                End If
                End If
                End If
                                    
    Next
Next
 For y = 6 To 6
    For x = 2 To line1
        If Cells(x, y).Value = ("планшет") Then
                If Cells(x, y).Offset(0, 3).Value > 35000 Then
                    If Cells(x, y).Offset(0, 6).Value = "" Then
                        Cells(x, y).Offset(0, 6).Value = ("«SimpleProtect» для телефонов» и планшетов стоимостью от 35001 до 150000 рублей")
                End If
                End If
                End If
    Next
Next




For y = 12 To 12
    For x = 2 To line1
        If Cells(x, y).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 0 до 15000 рублей") Then
                    If Cells(x, y).Offset(0, -1).Value = "" Then
                        Cells(x, y).Offset(0, -1).Value = ("11,0%")
                End If
                End If
    Next
Next
For y = 12 To 12
    For x = 2 To line1
        If Cells(x, y).Value = ("«SimpleProtect» для телефонов и планшетов стоимостью от 15001 до 35000 рублей") Then
            If Cells(x, y).Offset(0, -1).Value = "" Then
                Cells(x, y).Offset(0, -1).Value = ("7,50%")
                End If
                End If
                                    
    Next
Next
 For y = 12 To 12
    For x = 2 To line1
        If Cells(x, y).Value = ("«SimpleProtect» для телефонов» и планшетов стоимостью от 35001 до 150000 рублей") Then
                    If Cells(x, y).Offset(0, -1).Value = "" Then
                        Cells(x, y).Offset(0, -1).Value = ("6,50%")
                End If
                End If
    Next
Next
For y = 12 To 12
    For x = 2 To line1
        If Cells(x, y).Value = ("«Simple Protect» для телевизоров и смарт-часов") Then
                    If Cells(x, y).Offset(0, -1).Value = "" Then
                        Cells(x, y).Offset(0, -1).Value = ("4,50%")
                End If
                End If
    Next
Next
Range(Cells(2, 11), Cells(cp, 11)).Select
With Selection
 .NumberFormat = "0,00%"
 .Replace ".", ","
 .Formula = .Formula
End With

For y = 14 To 14
    For x = 2 To line1
        If Cells(x, y).Value = "" Then
            Cells(x, y).FormulaR1C1 = "=RC[-5]*RC[-3]"
                End If
   Next
Next

For y = 13 To 13
    For x = 2 To line1
        If Cells(x, y).Value = "" Then
            Cells(x, y).FormulaR1C1 = "=RC[-3]-RC[1]"
                End If
   Next
Next

Range(Cells(2, 15), Cells(2, 16)).Select
 Selection.AutoFill Destination:=Range(Cells(2, 15), Cells(line1, 16)), Type:=xlFillDefault
        Range(Cells(2, 13), Cells(line1, 14)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(Cells(1, 1), Cells(Line, od)).Borders.LineStyle = True
    Range(Cells(1, 1), Cells(Line, od)).Font.Size = 13
        Range(Cells(1, 1), Cells(Line, od)).Font.Name = "timesnewroman"
    Range(Cells(2, 9), Cells(Line, 8)).NumberFormat = "0"
    Range(Cells(1, 1), Cells(line1, od)).EntireColumn.AutoFit
    Columns("A:N").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .ReadingOrder = xlContext
        End With
    Range(Cells(1, 1), Cells(1, od)).Select
        With Selection
            .Font.Bold = True
            .Interior.Color = RGB(222, 200, 34)
        End With
MsgBox ("форматирование завершено")

End Sub
Public Sub выгрузка_месяца()

Call dims
Worksheets("Общий реестр продаж").Activate
            Range(Cells(1, 1), Cells(line1, od)).AutoFilter Field:=15, Criteria1:= _
                newname
    Range(Cells(1, 1), Cells(line1, 14)).Copy
    Workbooks.Add
   Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    With Selection
        .Borders.LineStyle = True
    End With
    Range(Cells(1, 1), Cells(1, 14)).Font.Bold = True
    Range(Cells(1, 1), Cells(1, 14)).Interior.Color = RGB(222, 200, 34)
    Range("N:N").NumberFormat = "0"
    Range("K:K").NumberFormat = "0.00%"
    ActiveWorkbook.SaveAs Filename:="Z:\Simple Protect\Проект Самсунг (1galaxy.ru)\Реестры Акты\Реестры\2020\" & newname & ".xls" 'Z:\Simple Protect\Проект Самсунг (1galaxy.ru)\Реестры Акты\Реестры\2020
    ActiveWorkbook.Close
    Workbooks("Ремонты+продажи.xlsm").Activate
    ActiveSheet.ShowAllData

End Sub
