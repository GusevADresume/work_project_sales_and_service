Attribute VB_Name = "Module7"
Option Explicit
Dim sp, nn, longs As String
Dim dat, datp As Date
Dim x, y As Integer
Dim sumpr, sumrem, sumkvs, sumvsp, sumb As Double


Sub Analysis()
Application.ScreenUpdating = False
Worksheets("общий реестр продаж").Activate
longs = Cells(2, 16).End(xlDown).Row
dat = CDate(Date)
Range(Cells(2, 17), Cells(longs, 17)).Select
With Selection
 .NumberFormat = "#,##0.00"
 .Replace ",", "."
 .Formula = .Formula
End With

For y = 17 To 17
    For x = 2 To longs
sp = Cells(x, y).Offset(0, -3).Value
datp = CDate(Cells(x, y).Offset(0, -13))
If sp / 365 * (dat - datp) <= sp Then
Cells(x, y).Value = sp / 365 * (dat - datp)
Else
Cells(x, y).Value = sp
End If
    
Next
    Next
sumpr = WorksheetFunction.Sum(Range(Cells(2, 17), Cells(longs, 17))) 'Сумма вознаграждений
sumrem = WorksheetFunction.Sum(Range(Cells(2, 18), Cells(longs, 18))) 'сумма ремонтов
sumkvs = WorksheetFunction.Sum(Range(Cells(2, 13), Cells(longs, 13))) 'сумма кв самсунг
sumvsp = WorksheetFunction.Sum(Range(Cells(2, 14), Cells(longs, 14))) 'сумма вознаграждения SP
Worksheets("Анализ").Activate
sumb = WorksheetFunction.Sum(Range(Cells(7, 2), Cells(8, 2)))
Range("B5").Value = sumpr
Range("B6").Value = sumrem
Range("B7").Value = sumkvs
Range("B8").Value = sumvsp
Range("B9").Value = Cells(5, 2) / 120 * 100
Range("B10").Value = sumrem / Cells(9, 2).Value
sumb = Cells(7, 2) + Cells(8, 2)
Range("B11") = Cells(6, 2) / sumb

End Sub

Sub procent()
Range("B11").Value = Cells(5, 2) / 120 * 100

End Sub
