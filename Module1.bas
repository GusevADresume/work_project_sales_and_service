Attribute VB_Name = "Module1"
Option Explicit
Dim d As String
Dim x, y, C, V As Integer
Dim W, Q As Integer
Dim shop, datesale As Object
Dim k, l As Integer
Dim g As String
Dim n As Range
Dim f As Variant

Sub data()
'проставление даты заведения убытка
Application.ScreenUpdating = False
d = Range("C2").End(xlDown).Row
    For y = 2 To 2
        For x = 1 To d

            If Cells(x, y).Value = "" Then
                Cells(x, y).Value = Date
            End If
        Next
    Next
'нумерация
    Range("A2") = 1
        Range("A3") = 2
            Range(Cells(2, 1), Cells(3, 1)).Select
                Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(d, 1)), Type:=xlFillDefault
'проставляем дату изменения статуса
    For y = 14 To 14
        For x = 1 To d
            If Cells(x, y).Value = "" Then
                Cells(x, y).Value = Date
            End If
        Next
    Next
'формулы проверки наличия
   'Shop

 Sheets("Claims").Select
d = Range("C2").End(xlDown).Row
For W = 16 To 16
    For Q = 2 To d
        If Cells(Q, W).Value = "" Then
            Sheets("Справочник").Range("AP2").Copy Cells(Q, W)
              End If
Next
        Next
'date
 Sheets("Claims").Select
d = Range("C2").End(xlDown).Row
For W = 17 To 17
    For Q = 2 To d
        If Cells(Q, W).Value = "" Then
            Sheets("Справочник").Range("AQ2").Copy Cells(Q, W)
              End If
Next
        Next

Columns("P:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'формулы франшиз
For C = 19 To 19
    For V = 2 To d
        If Cells(V, C).Value = "" Then
            If Cells(V, C).Offset(0, -2) > Worksheets("Справочник").Cells(3, 67) Then
                Cells(V, C).Value = ""
            Else
                If Cells(V, C).Offset(0, -10) = ("Более 20000 руб") Then
                    Cells(V, C).Value = 3000
                        Else
                            If Cells(V, C).Offset(0, -10).Value = ("Менее 20000 руб") Then
                               Cells(V, C).Value = 1500
        End If
            End If
                End If
                    End If
Next
    Next
'Дата запроса документов от клиента
Cells(2, 13).Activate
Do While ActiveCell.Value <> ""
    If ActiveCell.Value = "От клиента запрошены доп. Документы" Then
      If ActiveCell.Offset(0, 10).Value = "" Then
        ActiveCell.Offset(0, 10).Value = Date
            End If
        ActiveCell.Offset(1, 0).Activate
    Else
        ActiveCell.Offset(1, 0).Activate
    End If
Loop
' Дата Направления в СЦ
Cells(2, 13).Activate
Do While ActiveCell.Value <> ""
    If ActiveCell.Value = "Клиент направлен в СЦ" Then
      If ActiveCell.Offset(0, 11).Value = "" Then
        ActiveCell.Offset(0, 11).Value = Date
        End If
        ActiveCell.Offset(1, 0).Activate
    Else
        ActiveCell.Offset(1, 0).Activate
    End If
Loop
'Дата Получения доков на согласование
Cells(2, 13).Activate
Do While ActiveCell.Value <> ""
    If ActiveCell.Value = "На согласовании диагностики" Then
      If ActiveCell.Offset(0, 12).Value = "" Then
        ActiveCell.Offset(0, 12).Value = Date
        End If
        ActiveCell.Offset(1, 0).Activate
    Else
        ActiveCell.Offset(1, 0).Activate
    End If
Loop
' Направление на ремонт
Cells(2, 13).Activate
Do While ActiveCell.Value <> ""
    If ActiveCell.Value = "Направлено уведомление о ремонте" Then
      If ActiveCell.Offset(0, 13).Value = "" Then
        ActiveCell.Offset(0, 13).Value = Date
        End If
        ActiveCell.Offset(1, 0).Activate
    Else
        ActiveCell.Offset(1, 0).Activate
    End If
Loop
'Проверка номера карты к imei
For k = 10 To 10
    For l = 2 To d
        f = Cells(l, k)
            If Cells(l, k).Offset(0, -6).Interior.Color <> vbRed Then
            If Cells(l, k).Offset(0, -6).Interior.Color <> vbGreen Then
                Set n = Worksheets("Общий реестр продаж").Range("H1:H10000").Find(f)
                On Error GoTo msg
                g = n.Offset(0, -3)
                If Cells(l, k).Offset(0, -6).Value = g Then
                    Cells(l, k).Offset(0, -6).Interior.Color = vbGreen
                Else
                    Cells(l, k).Offset(0, -6).Interior.Color = vbRed
            End If
            End If
            End If
            
            
Next
    Next

MsgBox ("Готово!")
Exit Sub
msg:
MsgBox ("Такой imei отсутствует")
Cells(l, k).Interior.Color = vbRed
Cells(l, k).Offset(0, -6).Interior.Color = vbRed
End Sub


