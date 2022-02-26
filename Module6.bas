Attribute VB_Name = "Module6"
Option Explicit
Dim months, url, urr, tex, nn, monts, dd, enddd As String
Dim longer, y, x, z As Integer
Dim pst As Integer
Function IsBookOpen(wbName As String) As Boolean
    Dim wbBook As Workbook
    For Each wbBook In Workbooks
        If wbBook.Name <> ThisWorkbook.Name Then
            If Windows(wbBook.Name).Visible Then
                If wbBook.Name = wbName Then IsBookOpen = True: Exit For
            End If
        End If
    Next wbBook
End Function
Sub UPLOtest()

months = ThisWorkbook.Worksheets("Команды").Range("R2").Value 'месяц
longer = ThisWorkbook.Worksheets("Общий реестр продаж").Range("A2").End(xlDown).Row ' длинна всего массива
nn = ThisWorkbook.Worksheets("Справочник").Range("BO2").Value 'номер авр
dd = CDate(months) 'дата начала периодав цифрах
enddd = DateSerial(Year(dd), Month(dd) + 1, 0) 'дата конца периода в цифрах

ThisWorkbook.Worksheets("Справочник").Range("BO2").Value = nn + 1
Dim sFldr$
Dim sfldr1$
sFldr = "C:\Users\ivanp\Desktop\Teamplate AVR\авр\" & months & "\"
sfldr1 = "C:\Users\ivanp\Desktop\Teamplate AVR\авр" & months & "\.xlsx"
If Dir(sFldr, vbDirectory) = "" Then
          MkDir sFldr
        Else
 If Dir(months & ".xlsx") <> "" Then
 Kill months & ".xlsx"
End If
End If
ThisWorkbook.Worksheets("Справочник").Activate
For z = 61 To 61
    For urr = 2 To 14
    ThisWorkbook.Worksheets("Справочник").Activate
     url = Cells(urr, z)
     
    ThisWorkbook.Worksheets("Общий реестр продаж").Activate
   

    
    
ThisWorkbook.Worksheets("Общий реестр продаж").Activate
For y = 1 To 1
    For x = 2 To longer
            If Cells(x, y).Offset(0, 14).Value = months And Cells(x, y).Offset(0, 1).Value = url Then
                    On Error Resume Next
                    Workbooks.Open Filename:="C:\Users\ivanp\Desktop\Teamplate AVR\" & url & ".xlsx"
                    On Error GoTo 0
                    Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("G1").Value = ("№ " & nn & " от " & Date)
                    Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("D2").Value = ("c " & dd & " по " & enddd)
                    ThisWorkbook.Worksheets("Общий реестр продаж").Activate
            If Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5") = "" Then
            Cells(x, y).Offset(0, 4).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5") 'номер карты
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("C5") 'Название продукта
            Cells(x, y).Offset(0, 3).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("D5") 'дата продажи
            Cells(x, y).Offset(0, 5).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("E5") ' Тип устройства
            Cells(x, y).Offset(0, 6).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("F5") 'Название устройства
            Cells(x, y).Offset(0, 7).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("G5") ' imei
            Cells(x, y).Offset(0, 8).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("H5") ' цена устройства
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("I5") 'цена договора
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("A6") 'название продукта
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("B6") ' стоимость договора
            Cells(x, y).Offset(0, 12).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("C6") ' сумма ав
            Cells(x, y).Offset(0, 13).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("D6") ' вознаграждение SP
            Else
            Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Activate
    Rows("6:6").Select
    Selection.Insert Shift:=xlDown
    ThisWorkbook.Sheets("Общий реестр продаж").Activate
            Cells(x, y).Offset(0, 4).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 0)
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 1) 'Название продукта
            Cells(x, y).Offset(0, 3).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 2) 'дата продажи
            Cells(x, y).Offset(0, 5).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 3) ' Тип устройства
            Cells(x, y).Offset(0, 6).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 4) 'Название устройства
            Cells(x, y).Offset(0, 7).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 5) ' imei
            Cells(x, y).Offset(0, 8).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 6) ' цена устройства
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Range("B5").Offset(1, 7) 'цена договора
           Workbooks(url & ".xlsx").Worksheets("АВР").Activate
    Rows("7:7").Select
    Selection.Insert Shift:=xlDown
    ThisWorkbook.Sheets("Общий реестр продаж").Activate
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("A6").Offset(1, 0) 'название продукта
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("B6").Offset(1, 0) ' стоимость договора
            Cells(x, y).Offset(0, 12).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("C6").Offset(1, 0) ' сумма ав
            Cells(x, y).Offset(0, 13).Copy Workbooks(url & ".xlsx").Worksheets("АВР").Range("D6").Offset(1, 0) ' вознаграждение SP
         End If
         End If
         
   
       
Next
If IsBookOpen(url & ".xlsx") Then
         Workbooks(url & ".xlsx").Worksheets("Отчет о продажах").Activate
            ActiveWorkbook.SaveAs Filename:=("C:\Users\ivanp\Desktop\Teamplate AVR\авр\" & months & "\" & url & ".xlsx")
                ActiveWorkbook.Close
                 End If
Next
Next
Next

End Sub
Sub Макрос3()
Attribute Макрос3.VB_ProcData.VB_Invoke_Func = " \n14"

ThisWorkbook.Worksheets("Справочник").Activate
Range("BO2").Value = Range("BO2").Value + 1
nn = ThisWorkbook.Worksheets("Справочник").Range("BO2").Value
Range("BO3").Value = ("№ " & nn & " от " & Date)

End Sub

Sub UPLOAD_avr1()
Dim months As String
url = "ООО_Гармония"
months = ("Июнь 2019")
Call creat_path
If Dir("C:\Users\ivanp\Desktop\Teamplate AVR\" & url & ".xlsx") = "" Then
    MsgBox "нет файла"
Else
On Error Resume Next
    Workbooks.Open Filename:="C:\Users\ivanp\Desktop\Teamplate AVR\ООО_Гармония.xlsx"
    On Error GoTo 0
    Worksheets("Отчет о продажах").Activate
End If
    ActiveWorkbook.SaveAs Filename:=("C:\Users\ivanp\Desktop\Teamplate AVR\" & months & "\ООО_Гармония " & months & ".xlsx")
ActiveWorkbook.Close
End Sub

Sub creat_path()
Dim sFldr$
Dim sfldr1$
sFldr = "C:\Users\ivanp\Desktop\Teamplate AVR\авр\" & months & "\"
sfldr1 = "C:\Users\ivanp\Desktop\Teamplate AVR\авр" & months & "\.xlsx"
If Dir(sFldr, vbDirectory) = "" Then
          MkDir sFldr
        Else
 If Dir(months & ".xlsx") <> "" Then
 Kill months & ".xlsx"
End If
End If
End Sub

