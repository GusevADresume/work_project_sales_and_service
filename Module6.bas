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

months = ThisWorkbook.Worksheets("�������").Range("R2").Value '�����
longer = ThisWorkbook.Worksheets("����� ������ ������").Range("A2").End(xlDown).Row ' ������ ����� �������
nn = ThisWorkbook.Worksheets("����������").Range("BO2").Value '����� ���
dd = CDate(months) '���� ������ �������� ������
enddd = DateSerial(Year(dd), Month(dd) + 1, 0) '���� ����� ������� � ������

ThisWorkbook.Worksheets("����������").Range("BO2").Value = nn + 1
Dim sFldr$
Dim sfldr1$
sFldr = "C:\Users\ivanp\Desktop\Teamplate AVR\���\" & months & "\"
sfldr1 = "C:\Users\ivanp\Desktop\Teamplate AVR\���" & months & "\.xlsx"
If Dir(sFldr, vbDirectory) = "" Then
          MkDir sFldr
        Else
 If Dir(months & ".xlsx") <> "" Then
 Kill months & ".xlsx"
End If
End If
ThisWorkbook.Worksheets("����������").Activate
For z = 61 To 61
    For urr = 2 To 14
    ThisWorkbook.Worksheets("����������").Activate
     url = Cells(urr, z)
     
    ThisWorkbook.Worksheets("����� ������ ������").Activate
   

    
    
ThisWorkbook.Worksheets("����� ������ ������").Activate
For y = 1 To 1
    For x = 2 To longer
            If Cells(x, y).Offset(0, 14).Value = months And Cells(x, y).Offset(0, 1).Value = url Then
                    On Error Resume Next
                    Workbooks.Open Filename:="C:\Users\ivanp\Desktop\Teamplate AVR\" & url & ".xlsx"
                    On Error GoTo 0
                    Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("G1").Value = ("� " & nn & " �� " & Date)
                    Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("D2").Value = ("c " & dd & " �� " & enddd)
                    ThisWorkbook.Worksheets("����� ������ ������").Activate
            If Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5") = "" Then
            Cells(x, y).Offset(0, 4).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5") '����� �����
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("C5") '�������� ��������
            Cells(x, y).Offset(0, 3).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("D5") '���� �������
            Cells(x, y).Offset(0, 5).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("E5") ' ��� ����������
            Cells(x, y).Offset(0, 6).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("F5") '�������� ����������
            Cells(x, y).Offset(0, 7).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("G5") ' imei
            Cells(x, y).Offset(0, 8).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("H5") ' ���� ����������
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("I5") '���� ��������
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("A6") '�������� ��������
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("B6") ' ��������� ��������
            Cells(x, y).Offset(0, 12).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("C6") ' ����� ��
            Cells(x, y).Offset(0, 13).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("D6") ' �������������� SP
            Else
            Workbooks(url & ".xlsx").Worksheets("����� � ��������").Activate
    Rows("6:6").Select
    Selection.Insert Shift:=xlDown
    ThisWorkbook.Sheets("����� ������ ������").Activate
            Cells(x, y).Offset(0, 4).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 0)
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 1) '�������� ��������
            Cells(x, y).Offset(0, 3).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 2) '���� �������
            Cells(x, y).Offset(0, 5).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 3) ' ��� ����������
            Cells(x, y).Offset(0, 6).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 4) '�������� ����������
            Cells(x, y).Offset(0, 7).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 5) ' imei
            Cells(x, y).Offset(0, 8).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 6) ' ���� ����������
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("����� � ��������").Range("B5").Offset(1, 7) '���� ��������
           Workbooks(url & ".xlsx").Worksheets("���").Activate
    Rows("7:7").Select
    Selection.Insert Shift:=xlDown
    ThisWorkbook.Sheets("����� ������ ������").Activate
            Cells(x, y).Offset(0, 11).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("A6").Offset(1, 0) '�������� ��������
            Cells(x, y).Offset(0, 9).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("B6").Offset(1, 0) ' ��������� ��������
            Cells(x, y).Offset(0, 12).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("C6").Offset(1, 0) ' ����� ��
            Cells(x, y).Offset(0, 13).Copy Workbooks(url & ".xlsx").Worksheets("���").Range("D6").Offset(1, 0) ' �������������� SP
         End If
         End If
         
   
       
Next
If IsBookOpen(url & ".xlsx") Then
         Workbooks(url & ".xlsx").Worksheets("����� � ��������").Activate
            ActiveWorkbook.SaveAs Filename:=("C:\Users\ivanp\Desktop\Teamplate AVR\���\" & months & "\" & url & ".xlsx")
                ActiveWorkbook.Close
                 End If
Next
Next
Next

End Sub
Sub ������3()
Attribute ������3.VB_ProcData.VB_Invoke_Func = " \n14"

ThisWorkbook.Worksheets("����������").Activate
Range("BO2").Value = Range("BO2").Value + 1
nn = ThisWorkbook.Worksheets("����������").Range("BO2").Value
Range("BO3").Value = ("� " & nn & " �� " & Date)

End Sub

Sub UPLOAD_avr1()
Dim months As String
url = "���_��������"
months = ("���� 2019")
Call creat_path
If Dir("C:\Users\ivanp\Desktop\Teamplate AVR\" & url & ".xlsx") = "" Then
    MsgBox "��� �����"
Else
On Error Resume Next
    Workbooks.Open Filename:="C:\Users\ivanp\Desktop\Teamplate AVR\���_��������.xlsx"
    On Error GoTo 0
    Worksheets("����� � ��������").Activate
End If
    ActiveWorkbook.SaveAs Filename:=("C:\Users\ivanp\Desktop\Teamplate AVR\" & months & "\���_�������� " & months & ".xlsx")
ActiveWorkbook.Close
End Sub

Sub creat_path()
Dim sFldr$
Dim sfldr1$
sFldr = "C:\Users\ivanp\Desktop\Teamplate AVR\���\" & months & "\"
sfldr1 = "C:\Users\ivanp\Desktop\Teamplate AVR\���" & months & "\.xlsx"
If Dir(sFldr, vbDirectory) = "" Then
          MkDir sFldr
        Else
 If Dir(months & ".xlsx") <> "" Then
 Kill months & ".xlsx"
End If
End If
End Sub

