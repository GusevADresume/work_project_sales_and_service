Attribute VB_Name = "Module2"
Option Explicit
Sub Mail_Service_Definition()
Application.ScreenUpdating = False
Dim x, y As String
Dim SC, NSC As Object
Sheets("Claims").Select
x = Range("AF1")
y = Cells(x, 12).Value
Sheets("����������").Select
Set SC = Columns("E:E").Find(y).Offset(0, 1)
If Not SC Is Nothing Then
Range("I2") = SC
End If
Set NSC = Columns("E:E").Find(y).Offset(0, -1)
If Not NSC Is Nothing Then
Range("L2") = NSC
End If
Sheets("Claims").Activate
End Sub

Sub Addres_Service_Definition()
Application.ScreenUpdating = False
Dim x, y As String
Dim ASC, NSC As Object
Sheets("Claims").Select
x = Range("AF1")
y = Cells(x, 12).Value
Sheets("����������").Select
Set ASC = Columns("E:E").Find(y).Offset(0, 2)
If Not ASC Is Nothing Then
Range("J2") = ASC
End If
Set NSC = Columns("E:E").Find(y).Offset(0, -1)
If Not NSC Is Nothing Then
Range("L2") = NSC
End If
Sheets("Claims").Activate
End Sub

Sub mail_to_service()
Application.ScreenUpdating = False
Dim x, y As String

x = Range("AF1")
y = Cells(x, 3) & " " & Cells(x, 7) & " " & Cells(x, 8) & " " & Cells(x, 10)
 Call Mail_Service_Definition
 Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("����������").Range("i2").Value  '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = y + Worksheets("����������").Range("L2").Value   '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = Worksheets("����������").Range("AH2").Value    '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub

Sub mail_to_client()
Application.ScreenUpdating = False
Dim x, y As String
x = Range("AF1")
y = Cells(x, 28)
Call Addres_Service_Definition
Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = y  '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = ("����������� �� �����������")    '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = Worksheets("����������").Range("K2").Value + Worksheets("����������").Range("J2").Value + (", ��� ���� ���������� ����� ��������� ����� � �������")   '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub

Sub ������������������()
Application.ScreenUpdating = False
Sheets("Claims").Activate
Dim x, y As String
x = Cells(1, 32)
y = Cells(x, 13)
If y = "���������� ����������� � �������" Then
Call mail_if_accept
Else: Call mail_if_fail
End If

End Sub


Sub mail_if_accept()
Application.ScreenUpdating = False
Dim x, y As String

x = Range("AF1")
y = Cells(x, 3) & " " & Cells(x, 7) & " " & Cells(x, 8) & " " & Cells(x, 10)
Call Mail_Service_Definition
 Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("����������").Range("i2").Value  '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = y    '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = ("�������, ������ ����������, ������� " & Cells(x, 19) & " ������") '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub


Sub mail_if_fail()
Application.ScreenUpdating = False
Dim x, y As String

x = Range("AF1")
y = Cells(x, 3) & " " & Cells(x, 7) & " " & Cells(x, 8) & " " & Cells(x, 10)
Call Mail_Service_Definition
 Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("����������").Range("i2").Value  '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = y    '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = ("�������, ������ �� ����������, ������ ������ ���������� ��� �������") '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub

Sub inquiryservice()
Application.ScreenUpdating = False
Sheets("Claims").Activate
Dim x, y As String
x = Cells(1, 32)
y = Cells(x, 13)
If y = "������ �� � �������" Then
Call mail_if_inquiry
Else: Call mail_to_service
End If

End Sub
Sub mail_if_inquiry()
Application.ScreenUpdating = False
Dim x, y As String

x = Range("AF1")
y = Cells(x, 3) & " " & Cells(x, 7) & " " & Cells(x, 8) & " " & Cells(x, 10) & " " & Cells(x, 11)

 Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = "insurance@aide.ru;ikulagina@aide.ru" '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = y    '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = ("�������, ������ ����, ������ ��������� �� � ������ " & Cells(x, 11) & " .") '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub

Sub variable_to_client()
Application.ScreenUpdating = False
Sheets("Claims").Activate
Dim x, y As String
x = Cells(1, 32)
y = Cells(x, 13)
If y = "������ �� � �������" Then
Call mail_wait_to_client
Else: Call mail_to_client
End If

End Sub
Sub mail_wait_to_client()
Application.ScreenUpdating = False
Dim x, y As String
x = Range("AF1")
y = Cells(x, 28)
Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    '������� ������������ � Outlook ���� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ������� ������� ���������� ��� ��������� ��������� �������
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = y  '�����, ����� �������� sTo= Range("A1").Value)
    sSubject = ("����������� �� ������")    '����, ����� �������� - sSubject = Range("A2").Value)
    sBody = ("��������� ������! ����� ������ �� ������ �������, �� �������� ��������� ����� � ����� ������� � ����� ��� �������� �����!")    '����� ����� �������� - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    '��������, ����� ������� ���� � ����� - sAttachment = Range("A4").Value)
 
    '������� ���������
    With objMail
        .To = sTo '�����
        .CC = "" '�����
        .BCC = "" '�������
        .Subject = sSubject '����
        .Body = sBody '�����
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment '������ ��������
            '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
        'End If
        .Display 'Display, ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub
