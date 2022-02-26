Attribute VB_Name = "Module2"
Option Explicit
Sub Mail_Service_Definition()
Application.ScreenUpdating = False
Dim x, y As String
Dim SC, NSC As Object
Sheets("Claims").Select
x = Range("AF1")
y = Cells(x, 12).Value
Sheets("Справочник").Select
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
Sheets("Справочник").Select
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("Справочник").Range("i2").Value  'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = y + Worksheets("Справочник").Range("L2").Value   'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = Worksheets("Справочник").Range("AH2").Value    'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = y  'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = ("Направление на диагностику")    'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = Worksheets("Справочник").Range("K2").Value + Worksheets("Справочник").Range("J2").Value + (", При себе необходимо иметь сервисную карту и паспорт")   'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub

Sub определитьфраншизу()
Application.ScreenUpdating = False
Sheets("Claims").Activate
Dim x, y As String
x = Cells(1, 32)
y = Cells(x, 13)
If y = "Направлено уведомление о ремонте" Then
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("Справочник").Range("i2").Value  'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = y    'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = ("Коллеги, ремонт согласован, доплата " & Cells(x, 19) & " рублей") 'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = Worksheets("Справочник").Range("i2").Value  'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = y    'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = ("Коллеги, ремонт не согласован, просим выдать устройство без ремонта") 'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
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
If y = "Запрос СЦ в регионе" Then
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = "insurance@aide.ru;ikulagina@aide.ru" 'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = y    'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = ("Коллеги, добрый день, просим подобрать сц в городе " & Cells(x, 11) & " .") 'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
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
If y = "Запрос СЦ в регионе" Then
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
    'Пробуем подключиться к Outlook усли уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    err.Clear 'Outlook Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    objOutlookApp.Session.Logon
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не удалось создать приложение или экземпляр сообщения выходим
    If err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = y  'адрес, можно заменить sTo= Range("A1").Value)
    sSubject = ("Направление на ремонт")    'Тема, можно заменить - sSubject = Range("A2").Value)
    sBody = ("Уважаемый клиент! Вашая заявка на ремонт принята, мы подберем сервисный центр в вашем регионе и дадим вам обратную связь!")    'Текст можно заменить - sBody = Range("A3").Value)
    'sAttachment = "C:\Temp\?????1.xls"    'вложение, можно указать путь к файлу - sAttachment = Range("A4").Value)
 
    'создаем сообщение
    With objMail
        .To = sTo 'адрес
        .CC = "" 'копия
        .BCC = "" 'скрытая
        .Subject = sSubject 'тема
        .Body = sBody 'текст
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        'If Dir(sAttachment, 16) <> "" Then
            '.Attachments.Add sAttachment 'просто вложение
            'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
        'End If
        .Display 'Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub
