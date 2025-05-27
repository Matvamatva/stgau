Const EVENT_LOG_PATH = "Security"
Dim objWMIService, colEvents, objEvent
Dim eventCode, EventUser, Domain, IPAddress

eventCode = "4624" ' Код события для успешного входа

' Подключение к WMI
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colEvents = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent WHERE Logfile = '" & EVENT_LOG_PATH & "' AND EventCode = '" & eventCode & "'")

If colEvents.Count = 0 Then
    Wscript.Echo "Нет записей о входах в журнале."
Else
    For Each objEvent In colEvents
        EventUser = objEvent.InsertionStrings(5) ' Имя пользователя
        Domain = objEvent.InsertionStrings(6) ' Домен
        IPAddress = objEvent.InsertionStrings(18) ' IP-адрес

        Wscript.Echo "Время: " & objEvent.TimeGenerated & vbCrLf & _
                     "Имя пользователя: " & EventUser & vbCrLf & _
                     "Домен: " & Domain & vbCrLf & _
                     "IP-адрес: " & IPAddress
    Next
End If