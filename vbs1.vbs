Const EVENT_LOG_PATH = "Security"
Dim objWMIService, colEvents, objEvent
Dim eventCode, EventUser, Domain, IPAddress

eventCode = "4624"


Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colEvents = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent WHERE Logfile = '" & EVENT_LOG_PATH & "' AND EventCode = '" & eventCode & "'")

If colEvents.Count = 0 Then
    Wscript.Echo "Íåò çàïèñåé î âõîäàõ â æóðíàëå."
Else
    For Each objEvent In colEvents
        EventUser = objEvent.InsertionStrings(5)
        Domain = objEvent.InsertionStrings(6)
        IPAddress = objEvent.InsertionStrings(18) ' IP-àäðåñ

        Wscript.Echo "Âðåìÿ: " & objEvent.TimeGenerated & vbCrLf & _
                     "Èìÿ ïîëüçîâàòåëÿ: " & EventUser & vbCrLf & _
                     "Äîìåí: " & Domain & vbCrLf & _
                     "IP-àäðåñ: " & IPAddress
    Next

End If
