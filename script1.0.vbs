Option Explicit

Dim fso, folderPath, logFile, logStream, computerName, scriptPath

folderPath = "\\192.168.1.3\обменник\Кондауров"
logFile = folderPath & "\log.txt"

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
computerName = CreateObject("WScript.Network").ComputerName
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder(folderPath)
    fso.CreateFolder(scriptPath & "\" & "reset")
    WriteToLog "Папка создана: " & folderPath
Else
    WriteToLog "Папка уже существует: " & folderPath
End If

If Err.Number <> 0 Then
    WriteToLog "Ошибка: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

Set fso = Nothing



Sub WriteToLog(message)
On Error Resume Next
    If Not fso.FileExists(logFile) Then
        Set logStream = fso.CreateTextFile(logFile, True)
    Else
        Set logStream = fso.OpenTextFile(logFile, 8, True)
    End If
If Err.Number <> 0 Then
    WriteToLog "Ошибка: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
    logStream.WriteLine "[" & Now & "]" & " " & message & "              // " & computerName
    logStream.Close
End Sub