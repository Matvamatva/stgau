Option Explicit

Dim fso, folderPath, resetPath, logFile, infoFile, computerName, scriptPath

folderPath = "\\192.168.1.3\обменник\Кондауров"
logFile = folderPath & "\log.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
computerName = CreateObject("WScript.Network").ComputerName
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

CheckAndCreateFoldersAndFiles

Do
    CheckAndCreateFoldersAndFiles
    On Error Resume Next
    If Err.Number <> 0 Then
        WriteToLog "Ошибка при копировании Кондауров -> reset: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    WScript.Sleep 60000 ' 60 000 мс = 1 минута
Loop

Sub CheckAndCreateFoldersAndFiles()
    On Error Resume Next

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
        WriteToLog "Папка создана: " & folderPath
    End If

    If Not fso.FolderExists(resetPath) Then
        fso.CreateFolder(resetPath)
        WriteToLog "Папка создана: " & resetPath
    End If

    If Not fso.FileExists(logFile) Then
        Dim logStream
        Set logStream = fso.CreateTextFile(logFile, True)
        logStream.Close
        WriteToLog "Файл создан: " & logFile
    End If

    If Err.Number <> 0 Then
        WriteToLog "Ошибка при проверке/создании папок или файлов: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub WriteToLog(message)
    On Error Resume Next
    Dim logStream
    If Not fso.FileExists(logFile) Then
        Set logStream = fso.CreateTextFile(logFile, True)
    Else
        Set logStream = fso.OpenTextFile(logFile, 8, True)
    End If

    If Err.Number <> 0 Then
        WriteToLog "Ошибка записи в лог: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    logStream.WriteLine "[" & Now & "] " & message & " // " & computerName
    logStream.Close
    On Error GoTo 0
End Sub
