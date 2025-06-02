 Option Explicit

Dim fso, folderPath, resetPath, logFile, infoFile, computerName, scriptPath

folderPath = "\\192.168.1.3\обменник\Кондауров"
logFile = folderPath & "\log.txt"
infoFile = folderPath & "\информация.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
computerName = CreateObject("WScript.Network").ComputerName
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
resetPath = scriptPath & "\reset"

' При запуске создаём папки и файлы (если нет)
CheckAndCreateFoldersAndFiles

' При запуске копируем содержимое reset в Кондауров
On Error Resume Next
CopyFolderContents resetPath, folderPath
If Err.Number <> 0 Then
    WriteToLog "Ошибка копирования reset -> Кондауров при запуске: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' Основной цикл - каждую минуту проверяем и копируем Кондауров -> reset
Do
    CheckAndCreateFoldersAndFiles
    On Error Resume Next
    CopyFolderContents folderPath, resetPath
    If Err.Number <> 0 Then
        WriteToLog "Ошибка при копировании Кондауров -> reset: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    WScript.Sleep 60000 ' 60 000 мс = 1 минута
Loop


' Процедура проверки и создания папок и файлов
Sub CheckAndCreateFoldersAndFiles()
    On Error Resume Next

    ' Проверяем папку "Кондауров"
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
        WriteToLog "Папка создана: " & folderPath
    End If

    ' Проверяем папку "reset"
    If Not fso.FolderExists(resetPath) Then
        fso.CreateFolder(resetPath)
        WriteToLog "Папка создана: " & resetPath
    End If

    ' Проверяем файл "информация.txt"
    If Not fso.FileExists(infoFile) Then
        Dim infoStream
        Set infoStream = fso.CreateTextFile(infoFile, True)
        infoStream.WriteLine "Данная папка подкреплена скриптом, который сохраняет всё существующее в ней после очищения обменника в понедельник. Скрипт может выдавать ошибки или неправильно работать, они записываются в log файле. Автор: Кондауров"
        infoStream.Close
        WriteToLog "Файл создан: " & infoFile
    End If

    ' Проверяем файл "log.txt"
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


' Процедура копирования содержимого папки source в destination
Sub CopyFolderContents(source, destination)
    Dim srcFolder, destFolder, file, subfolder

    On Error Resume Next
    Set srcFolder = fso.GetFolder(source)
    Set destFolder = fso.GetFolder(destination)

    If Err.Number <> 0 Then
        WriteToLog "Ошибка доступа к папкам: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Удаляем все из папки назначения перед копированием
    On Error Resume Next
    For Each file In destFolder.Files
        fso.DeleteFile file.Path, True
        If Err.Number <> 0 Then
            WriteToLog "Ошибка удаления файла " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next
    For Each subfolder In destFolder.SubFolders
        fso.DeleteFolder subfolder.Path, True
        If Err.Number <> 0 Then
            WriteToLog "Ошибка удаления папки " & subfolder.Path & ": " & Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0

    ' Копируем файлы из исходной папки
    On Error Resume Next
    For Each file In srcFolder.Files
        fso.CopyFile file.Path, destFolder.Path & "\", True
        If Err.Number <> 0 Then
            WriteToLog "Ошибка копирования файла " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next

    ' Копируем подпапки рекурсивно
    For Each subfolder In srcFolder.SubFolders
        CopySubFolder subfolder.Path, destFolder.Path & "\" & subfolder.Name
    Next
    On Error GoTo 0
End Sub

' Рекурсивное копирование подпапки
Sub CopySubFolder(source, destination)
    On Error Resume Next
    If Not fso.FolderExists(destination) Then
        fso.CreateFolder(destination)
        If Err.Number <> 0 Then
            WriteToLog "Ошибка создания папки " & destination & ": " & Err.Description
            Err.Clear
            Exit Sub
        End If
    End If
    On Error GoTo 0

    Dim folder, file, subfolder
    Set folder = fso.GetFolder(source)

    On Error Resume Next
    For Each file In folder.Files
        fso.CopyFile file.Path, destination & "\", True
        If Err.Number <> 0 Then
            WriteToLog "Ошибка копирования файла " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next

    For Each subfolder In folder.SubFolders
        CopySubFolder subfolder.Path, destination & "\" & subfolder.Name
    Next
    On Error GoTo 0
End Sub


' Процедура записи в лог
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
