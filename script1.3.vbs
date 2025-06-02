 Option Explicit

Dim fso, folderPath, resetPath, logFile, infoFile, computerName, scriptPath

folderPath = "\\192.168.1.3\îáìåííèê\123123фывфывфывфв"
logFile = folderPath & "\log.txt"
infoFile = folderPath & "\èíôîðìàöèÿ.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
computerName = CreateObject("WScript.Network").ComputerName
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
resetPath = scriptPath & "\reset"

' Ïðè çàïóñêå ñîçäà¸ì ïàïêè è ôàéëû (åñëè íåò)
CheckAndCreateFoldersAndFiles

' Ïðè çàïóñêå êîïèðóåì ñîäåðæèìîå reset â Êîíäàóðîâ
On Error Resume Next
CopyFolderContents resetPath, folderPath
If Err.Number <> 0 Then
    WriteToLog "Îøèáêà êîïèðîâàíèÿ reset -> Êîíäàóðîâ ïðè çàïóñêå: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' Îñíîâíîé öèêë - êàæäóþ ìèíóòó ïðîâåðÿåì è êîïèðóåì Êîíäàóðîâ -> reset
Do
    CheckAndCreateFoldersAndFiles
    On Error Resume Next
    CopyFolderContents folderPath, resetPath
    If Err.Number <> 0 Then
        WriteToLog "Îøèáêà ïðè êîïèðîâàíèè Êîíäàóðîâ -> reset: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    WScript.Sleep 60000 ' 60 000 ìñ = 1 ìèíóòà
Loop


' Ïðîöåäóðà ïðîâåðêè è ñîçäàíèÿ ïàïîê è ôàéëîâ
Sub CheckAndCreateFoldersAndFiles()
    On Error Resume Next

    ' Ïðîâåðÿåì ïàïêó "Êîíäàóðîâ"
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
        WriteToLog "Ïàïêà ñîçäàíà: " & folderPath
    End If

    ' Ïðîâåðÿåì ïàïêó "reset"
    If Not fso.FolderExists(resetPath) Then
        fso.CreateFolder(resetPath)
        WriteToLog "Ïàïêà ñîçäàíà: " & resetPath
    End If

    ' Ïðîâåðÿåì ôàéë "èíôîðìàöèÿ.txt"
    If Not fso.FileExists(infoFile) Then
        Dim infoStream
        Set infoStream = fso.CreateTextFile(infoFile, True)
        infoStream.WriteLine "Äàííàÿ ïàïêà ïîäêðåïëåíà ñêðèïòîì, êîòîðûé ñîõðàíÿåò âñ¸ ñóùåñòâóþùåå â íåé ïîñëå î÷èùåíèÿ îáìåííèêà â ïîíåäåëüíèê. Ñêðèïò ìîæåò âûäàâàòü îøèáêè èëè íåïðàâèëüíî ðàáîòàòü, îíè çàïèñûâàþòñÿ â log ôàéëå. Àâòîð: Êîíäàóðîâ"
        infoStream.Close
        WriteToLog "Ôàéë ñîçäàí: " & infoFile
    End If

    ' Ïðîâåðÿåì ôàéë "log.txt"
    If Not fso.FileExists(logFile) Then
        Dim logStream
        Set logStream = fso.CreateTextFile(logFile, True)
        logStream.Close
        WriteToLog "Ôàéë ñîçäàí: " & logFile
    End If

    If Err.Number <> 0 Then
        WriteToLog "Îøèáêà ïðè ïðîâåðêå/ñîçäàíèè ïàïîê èëè ôàéëîâ: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


' Ïðîöåäóðà êîïèðîâàíèÿ ñîäåðæèìîãî ïàïêè source â destination
Sub CopyFolderContents(source, destination)
    Dim srcFolder, destFolder, file, subfolder

    On Error Resume Next
    Set srcFolder = fso.GetFolder(source)
    Set destFolder = fso.GetFolder(destination)

    If Err.Number <> 0 Then
        WriteToLog "Îøèáêà äîñòóïà ê ïàïêàì: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Óäàëÿåì âñå èç ïàïêè íàçíà÷åíèÿ ïåðåä êîïèðîâàíèåì
    On Error Resume Next
    For Each file In destFolder.Files
        fso.DeleteFile file.Path, True
        If Err.Number <> 0 Then
            WriteToLog "Îøèáêà óäàëåíèÿ ôàéëà " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next
    For Each subfolder In destFolder.SubFolders
        fso.DeleteFolder subfolder.Path, True
        If Err.Number <> 0 Then
            WriteToLog "Îøèáêà óäàëåíèÿ ïàïêè " & subfolder.Path & ": " & Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0

    ' Êîïèðóåì ôàéëû èç èñõîäíîé ïàïêè
    On Error Resume Next
    For Each file In srcFolder.Files
        fso.CopyFile file.Path, destFolder.Path & "\", True
        If Err.Number <> 0 Then
            WriteToLog "Îøèáêà êîïèðîâàíèÿ ôàéëà " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next

    ' Êîïèðóåì ïîäïàïêè ðåêóðñèâíî
    For Each subfolder In srcFolder.SubFolders
        CopySubFolder subfolder.Path, destFolder.Path & "\" & subfolder.Name
    Next
    On Error GoTo 0
End Sub

' Ðåêóðñèâíîå êîïèðîâàíèå ïîäïàïêè
Sub CopySubFolder(source, destination)
    On Error Resume Next
    If Not fso.FolderExists(destination) Then
        fso.CreateFolder(destination)
        If Err.Number <> 0 Then
            WriteToLog "Îøèáêà ñîçäàíèÿ ïàïêè " & destination & ": " & Err.Description
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
            WriteToLog "Îøèáêà êîïèðîâàíèÿ ôàéëà " & file.Path & ": " & Err.Description
            Err.Clear
        End If
    Next

    For Each subfolder In folder.SubFolders
        CopySubFolder subfolder.Path, destination & "\" & subfolder.Name
    Next
    On Error GoTo 0
End Sub


' Ïðîöåäóðà çàïèñè â ëîã
Sub WriteToLog(message)
    On Error Resume Next
    Dim logStream
    If Not fso.FileExists(logFile) Then
        Set logStream = fso.CreateTextFile(logFile, True)
    Else
        Set logStream = fso.OpenTextFile(logFile, 8, True)
    End If

    If Err.Number <> 0 Then
        WriteToLog "Îøèáêà çàïèñè â ëîã: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    logStream.WriteLine "[" & Now & "] " & message & " // " & computerName
    logStream.Close
    On Error GoTo 0
End Sub
