Dim username, password
Dim objNetwork

Set objNetwork = CreateObject("WScript.Network")

username = InputBox("Введите имя пользователя:")
password = InputBox("Введите пароль:")

On Error Resume Next
objNetwork.MapNetworkDrive "Z:", "\\192.168.1.3\обменник", username, password

If Err.Number = 0 Then
    MsgBox "Аутентификация прошла успешно!"
Else
    MsgBox "Аутентификация не прошла! Ошибка: " & Err.Description
End If


On Error GoTo 0
