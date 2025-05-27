Dim username, password
Dim objNetwork

Set objNetwork = CreateObject("WScript.Network")

username = InputBox("Введите имя пользователя:")
password = InputBox("Введите пароль:")

' Попытка подключения к ресурсу с использованием введенного пользователя и пароля
On Error Resume Next
objNetwork.MapNetworkDrive "Z:", "\\192.168.1.3\обменник", username, password

If Err.Number = 0 Then
    MsgBox "Аутентификация прошла успешно!"
    ' Удаление сетевого привода
    ' objNetwork.RemoveNetworkDrive "Z:"
Else
    MsgBox "Аутентификация не прошла! Ошибка: " & Err.Description
End If

On Error GoTo 0