dim age

age = InputBox("Enter your age: ")

if age < 18 Then
    MsgBox "You are too young."
    WScript.Echo "You are too young."
ElseIf age < 45 Then
    MsgBox "You are still too young."
    WScript.Echo "You are still too young."
ElseIf age < 70 Then
    MsgBox "You are getting older"
    WScript.Echo "You are getting older"
Else
    MsgBox "You are too old."
    WScript.Echo "You are too old."
End if