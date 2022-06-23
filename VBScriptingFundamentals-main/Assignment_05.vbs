option explicit

Dim John_Age, Kerry_Age

John_Age = Cint(InputBox("Enter John's age: "))
Kerry_Age = Cint(InputBox("Enter Kerry's age: "))

do while John_Age < 1 or Kerry_Age < 1
    John_Age = Cint(InputBox("Age has to be greater than 0, Enter John's age: "))
    Kerry_Age = Cint(InputBox("Age has to be greater than 0, Enter Kerry's age: "))
    if John_Age > 1 and Kerry_Age > 1 Then Exit do
loop

If John_Age > Kerry_Age Then
    MsgBox "John is older"
ElseIf John_Age < Kerry_Age Then
    MsgBox "Kerry is older"
End If