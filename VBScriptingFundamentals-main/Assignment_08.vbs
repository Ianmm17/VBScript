option explicit

Dim UserInput, i, result

UserInput = InputBox("enter number between 1-100: ")

If UserInput < 1 OR UserInput > 100 Then
    MsgBox "did not enter number between 1-100"
Else
    For i = 1 to UserInput
        result = i Mod 10
        if result = 0 Then
            MsgBox i
        end if
    Next
End if