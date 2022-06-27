option explicit

Dim UserInput, num

UserInput = InputBox("Enter number between 1-25")

If UserInput < 1 or UserInput > 25 Then
    MsgBox "Did not enter number between 1-25"
Else
    for num = 1 to UserInput
    MsgBox num
    Next
End If

