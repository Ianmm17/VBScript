option explicit

Const SITE_TITLE = "ian.miller"

Function Add(num1, num2)
    Add = num1 + num2
End Function

Function Subtract(num1, num2)
    Subtract = num1 - num2
End Function

Function Multiply(num1, num2)
    Multiply = num1 * num2
End Function

Function Division(num1, num2)
    Division = num1 / num2
End Function

Function PowerOf(num1, num2)
    PowerOf = num1 ^ num2
End Function

Function DisplayMsg(stringMSG, result)
    MsgBox stringMSG & " : " & result, 0, SITE_TITLE
End Function

Dim input1, input2, result, UserChoice

input1 = CInt(InputBox("enter first number: "))
input2 = CInt(InputBox("enter second number: "))

UserChoice = InputBox("Enter: " & vbNewLine & " 1 for Add" & vbNewLine & " 2 for Subtract" & vbNewLine & " 3 for Multiply" & vbNewLine & " 4 for Divide" & vbNewLine & " 5 for PowerOf")

Select Case UserChoice
    Case 1
        result = Add(input1, input2)
    Case 2
        result = Subtract(input1, input2)
    Case 3
        result = Multiply(input1, input2)
    Case 4
        result = Division(input1, input2)
    Case 5
        result = PowerOf(input1, input2)
    Case else
        MsgBox "Your selection is not VALID", 48, "ERROR"
End Select

If UserChoice > 0 AND UserChoice < 6 Then
    DisplayMsg "The result is ", result
End If