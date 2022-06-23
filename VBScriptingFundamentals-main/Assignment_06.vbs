option explicit

Dim Last_name

Last_name = InputBox("Please enter your last name: ")

select case Last_name

    case "Thompson"
        MsgBox Last_name & " position is President"
    case "Rooter"
        MsgBox Last_name & " position is Sr. Vice President"
    case "Cooper"
        MsgBox Last_name & " position is Vice President"
    case "Parker"
        MsgBox Last_name & " position is Parker"
    case else
        MsgBox "Hello " & Last_name & ", you are not part of the Management Team."
end select