option explicit

Dim TestArray, Item

TestArray = Array(1,2,3,4,5,6,7,8,9,10)

for Each Item in TestArray
    MsgBox Item, 0, "Item in array: "
Next