MyText = "The quick brown fox jumps over the lazy dog"

arrWords = Split(MyText, " ")

TotalNoOfWords = UBound(arrWords) + 1

MsgBox "Total number of words " & TotalNoOfWords, 64, "Word Count :"

text1 = mid(MyText, 21, 5)

MsgBox "Extracted Characters : " & text1, 64, "Extracted :"