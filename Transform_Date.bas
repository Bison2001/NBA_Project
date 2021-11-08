Attribute VB_Name = "Module2"
Sub datetrans()

Dim LString As String
Dim LArray() As String

For q = 176 To 284
    LString = Range("a" & q)
    LArray = Split(LString, ",")
    Range("a" & q) = LArray(1) + LArray(2)
Next
End Sub


