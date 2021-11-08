Attribute VB_Name = "Module3"
Sub odd()
For j = 176 To 284
    da = Range("a" & j)
    For i = 2 To 31
        If Range("e" & j) = Range("r" & i) Then
        a = Range("s" & i) + 68.99
        End If
    Next
    For p = 2 To 31
        If Range("c" & j) = Range("r" & p) Then
        b = Range("s" & p)
        End If
    Next
    For k = 176 To 284
        If DateAdd("d", 1, Range("a" & k)) = da Then
            If Range("c" & j) = Range("c" & k) Or Range("c" & j) = Range("e" & k) Then
            a = a + 39.9
            End If
            If Range("e" & j) = Range("c" & k) Or Range("e" & j) = Range("e" & k) Then
            a = a - 39.9
            End If
        End If
    Next
    
    EA = 1 / (1 + 10 ^ ((b - a) / 400))
    EB = 1 - EA
    
   Range("f" & j) = EA
   Range("d" & j) = EB
   Range("j" & j) = 0.84855 / EB
   Range("k" & j) = 0.84855 / EA
   Range("l" & j) = 0.95 / EB
   Range("m" & j) = 0.95 / EA
   
    
Next

End Sub
