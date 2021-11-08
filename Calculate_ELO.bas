Attribute VB_Name = "Module1"
Sub ELO()
Dim a As Double
Dim b As Double

For j = 2 To 175
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
    For k = 2 To 175
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
    
    For m = 2 To 32
        If Range("c" & j) = Range("r" & m) Then
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 15 Then
            Range("s" & m) = Range("s" & m) + 64 * (1 - EB)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 10 And Range("d" & j) - Range("f" & j) < 15 Then
            Range("s" & m) = Range("s" & m) + 48 * (1 - EB)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 6 And Range("d" & j) - Range("f" & j) < 10 Then
            Range("s" & m) = Range("s" & m) + 32 * (1 - EB)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) > 0 And Range("d" & j) - Range("f" & j) < 6 Then
            Range("s" & m) = Range("s" & m) + 16 * (1 - EB)
            End If
        End If
    Next
    For u = 2 To 32
        If Range("e" & j) = Range("r" & u) Then
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 15 Then
            Range("s" & u) = Range("s" & u) + 64 * (0 - EA)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 10 And Range("d" & j) - Range("f" & j) < 15 Then
            Range("s" & u) = Range("s" & u) + 64 * (0 - EA)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) >= 6 And Range("d" & j) - Range("f" & j) < 10 Then
            Range("s" & u) = Range("s" & u) + 64 * (0 - EA)
            End If
            If Range("d" & j) > Range("f" & j) And Range("d" & j) - Range("f" & j) > 0 And Range("d" & j) - Range("f" & j) < 6 Then
            Range("s" & u) = Range("s" & u) + 64 * (0 - EA)
            End If
        End If
    Next
    
    For n = 2 To 32
        If Range("e" & j) = Range("r" & n) Then
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 15 Then
            Range("s" & n) = Range("s" & n) + 64 * (1 - EA)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 10 And Range("f" & j) - Range("d" & j) < 15 Then
            Range("s" & n) = Range("s" & n) + 48 * (1 - EA)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 6 And Range("f" & j) - Range("d" & j) < 10 Then
            Range("s" & n) = Range("s" & n) + 32 * (1 - EA)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) > 0 And Range("f" & j) - Range("d" & j) < 6 Then
            Range("s" & n) = Range("s" & n) + 16 * (1 - EA)
            End If
        End If
    Next
    
    For v = 2 To 32
        If Range("c" & j) = Range("r" & v) Then
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 15 Then
            Range("s" & v) = Range("s" & v) + 64 * (0 - EB)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 10 And Range("f" & j) - Range("d" & j) < 15 Then
            Range("s" & v) = Range("s" & v) + 48 * (0 - EB)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) >= 6 And Range("f" & j) - Range("d" & j) < 10 Then
            Range("s" & v) = Range("s" & v) + 32 * (0 - EB)
            End If
            If Range("f" & j) > Range("d" & j) And Range("f" & j) - Range("d" & j) > 0 And Range("f" & j) - Range("d" & j) < 6 Then
            Range("s" & v) = Range("s" & v) + 16 * (0 - EB)
            End If
        End If
    Next
    
    
Next

            
End Sub
