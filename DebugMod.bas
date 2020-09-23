Attribute VB_Name = "DebugMod"
Option Explicit

Public Dbg As New Debugger

Public Sub Temp()
Dim a() As Single
ReDim a(1 To 9)
a(1) = 64547.65
a(2) = 64547.68
a(3) = 64547.76
a(4) = 64547.78
a(5) = 64547.82
a(6) = 64547.84
a(7) = 64548.12
a(8) = 64548.15
a(9) = 64548.18
Dim i As Integer
For i = 2 To 9
    Debug.Print i, a(i) - a(i - 1)
Next
End Sub

