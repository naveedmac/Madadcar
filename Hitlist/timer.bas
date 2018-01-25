Attribute VB_Name = "timer"
Public total As Double
Dim t1 As Double
Dim t2 As Double


Sub startcounting()
t1 = Time

'If Minute(t1) > 1 Then
't1 = 0
'MsgBox "1 minute"
'End If


End Sub
Sub stopcounting()
total = total + Time - t1
If Second(total) > 10 Then
MsgBox "grater"
Else
MsgBox "less"
End If
End Sub

