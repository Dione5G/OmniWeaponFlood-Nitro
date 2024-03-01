Attribute VB_Name = "Push_Pull_M"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub Push_Pull()
If GetKeyPress(vbKey2) Then
On Error Resume Next
SendKeys ("{bs}") + (":push x") + ("{enter}")
2:
End If
If GetKeyPress(vbKey2) Then
GoTo 2
End If
If GetKeyPress(vbKey3) Then
On Error Resume Next
SendKeys ("{bs}") + (":pull x") + ("{enter}")
3:
End If
If GetKeyPress(vbKey3) Then
GoTo 3
End If
If GetKeyPress(vbKey4) Then
On Error Resume Next
SendKeys ("{bs}") + (":moonwalk") + ("{enter}")
4:
End If
If GetKeyPress(vbKey4) Then
GoTo 4
End If
If GetKeyPress(vbKey5) Then
On Error Resume Next
SendKeys ("{bs}") + (":sit") + ("{enter}")
5:
End If
If GetKeyPress(vbKey5) Then
GoTo 5
End If
If GetKeyPress(vbKey6) Then
On Error Resume Next
SendKeys ("{bs}") + (Form1.Combo1.Text) + ("{enter}")
6:
End If
If GetKeyPress(vbKey6) Then
GoTo 6
End If
If GetKeyPress(vbKey8) Then
On Error Resume Next
Form2.Show
Logger.Show
Logger.Left = Form2.Left
Logger.Top = Form2.Top + Form2.Height
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = False
Form1.Hide
Form1.Timer1.Enabled = False
8:
End If
If GetKeyPress(vbKey8) Then
GoTo 8
End If
End Sub
