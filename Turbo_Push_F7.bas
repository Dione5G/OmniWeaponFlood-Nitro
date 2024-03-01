Attribute VB_Name = "Turbo_Push_F7"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub Turbo_Push()
If GetKeyPress(vbKey7) Then
On Error Resume Next
SendKeys ("{bs}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
7:
End If
If GetKeyPress(vbKey7) Then
GoTo 7
End If
End Sub
