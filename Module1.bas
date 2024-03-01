Attribute VB_Name = "OmniPush"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Dim pased As String
Dim X As Integer
Dim activ As String
Public Sub Omnipush_F8()
If GetKeyPress(vbKey1) Then
On Error Resume Next
SendKeys ("{bs}")
Logger.List1.Clear
Form2.Text3.Text = "0"
Form2.Text3.ForeColor = &H80FF80
1:
End If
If GetKeyPress(vbKey1) Then
GoTo 1
End If

If GetKeyPress(vbKey2) Then
On Error Resume Next
   pased = "1"
   SendKeys ("{bs}")
   SendKeys ("Susurrar") + (" ")
   SendKeys "^" + ("{a}")
   SendKeys "^" + ("{c}")
   SendKeys "{BACKSPACE}"
2:
End If
On Error Resume Next
If GetKeyPress(vbKey2) Then
   GoTo 2
End If

Copy = Clipboard.GetText

If Copy = "" Then
Else
   If pased = "1" Then
        Logger.List1.ForeColor = &H80FF80
        perla = Replace(Copy, "Susurrar ", "")
        limp = Replace(perla, " ", "")
        Logger.List1.AddItem " " & limp
        Form2.Text3.Text = Form2.Text3.Text + 1
        X = 0
        pased = "0"
      
   End If
End If
Clipboard.Clear
If Form2.Text3.Text = "30" Then
Form2.Text3.ForeColor = &H80FFFF
Logger.List1.ForeColor = &H80FFFF
End If
If Form2.Text3.Text = "60" Then
Form2.Text3.ForeColor = &H8080FF
Form2.adver.ForeColor = &H8080FF
Logger.List1.ForeColor = &H8080FF
Form2.adver.Caption = " Excess"
End If
If Form2.Text3.Text = "100" Then
Form2.Text3.ForeColor = vbRed
Form2.adver.ForeColor = vbRed
Logger.List1.ForeColor = vbRed
Form2.adver.Caption = " ¡Warning!"
End If
If Form2.Text3.Text = "300" Then
Form2.Text3.ForeColor = &HFF8080
Form2.adver.ForeColor = &HFF8080
Logger.List1.ForeColor = &HFF8080
Form2.adver.Caption = " God level"
End If
If Form2.Text3.Text = "500" Then
Form2.Text3.ForeColor = &HC000C0
Form2.adver.ForeColor = &HC000C0
Logger.List1.ForeColor = &HC000C0
Form2.adver.Caption = " Dione5G"
End If
If Form2.Text3.Text = "1000" Then
Form2.Text3.ForeColor = vbBlack
Form2.adver.ForeColor = vbBlack
Logger.List1.ForeColor = vbBlack
Form2.adver.BackStyle = 1
Form2.adver.Caption = " DarkNet"
End If
If GetKeyPress(vbKey3) Then
On Error Resume Next
pased = "0"
SendKeys ("{bs}")
activ = "ON"
X = 0
3:
End If
If GetKeyPress(vbKey3) Then
GoTo 3
End If
End Sub
Public Sub OmnipushB()
If activ = "ON" Then
  If X < Logger.List1.ListCount Then
     Clipboard.Clear
     Clipboard.SetText Logger.sep.Text & Logger.List1.List(X)
     SendKeys ("^") + ("v") + ("{enter}")
     X = X + 1
  Else
  activ = "OFF"
  End If
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
Clipboard.Clear
Clipboard.SetText Form2.Combo1.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
6:
End If
If GetKeyPress(vbKey6) Then
GoTo 6
End If
If GetKeyPress(vbKey8) Then
On Error Resume Next
Logger.Hide
Unload Logger
Form1.Show
Form2.Timer2.Enabled = False
Form2.Timer3.Enabled = False
Form2.Timer4.Enabled = False
Form2.Timer5.Enabled = False
Form2.Timer6.Enabled = False
Form2.omnipu.Enabled = False
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = True
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = True
Form2.Hide
Unload Form2
8:
End If
If GetKeyPress(vbKey8) Then
GoTo 8
End If
End Sub
