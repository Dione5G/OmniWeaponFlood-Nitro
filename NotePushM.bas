Attribute VB_Name = "NotePushM"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub NotePushC()
If GetKeyPress(vbKey1) Then
On Error Resume Next
Npush.Text1.Text = "-"
Npush.Text2.Text = "-"
Npush.Text3.Text = "-"
Npush.Text4.Text = "-"
Npush.Text5.Text = "-"
Npush.Text6.Text = "-"
Npush.Text7.Text = "-"
Npush.Text8.Text = "-"
Npush.cant.Caption = "0"
Npush.Caption = "° Dione5G - NotePush (0/8)"
Npush.Nl1.Visible = False
Npush.Nl2.Visible = False
Npush.Nl3.Visible = False
Npush.Nl4.Visible = False
Npush.Nl5.Visible = False
Npush.Nl6.Visible = False
Npush.Nl7.Visible = False
Npush.Nl8.Visible = False
1:
End If
If GetKeyPress(vbKey1) Then
GoTo 1
End If
If GetKeyPress(vbKey2) Then
On Error Resume Next
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
    If Npush.cant.Caption = "8" Then
    MsgBox "8/8 You cannot add more people"
    Else
      If Npush.Text1.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text1.Text = Npush.sete.Text & " " & limp
     Npush.cant.Caption = Npush.cant.Caption + 1
   Npush.Nl1.Visible = True
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      ElseIf Npush.Text2.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text2.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl2.Visible = True
      ElseIf Npush.Text3.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text3.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl3.Visible = True
      ElseIf Npush.Text4.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text4.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl4.Visible = True
      ElseIf Npush.Text5.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text5.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl5.Visible = True
      ElseIf Npush.Text6.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text6.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl6.Visible = True
      ElseIf Npush.Text7.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text7.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl7.Visible = True
      ElseIf Npush.Text8.Text = "-" Then
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Npush.Text8.Text = Npush.sete.Text & " " & limp
      Npush.cant.Caption = Npush.cant.Caption + 1
      Npush.Caption = "° Dione5G - NotePush (" + Npush.cant.Caption + "/8)"
      Npush.Nl8.Visible = True
      End If
      End If
      Clipboard.Clear
    End If
End Sub
Public Sub NotePushD()

If GetKeyPress(vbKeyNumpad9) Then
Form1.Show
Npush.Timer1.Enabled = False
Npush.copys.Enabled = False
Npush.contro.Enabled = False
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = True
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = True
Npush.Hide
Unload Npush
n11:
End If
If GetKeyPress(vbKeyNumpad9) Then
GoTo n11
End If
End Sub
