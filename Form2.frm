VERSION 5.00
Begin VB.Form Npush 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "° Dione5G - NotePush (0/8)"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":57E2
   ScaleHeight     =   2190
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer copys 
      Interval        =   120
      Left            =   600
      Top             =   2760
   End
   Begin VB.CommandButton automatic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto"
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton blo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Block"
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Minin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mini"
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox sete 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   4200
      MaxLength       =   100
      TabIndex        =   21
      Text            =   ":push"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Timer contro 
      Interval        =   120
      Left            =   1560
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   2760
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "-"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "-"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "-"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "-"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "-"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "-"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "-"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "-"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label modomini 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   23
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label cant 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Dione 
      BackStyle       =   0  'Transparent
      Caption         =   "D5G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.Image pikachu 
      Height          =   720
      Left            =   4560
      Picture         =   "Form2.frx":8267
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Ls3 
      BackStyle       =   0  'Transparent
      Caption         =   "NL9 Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Ls2 
      BackStyle       =   0  'Transparent
      Caption         =   "N2 Copy N*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Ls1 
      BackStyle       =   0  'Transparent
      Caption         =   "N1 DelAll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Nl8 
      BackStyle       =   0  'Transparent
      Caption         =   "n0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl7 
      BackStyle       =   0  'Transparent
      Caption         =   "n9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl6 
      BackStyle       =   0  'Transparent
      Caption         =   "n8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl5 
      BackStyle       =   0  'Transparent
      Caption         =   "n7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl4 
      BackStyle       =   0  'Transparent
      Caption         =   "n6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl3 
      BackStyle       =   0  'Transparent
      Caption         =   "n5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl2 
      BackStyle       =   0  'Transparent
      Caption         =   "n4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nl1 
      BackStyle       =   0  'Transparent
      Caption         =   "n3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   3480
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape aro 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3960
      Shape           =   2  'Oval
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Npush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '
Private Declare Function SetWindowPos _
    Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal wFlags As Long) As Long
Public dataset As String
Private Sub automatic_Click()
Form3.Show
Timer1.Enabled = False
contro.Enabled = False
automatic.Enabled = False
Form3.IDENTIFIC.Caption = "Note_Push"
Form3.selec.AddItem "F3"
Form3.selec.AddItem "F4"
Form3.selec.AddItem "F5"
Form3.selec.AddItem "F6"
Form3.selec.AddItem "F7"
Form3.selec.AddItem "F8"
Form3.selec.AddItem "F9"
Form3.selec.AddItem "F10"
Form3.controlsON.Enabled = True
End Sub
Private Sub blo_Click()
If blo.Caption = "Block" Then
            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE
            blo.Caption = "UnBlock"
            blo.BackColor = &HC0C0C0
            blo.ToolTipText = "Press to return to Normal Window Mode"
            MsgBox "ON: Always in Front Window Mode"
        ElseIf blo.Caption = "UnBlock" Then
            SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            blo.Caption = "Block"
            blo.BackColor = &HFFFFFF
            blo.ToolTipText = "Press to Keep the Window Always in Front"
            MsgBox "OFF: Window Mode Always in Front Disabled"
        ElseIf blo.Caption = "B" Then
            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE
            blo.Caption = "U"
            blo.BackColor = &HC0C0C0
            blo.ToolTipText = "Press to return to Normal Window Mode"
            MsgBox "ON: Press to return to Normal Window Mode"
        ElseIf blo.Caption = "U" Then
            SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            blo.Caption = "B"
            blo.BackColor = &HFFFFFF
            blo.ToolTipText = "Press to Keep the Window Always in Front"
            MsgBox "OFF: Window Mode Always in Front Disabled"
        End If
End Sub
Private Sub Command3_Click()
If Npush.Width = 4395 Then
Npush.Width = 2985
Command3.Top = 2000
Command3.Left = 120
Command3.Caption = "+"
Else
Command3.Caption = "-"
Npush.Width = 4395
Command3.Top = 2280
Command3.Left = 3960
End If
End Sub
Private Sub contro_Timer()
NotePushD
End Sub
Private Sub copys_Timer()
NotePushC
End Sub

Private Sub Form_Load()
Clipboard.Clear
Timer1.Enabled = True
contro.Enabled = True
copys.Enabled = True
Unload Logger
Npush.Top = Form1.Top
Npush.Left = Form1.Left
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Logger
Unload Form3
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Minin_Click()
If Minin.Caption = "Mini" Then
Minin.Caption = "M"
Minin.BackColor = &HC0C0C0
ElseIf Minin.Caption = "M" Then
Minin.Caption = "Mini"
Minin.BackColor = &HFFFFFF
End If
If blo.Caption = "B" Then
blo.Caption = "Block"
ElseIf blo.Caption = "Block" Then
blo.Caption = "B"
ElseIf blo.Caption = "UnBlock" Then
blo.Caption = "U"
ElseIf blo.Caption = "U" Then
blo.Caption = "UnBlock"
End If
npushminim
End Sub
Function TextRTx(ByVal KeyAscii As Integer)
If InStr("-_=.:qwertyuiopasdfghjklñzxcvbnmQWERTYUIOPASDFGHJKLÑZXCVBNM1234567890", Chr(KeyAscii)) = 0 Then
Else
TextRTx = KeyAscii
End If
If KeyAscii = 8 Then
TextRTx = KeyAscii
End If
End Function
Private Sub sete_DblClick()
sete.Text = ""
End Sub
Private Sub sete_KeyPress(KeyAscii As Integer)
KeyAscii = TextRTx(KeyAscii)
End Sub
Private Sub Timer1_Timer()
If Text1.Text <> "-" Then
If GetKeyPress(vbKey3) Then
On Error Resume Next
If InStr(Text1.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text1.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text1.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N1:
End If
End If
If GetKeyPress(vbKey3) Then
GoTo N1
End If
End If
If Text2.Text <> "-" Then
If GetKeyPress(vbKey4) Then
On Error Resume Next
If InStr(Text2.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text2.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text2.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N2:
End If
End If
If GetKeyPress(vbKey4) Then
GoTo N2
End If
End If

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text3.Text <> "-" Then
If GetKeyPress(vbKey5) Then
On Error Resume Next
If InStr(Text3.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text3.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text3.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N3:
End If
End If
If GetKeyPress(vbKey5) Then
GoTo N3
End If
End If


'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text4.Text <> "-" Then
If GetKeyPress(vbKey6) Then
On Error Resume Next
If InStr(Text4.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text4.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text4.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N4:
End If
End If
If GetKeyPress(vbKey6) Then
GoTo N4
End If
End If

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text5.Text <> "-" Then
If GetKeyPress(vbKey7) Then
On Error Resume Next
If InStr(Text5.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text5.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text5.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N5:
End If
End If
If GetKeyPress(vbKey7) Then
GoTo N5
End If
End If

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text6.Text <> "-" Then
If GetKeyPress(vbKey8) Then
On Error Resume Next
If InStr(Text6.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text6.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text6.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N6:
End If
End If
If GetKeyPress(vbKey8) Then
GoTo N6
End If
End If

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text7.Text <> "-" Then
If GetKeyPress(vbKey9) Then
On Error Resume Next
If InStr(Text7.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text7.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text7.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N7:
End If
End If
If GetKeyPress(vbKey9) Then
GoTo N7
End If
End If

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Text8.Text <> "-" Then
If GetKeyPress(vbKey0) Then
On Error Resume Next
If InStr(Text8.Text, "}") Then
MsgBox "Cannot use character ( } )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "{") Then
MsgBox "Cannot use character ( { )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "%") Then
MsgBox "Cannot use character ( % )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "^") Then
MsgBox "Cannot use character ( ^ )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "~") Then
MsgBox "Cannot use character ( ~ )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "+") Then
MsgBox "Cannot use character ( + )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, ")") Then
MsgBox "Cannot use character ( ) )" + vbCrLf + "  in this version "
ElseIf InStr(Text8.Text, "(") Then
MsgBox "Cannot use character ( ( )" + vbCrLf + "  in this version "
Else
Clipboard.Clear
Clipboard.SetText Text8.Text
SendKeys ("{bs}") + ("^") + ("v") + ("{enter}")
N8:
End If
End If
If GetKeyPress(vbKey0) Then
GoTo N8
End If
End If
End Sub

