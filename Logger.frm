VERSION 5.00
Begin VB.Form Logger 
   BackColor       =   &H00404040&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "              List-Users OWF"
   ClientHeight    =   2415
   ClientLeft      =   2925
   ClientTop       =   1770
   ClientWidth     =   2505
   ControlBox      =   0   'False
   Icon            =   "Logger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0080FF80&
      Height          =   1980
      ItemData        =   "Logger.frx":57E2
      Left            =   120
      List            =   "Logger.frx":57E4
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton ListBlock 
      Caption         =   "Block"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton loki 
      Caption         =   "Unlock"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox sep 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   840
      MaxLength       =   100
      TabIndex        =   0
      Text            =   ":push"
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   600
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   3000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Function setlogger(ByVal KeyAscii As Integer)
If InStr("-_=.:qwertyuiopasdfghjklñzxcvbnmQWERTYUIOPASDFGHJKLÑZXCVBNM1234567890", Chr(KeyAscii)) = 0 Then
Else
setlogger = KeyAscii
End If
If KeyAscii = 8 Then
setlogger = KeyAscii
End If
End Function
Private Sub Form_Load()
Unload Npush
End Sub
Private Sub List1_DblClick()
If List1.ListIndex <> -1 Then
   List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub ListBlock_Click()
If ListBlock.Caption = "Block" Then
    loki.Caption = "Lock"
    loki.Enabled = False
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE
    ListBlock.Caption = "UnBlock"
    ListBlock.BackColor = &H404040
    ListBlock.ToolTipText = "Press to return to Normal Window Mode"
    MsgBox "ON: List-User Always in front"
    Timer1.Enabled = False
    ElseIf ListBlock.Caption = "UnBlock" Then
    loki.Caption = "Unlock"
    loki.Enabled = True
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    ListBlock.Caption = "Block"
    ListBlock.BackColor = &HFFFFFF
    ListBlock.ToolTipText = "Press to Keep the Window Always in Front"
    MsgBox "OFF: List-User Always in front Disabled"
    Timer1.Enabled = True
End If
End Sub
Private Sub loki_Click()
If loki.Caption = "Unlock" Then
loki.Caption = "Lock"
Timer1.Enabled = False
ElseIf loki.Caption = "Lock" Then
loki.Caption = "Unlock"
Timer1.Enabled = True
End If
End Sub
Private Sub sep_DblClick()
sep.Text = ""
End Sub
Private Sub sep_KeyPress(KeyAscii As Integer)
KeyAscii = setlogger(KeyAscii)
End Sub

Private Sub Timer1_Timer()
Logger.Left = Form2.Left
Logger.Top = Form2.Top + Form2.Height
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Text2.Width = (Logger.Width - 480)
Text2.Height = (Logger.Height - 1000)
End Sub
