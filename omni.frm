VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "º  Dione5G - OmniWeaponFlood"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   870
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "omni.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "omni.frx":57E2
   ScaleHeight     =   1800
   ScaleMode       =   0  'User
   ScaleWidth      =   5415
   Begin VB.Timer omnipu 
      Interval        =   132
      Left            =   2160
      Top             =   2400
   End
   Begin VB.CommandButton automatic 
      Caption         =   "Auto"
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton numpush 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Npush"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton minies 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mini"
      Height          =   255
      Left            =   2760
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton blocker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Block"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Width           =   735
   End
   Begin VB.Timer Timer6 
      Interval        =   30
      Left            =   1440
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   21
      Text            =   "Add User"
      ToolTipText     =   "Toca una vez para Borrar el contenido de esta caja"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      ToolTipText     =   "Cantidad de Usuarios Agregados"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton addus 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   300
      Left            =   840
      TabIndex        =   18
      ToolTipText     =   "Puedes Presionar ""Enter"" para mas comodidad"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton addc 
      Caption         =   "Asig"
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      ToolTipText     =   "Asignar Tiempo a los Clicks"
      Top             =   1050
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "omni.frx":7B5B
      Left            =   1920
      List            =   "omni.frx":7B71
      TabIndex        =   11
      Text            =   "200"
      ToolTipText     =   "Tiempo 1000 = 1 seg / recomendado: 200 mlseg (Puedes editar a gusto)"
      Top             =   1050
      Width           =   855
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6000
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Interval        =   90
      Left            =   240
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Interval        =   150
      Left            =   6600
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6000
      Top             =   120
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "omni.frx":7B93
      Left            =   120
      List            =   "omni.frx":7BB5
      TabIndex        =   8
      Text            =   ":pikachu"
      ToolTipText     =   "F6 Mascotas (Puedes Personalizar)"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label fx6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label L5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Tag             =   ":sit"
      ToolTipText     =   "F5 :sit (Sentarse)"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label L4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MoonWalk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Tag             =   ":sit"
      ToolTipText     =   "F4 :Moonwalk (Caminar de reversa)"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label L3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OmniPush"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   ":moonwalk"
      ToolTipText     =   "F3 OmniPush (Lanzar OmniPush)"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label L2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "AddUser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   ":pull x"
      ToolTipText     =   "F2 AddUser (Adherir Usuario)"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label L1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DelList"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   ":push x"
      ToolTipText     =   "F1 DelList (Borrar todo lo de la Caja)"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label modomini 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image pikachu 
      Height          =   720
      Left            =   4800
      Picture         =   "omni.frx":7C10
      ToolTipText     =   "Pikachu"
      Top             =   -120
      Width           =   720
   End
   Begin VB.Label adver 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.Label fx9 
      BackStyle       =   0  'Transparent
      Caption         =   "NL7 One"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      ToolTipText     =   "AutoClick de un disparo por Seg asignado"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1150
      Width           =   135
   End
   Begin VB.Label fx11 
      BackStyle       =   0  'Transparent
      Caption         =   "NL9  off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      ToolTipText     =   "Detener AutoClic"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label fx10 
      BackStyle       =   0  'Transparent
      Caption         =   "NL8 Doble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      ToolTipText     =   "AutoDobleClick a disparo por Seg asignado"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label fx8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N8 off"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "F8 Empujar 2 o mas usuarios a la vez (Cuidado con el flood)"
      Top             =   70
      Width           =   735
   End
   Begin VB.Label Dione 
      BackStyle       =   0  'Transparent
      Caption         =   "Dione5G"
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
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "Iris"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label fx7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N7 TP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      ToolTipText     =   "F7 TurboPush x30 "":push x"""
      Top             =   70
      Width           =   735
   End
   Begin VB.Label fx5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label fx4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label fx3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.Label fx2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label fx1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Shape aro 
      BorderColor     =   &H00FFFF80&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3600
      Shape           =   2  'Oval
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Private Const KEY_TOGGLED As Integer = &H1
Private Const KEY_PRESSED As Integer = &H1000
Option Explicit
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '
Const SW_SHOWNORMAL = 1
Private Declare Function SetWindowPos _
    Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal wFlags As Long) As Long
Private Sub addc_Click()
If Combo2.Text > 65535 Then
MsgBox "It is not allowed to add more than 60,000 mls."
Else
Timer2.Interval = Combo2.Text
Timer4.Interval = Combo2.Text
End If
End Sub
Private Sub automatic_Click()
Form3.Show
Timer2.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
numpush.Enabled = False
automatic.Enabled = False
Form3.IDENTIFIC.Caption = "OmniPush"
Form3.selec.AddItem "N3"
Form3.selec.AddItem "N4"
Form3.selec.AddItem "N5"
Form3.selec.AddItem "N6"
Form3.controlsON.Enabled = True
End Sub
Private Sub blocker_Click()
If blocker.Caption = "Block" Then
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE
    blocker.Caption = "UnBlock"
    blocker.BackColor = &HC0C0C0
    blocker.ToolTipText = "Press to return to Normal Window Mode"
    MsgBox "ON: Always in Front Window Mode"
ElseIf blocker.Caption = "UnBlock" Then
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    blocker.Caption = "Block"
    blocker.BackColor = &HFFFFFF
    blocker.ToolTipText = "Press to Keep the Window Always in Front"
    MsgBox "OFF: Window Mode Always in Front Disabled"
ElseIf blocker.Caption = "B" Then
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE
    blocker.Caption = "U"
    blocker.BackColor = &HC0C0C0
    blocker.ToolTipText = "Press to return to Normal Window Mode"
    MsgBox "ON: Always in Front Window Mode"
ElseIf blocker.Caption = "U" Then
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    blocker.Caption = "B"
    blocker.BackColor = &HFFFFFF
    blocker.ToolTipText = "Press to Keep the Window Always in Front"
    MsgBox "OFF: Window Mode Always in Front Disabled"
End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = TextRT(KeyAscii)
End Sub
Function TextRT(ByVal KeyAscii As Integer)
If InStr("-_=.:qwertyuiopasdfghjklñzxcvbnmQWERTYUIOPASDFGHJKLÑZXCVBNM1234567890 ", Chr(KeyAscii)) = 0 Then
Else
TextRT = KeyAscii
End If
If KeyAscii = 8 Then
TextRT = KeyAscii
End If
End Function
Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Function numero(ByVal KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
Else
numero = KeyAscii
End If
If KeyAscii = 8 Then
numero = KeyAscii
End If
End Function
Private Sub Command1_Click()
Timer2.Interval = Combo2.Text
Timer4.Interval = Combo2.Text
End Sub
Private Sub addus_Click()
  If Text3.Text = "0" Then
      Logger.List1.ForeColor = &H80FF80
Logger.List1.Text = Logger.sep.Text & " " & Text1.Text
Text1.Text = ""
Text3.Text = Text3.Text + 1
Else
Logger.List1.AddItem Logger.sep.Text & " " & Text1.Text
Text1.Text = ""
Text3.Text = Text3.Text + 1
End If
If Text3.Text = "30" Then
Text3.ForeColor = &H80FFFF
Logger.List1.ForeColor = &H80FFFF
End If
If Text3.Text = "60" Then
Text3.ForeColor = &H8080FF
adver.ForeColor = &H8080FF
Logger.List1.ForeColor = &H8080FF
adver.Caption = " Excess"
End If
If Text3.Text = "100" Then
Text3.ForeColor = vbRed
adver.ForeColor = vbRed
Logger.List1.ForeColor = vbRed
adver.Caption = " ¡Warning!"
End If
If Text3.Text = "300" Then
Text3.ForeColor = &HFF8080
adver.ForeColor = &HFF8080
Logger.List1.ForeColor = &HFF8080
adver.Caption = " God Level"
End If
If Text3.Text = "500" Then
Text3.ForeColor = &HC000C0
adver.ForeColor = &HC000C0
Logger.List1.ForeColor = &HC000C0
adver.Caption = " Dione5G"
End If
If Text3.Text = "1000" Then
Text3.ForeColor = vbBlack
adver.ForeColor = vbBlack
Logger.List1.ForeColor = vbBlack
adver.BackStyle = 1
adver.Caption = " DarkNet"
End If
End Sub
Private Sub Dione_Click()
ShellExecute Me.hWnd, "open", "https://www.youtube.com/channel/UC7L97ZkpVRVceO-sqE9NhJw", "", "C:\\", SW_SHOWNORMAL
ShellExecute Me.hWnd, "open", "https://www.instagram.com/Dione5G", "", "C:\\", SW_SHOWNORMAL
ShellExecute Me.hWnd, "open", "https://github.com/Dione5G", "", "C:\\", SW_SHOWNORMAL
End Sub
Private Sub Form_Load()
Form2.Top = Form1.Top
Form2.Left = Form1.Left
Timer2.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = False
Timer5.Enabled = True
Timer6.Enabled = True
omnipu.Enabled = True
Clipboard.Clear
blocker.ToolTipText = "Press to Keep the window Always Visible"
Logger.Show
End Sub

Private Sub Image1_Click()

End Sub

Private Sub minies_Click()
If minies.Caption = "Mini" Then
minies.Caption = "M"
minies.BackColor = &HC0C0C0
ElseIf minies.Caption = "M" Then
minies.Caption = "Mini"
minies.BackColor = &HFFFFFF
End If
If blocker.Caption = "B" Then
blocker.Caption = "Block"
ElseIf blocker.Caption = "Block" Then
blocker.Caption = "B"
ElseIf blocker.Caption = "UnBlock" Then
blocker.Caption = "U"
ElseIf blocker.Caption = "U" Then
blocker.Caption = "UnBlock"
End If
omnimodomini
End Sub
Private Sub numpush_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Npush.Show
Form2.Hide
Unload Form2
End Sub
Private Sub omnipu_Timer()
Omnipush_F8
End Sub
Private Sub Text1_Click()
Text1.Text = ""
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = TextRT(KeyAscii)
End Sub
Private Sub Timer2_Timer()
oneclic
End Sub
Private Sub Timer4_Timer()
dobleclic
End Sub
Private Sub Timer3_Timer()
autoclickOmni
End Sub
Private Sub Timer5_Timer()
Turbo_Push
End Sub
Private Sub Timer6_Timer()
OmnipushB
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Logger
Unload Form3
End Sub
