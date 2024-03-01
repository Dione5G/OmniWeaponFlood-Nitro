VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic Control"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2850
   ClipControls    =   0   'False
   DrawMode        =   1  'Blackness
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "auto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2145
   ScaleMode       =   0  'User
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "set"
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox seg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   23
      Text            =   "00"
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox min 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   2
      TabIndex        =   22
      Text            =   "00"
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox hor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   2
      TabIndex        =   21
      Text            =   "00"
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox buffer 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox filtro 
      BackColor       =   &H008080FF&
      Height          =   735
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox blank 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Timer code 
      Interval        =   100
      Left            =   5280
      Top             =   0
   End
   Begin VB.Timer relo 
      Interval        =   100
      Left            =   2160
      Top             =   3120
   End
   Begin VB.Timer mftcontro2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   840
   End
   Begin VB.Timer mftcontro 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "B"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.Timer ONs 
      Interval        =   100
      Left            =   4560
      Top             =   360
   End
   Begin VB.Timer controlsON 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   1320
   End
   Begin VB.CommandButton stex 
      BackColor       =   &H00FFFFFF&
      Caption         =   "use text"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin VB.ComboBox selec 
      BackColor       =   &H00404040&
      ForeColor       =   &H008080FF&
      Height          =   315
      ItemData        =   "auto.frx":57E2
      Left            =   120
      List            =   "auto.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "add"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "200"
      ToolTipText     =   "Programar Segundos de Intervalo en Candidad de Veces"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "1000"
      ToolTipText     =   "Programar Segundos de Intervalo"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox C 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "3"
      ToolTipText     =   "Cantidad de Mensajes a Soltar por Segundo"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox V 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "Total de Veces a Repetir cada Segundos"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox MFT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   120
      MaxLength       =   100
      TabIndex        =   2
      Text            =   "Press ""use text"" enabled"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Timer contr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   1320
   End
   Begin VB.Label tim 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00:00:00"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Programar Ejecucion con Reloj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label IDENTIFIC 
      Alignment       =   2  'Center
      Caption         =   "none"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label R 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.Label reloj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Tag             =   ":pull x"
      ToolTipText     =   "Reloj"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label segc 
      BackStyle       =   0  'Transparent
      Caption         =   "Seg C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "F8 Empujar 2 o mas usuarios a la vez (Cuidado con el flood)"
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label segv 
      BackStyle       =   0  'Transparent
      Caption         =   "Seg V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "F8 Empujar 2 o mas usuarios a la vez (Cuidado con el flood)"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label fx3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N3 Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label fx2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N2 On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image pikachu 
      Height          =   720
      Left            =   1320
      Picture         =   "auto.frx":57E6
      ToolTipText     =   "Pikachu"
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   0
      Picture         =   "auto.frx":AFC8
      Stretch         =   -1  'True
      Top             =   -1560
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
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
Function numero(ByVal KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
Else
numero = KeyAscii
End If
If KeyAscii = 8 Then
numero = KeyAscii
End If
End Function
Function tiempoxx(ByVal KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
Else
tiempoxx = KeyAscii
End If
If KeyAscii = 8 Then
tiempoxx = KeyAscii
End If
End Function
Private Sub C_DblClick()
C.Text = ""
End Sub
Private Sub code_Timer()
If selec.Enabled = True Then
                        If IDENTIFIC.Caption = "Push_Pull" Then
If selec.Text = "N2" Then
blank.Text = ":push x"
ElseIf selec.Text = "N3" Then
blank.Text = ":pull x"
ElseIf selec.Text = "N4" Then
blank.Text = ":moonwalk"
ElseIf selec.Text = "N5" Then
blank.Text = ":sit"
ElseIf selec.Text = "N6" Then
blank.Text = Form1.Combo1.Text
End If
                         ElseIf IDENTIFIC.Caption = "OmniPush" Then
If selec.Text = "N3" Then
blank.Text = Logger.List1.Text
ElseIf selec.Text = "N4" Then
blank.Text = ":moonwalk"
ElseIf selec.Text = "N5" Then
blank.Text = ":sit"
ElseIf selec.Text = "N6" Then
blank.Text = Form2.Combo1.Text
End If
                         ElseIf IDENTIFIC.Caption = "Note_Push" Then
If selec.Text = "N3" Then
blank.Text = Npush.Text1.Text
ElseIf selec.Text = "N4" Then
blank.Text = Npush.Text2.Text
ElseIf selec.Text = "N5" Then
blank.Text = Npush.Text3.Text
ElseIf selec.Text = "N6" Then
blank.Text = Npush.Text4.Text
ElseIf selec.Text = "N7" Then
blank.Text = Npush.Text5.Text
ElseIf selec.Text = "N8" Then
blank.Text = Npush.Text6.Text
ElseIf selec.Text = "N9" Then
blank.Text = Npush.Text7.Text
ElseIf selec.Text = "N10" Then
blank.Text = Npush.Text8.Text
End If
                                       End If
   ElseIf MFT.Enabled = True Then
blank.Text = MFT.Text
End If
If InStr(blank.Text, "{") Then
bo1 = Replace(blank.Text, "{", "°coder1°coder1°")
filtro.Text = bo1
Else
bo1 = blank.Text
filtro.Text = bo1
End If
If InStr(bo1, "}") Then
bo2 = Replace(bo1, "}", "°coder2°coder2°")
filtro.Text = bo2
Else
bo2 = bo1
filtro.Text = bo2
End If
If InStr(bo2, "(") Then
bo3 = Replace(bo2, "(", "°coder3°coder3°")
filtro.Text = bo3
Else
bo3 = bo2
filtro.Text = bo3
End If
If InStr(bo3, ")") Then
bo4 = Replace(bo3, ")", "°coder4°coder4°")
filtro.Text = bo4
Else
bo4 = bo3
filtro.Text = bo4
End If
If InStr(bo4, "+") Then
bo5 = Replace(bo4, "+", "°coder5°coder5°")
filtro.Text = bo5
Else
bo5 = bo4
filtro.Text = bo5
End If
If InStr(bo5, "^") Then
bo6 = Replace(bo5, "^", "°coder6°coder6°")
filtro.Text = bo6
Else
bo6 = bo5
filtro.Text = bo6
End If
If InStr(bo6, "%") Then
bo7 = Replace(bo6, "%", "°coder7°coder7°")
filtro.Text = bo7
Else
bo7 = bo6
filtro.Text = bo7
End If
If InStr(bo7, "~") Then
bo8 = Replace(bo7, "~", "°coder8°coder8°")
filtro.Text = bo8
Else
bo8 = bo7
filtro.Text = bo8
End If
If InStr(filtro.Text, "°coder1°coder1°") Then
to1 = Replace(filtro.Text, "°coder1°coder1°", "{{}")
buffer.Text = to1
Else
to1 = filtro.Text
buffer.Text = to1
End If
If InStr(to1, "°coder2°coder2°") Then
to2 = Replace(to1, "°coder2°coder2°", "{}}")
buffer.Text = to2
Else
to2 = to1
buffer.Text = to2
End If
If InStr(to2, "°coder3°coder3°") Then
to3 = Replace(to2, "°coder3°coder3°", "{(}")
buffer.Text = to3
Else
to3 = to2
buffer.Text = to3
End If
If InStr(to3, "°coder4°coder4°") Then
to4 = Replace(to3, "°coder4°coder4°", "{)}")
buffer.Text = to4
Else
to4 = to3
buffer.Text = to4
End If
If InStr(to4, "°coder5°coder5°") Then
to5 = Replace(to4, "°coder5°coder5°", "{+}")
buffer.Text = to5
Else
to5 = to4
buffer.Text = to5
End If
If InStr(to5, "°coder6°coder6°") Then
to6 = Replace(to5, "°coder6°coder6°", "{^}")
buffer.Text = to6
Else
to6 = to5
buffer.Text = to6
End If
If InStr(to6, "°coder7°coder7°") Then
to7 = Replace(to6, "°coder7°coder7°", "{%}")
buffer.Text = to7
Else
to7 = to6
buffer.Text = to7
End If
If InStr(to7, "°coder8°coder8°") Then
to8 = Replace(to7, "°coder8°coder8°", "{~}")
buffer.Text = to8
End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "B" Then
                            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE
                            Command1.Caption = "U"
                            Command1.BackColor = &HC0C0C0
                            Command1.ToolTipText = "Press to return to Normal Window Mode"
                            MsgBox "ON: Always in Front Window Mode"
                            ElseIf Command1.Caption = "U" Then
                              SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                            Command1.Caption = "B"
                            Command1.BackColor = &HFFFFFF
                            Command1.ToolTipText = "Press to Keep the Window Always in Front"
                            MsgBox "OFF: Window Mode Always in Front Disabled"
                            End If
End Sub
Private Sub Command3_Click()
If Text5.Text = "" Then
MsgBox "Quantity Range is empty (Insert 1000 mls(1sec) or More)"
ElseIf Text5.Text > 65535 Then
MsgBox "It is not allowed to add more than 65,535 Mls"
Else
Timer2.Interval = Text5.Text
End If
End Sub
Private Sub Command4_Click()
If Command4.Caption = "H" Then
reloj.ForeColor = &H80FF80
Command4.Caption = "X"
hor.Text = Format(Now, "hh")
min.Text = Format(Now, "nn")
seg.Text = Format(Now, "ss")
tim.Caption = Format(Now, "hh:mm:ss")
Form3.Height = 3480
tim.ForeColor = &H404040
ElseIf Command4.Caption = "X" Then
reloj.ForeColor = &HFFFFFF
Command4.Caption = "H"
Form3.Height = 2565
End If
End Sub
Private Sub Command5_Click()
If hor.Text = "" Then
hor.Text = "00"
MsgBox "Do not leave the Hour Empty"
ElseIf min.Text = "" Then
min.Text = "00"
MsgBox "Do not leave empty minutes"
ElseIf seg.Text = "" Then
seg.Text = "00"
MsgBox "Don't leave empty seconds"
ElseIf CStr(Len(hor.Text)) < 2 Then
hor.Text = "0" & hor.Text
ElseIf CStr(Len(min.Text)) < 2 Then
min.Text = "0" & min.Text
ElseIf CStr(Len(seg.Text)) < 2 Then
seg.Text = "0" & seg.Text
Else
tim.Caption = hor.Text & ":" & min.Text & ":" & seg.Text
MsgBox "The messages will be executed automatically" & vbCrLf & "at " & tim.Caption
tim.Enabled = True
End If
End Sub
Private Sub contr_Timer()
If GetKeyPress(vbKey2) Then
If V.Text = "" Then
MsgBox "Times Box is Empty (Insert 1 or More)"
ElseIf C.Text = "" Then
MsgBox "Quantity Box is Empty (Insert 1 or More)"
Else
If selec.Text = "" Then
Timer1.Enabled = False
Timer2.Enabled = False
fx2.ForeColor = &HFF00&
fx2.ForeColor = &HFFC0C0
MsgBox "Select Available Options"
Else
fx2.ForeColor = &HFF00&
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
End If
End If
fcx2:
End If
If GetKeyPress(vbKey2) Then
GoTo fcx2
End If
If GetKeyPress(vbKey3) Then
C.Enabled = True
V.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
fx2.ForeColor = &HFFC0C0
End If
fx2.Caption = "N2 On"
fx3.Caption = "N3 Off"
fx2.ToolTipText = "Activate"
fx3.ToolTipText = "Stop"
End Sub
Private Sub controlsON_Timer()
If GetKeyPress(vbKey4) Then
If V.Text = "" Then
MsgBox "Times Box is Empty (Insert 1 or More)"
ElseIf C.Text = "" Then
MsgBox "Quantity Box is Empty (Insert 1 or More)"
Else
If selec.Text = "" Then
Timer1.Enabled = False
Timer2.Enabled = False
MsgBox "Select Available Options"
fx2.ForeColor = &HFF00&
fx2.ForeColor = &HFFC0C0
fx2.ForeColor = &HFF00&
fx2.ForeColor = &HFFC0C0
fx2.ForeColor = &HFF00&
fx2.ForeColor = &HFFC0C0
fx2.ForeColor = &HFF00&
fx2.ForeColor = &HFFC0C0
Else
fx2.ForeColor = &HFFC0C0
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
End If
End If
fcx3:
End If
If GetKeyPress(vbKey4) Then
GoTo fcx3
End If
If GetKeyPress(vbKey5) Then
C.Enabled = True
V.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
fx2.ForeColor = &HFFC0C0
End If
fx2.Caption = "F4 On"
fx3.Caption = "F5 Off"
fx2.ToolTipText = "Activar"
fx3.ToolTipText = "Detener"
End Sub
Private Sub Form_Load()
Command3.ToolTipText = "Press to Keep the window Always Visible"
ONs.Enabled = True
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub C_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub C_KeyUp(KeyCode As Integer, Shift As Integer)
R.Caption = C.Text
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub Command2_Click()
If Text4.Text = "" Then
MsgBox "Times Interval is empty (Insert 1000mls(1sec) or More)"
ElseIf Text4.Text > 65535 Then
MsgBox "It is not allowed to add more than 65,535 Mls"
Else
Timer1.Interval = Text4.Text
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub Form_Unload(Cancel As Integer)
ONs.Enabled = True
If IDENTIFIC.Caption = "Push_Pull" Then
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = True
Form1.numpush.Enabled = True
Form1.automatic.Enabled = True
ElseIf IDENTIFIC.Caption = "OmniPush" Then
Form2.Timer2.Enabled = False
Form2.Timer4.Enabled = False
Form2.Timer5.Enabled = True
Form2.Timer6.Enabled = True
Form2.numpush.Enabled = True
Form2.automatic.Enabled = True
ElseIf IDENTIFIC.Caption = "Note_Push" Then
Npush.Timer1.Enabled = True
Npush.contro.Enabled = True
Npush.automatic.Enabled = True
End If
Form3.contr.Enabled = False
Form3.controlsON.Enabled = False
End Sub
Private Sub hor_DblClick()
hor.Text = ""
End Sub

Private Sub min_DblClick()
min.Text = ""
End Sub
Private Sub seg_DblClick()
seg.Text = ""
End Sub
Private Sub hor_KeyPress(KeyAscii As Integer)
KeyAscii = tiempoxx(KeyAscii)
End Sub
Private Sub MFT_DblClick()
MFT.Text = ""
End Sub
Private Sub mftcontro_Timer()
If GetKeyPress(vbKey2) Then
If V.Text = "" Then
MsgBox "Times Box is Empty (Insert 1 or More)"
ElseIf C.Text = "" Then
MsgBox "Quantity Box is Empty (Insert 1 or More)"
Else
fx2.ForeColor = &HFF00&
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
stex.Enabled = False
End If
fcx5:
End If
If GetKeyPress(vbKey2) Then
GoTo fcx5
End If
If GetKeyPress(vbKey3) Then
C.Enabled = True
V.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
fx2.ForeColor = &HFFC0C0
stex.Enabled = True
End If
fx2.Caption = "N2 On"
fx3.Caption = "N3 Off"
fx2.ToolTipText = "Activate"
fx3.ToolTipText = "Stop"
End Sub
Private Sub mftcontro2_Timer()
If GetKeyPress(vbKey4) Then
If V.Text = "" Then
MsgBox "Times Box is Empty (Insert 1 or More)"
ElseIf C.Text = "" Then
MsgBox "Quantity Box is Empty (Insert 1 or More)"
Else
fx2.ForeColor = &HFF00&
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
stex.Enabled = False
End If
fcx4:
End If
If GetKeyPress(vbKey4) Then
GoTo fcx4
End If
If GetKeyPress(vbKey5) Then
C.Enabled = True
V.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
fx2.ForeColor = &HFFC0C0
stex.Enabled = True
End If
fx2.Caption = "N4 On"
fx3.Caption = "N5 Off"
fx2.ToolTipText = "Activate"
fx3.ToolTipText = "Stop"
End Sub
Private Sub min_KeyPress(KeyAscii As Integer)
KeyAscii = tiempoxx(KeyAscii)
End Sub
Private Sub ONs_Timer()
If MFT.Enabled = False Then
If IDENTIFIC.Caption = "Push_Pull" Then
contr.Enabled = True
mftcontro.Enabled = False
ONs.Enabled = False
ElseIf IDENTIFIC.Caption = "OmniPush" Then
controlsON.Enabled = True
mftcontro2.Enabled = False
ONs.Enabled = False
ElseIf IDENTIFIC.Caption = "Note_Push" Then
controlsON.Enabled = True
mftcontro2.Enabled = False
ONs.Enabled = False
End If
Else
If IDENTIFIC.Caption = "Push_Pull" Then
contr.Enabled = False
mftcontro.Enabled = True
ONs.Enabled = False
ElseIf IDENTIFIC.Caption = "OmniPush" Then
mftcontro2.Enabled = True
controlsON.Enabled = False
ONs.Enabled = False
ElseIf IDENTIFIC.Caption = "Note_Push" Then
mftcontro2.Enabled = True
controlsON.Enabled = False
ONs.Enabled = False
End If
End If
End Sub
Private Sub relo_Timer()
reloj.Caption = Format(Now, "hh:mm:ss")
If hor.Text = "" Then
Else
If hor.Text > 23 Then
hor.Text = "00"
End If
End If
If min.Text = "" Then
Else
If min.Text > 59 Then
min.Text = "59"
End If
End If
If seg.Text = "" Then
Else
If seg.Text > 59 Then
seg.Text = "59"
End If
End If
If tim.Enabled = True Then
tim.ForeColor = &HFF00&
If reloj.Caption = tim.Caption Then
If MFT.Enabled = True Then
If V.Text = "" Then
tim.ForeColor = &H404040
tim.Enabled = False
MsgBox "The Times Box is Empty (No Execution Made)"
ElseIf C.Text = "" Then
tim.ForeColor = &H404040
tim.Enabled = False
MsgBox "The Quantity Box is Empty (No Execution was Made)"
Else
fx2.ForeColor = &HFF00&
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
stex.Enabled = False
tim.Enabled = False
tim.ForeColor = &H404040
End If
Else
If selec.Text = "" Then
tim.ForeColor = &H404040
tim.Enabled = False
MsgBox "Select one of the F's or Press Use Text!"
Else
If V.Text = "" Then
tim.ForeColor = &H404040
tim.Enabled = False
MsgBox "The Times Box is Empty (No Execution Made)"
ElseIf C.Text = "" Then
tim.ForeColor = &H404040
tim.Enabled = False
MsgBox "The Quantity Box is Empty (No Execution was Made)"
Else
fx2.ForeColor = &HFF00&
Timer1.Enabled = True
C.Enabled = False
V.Enabled = False
stex.Enabled = False
tim.Enabled = False
tim.ForeColor = &H404040
End If
End If
End If
End If
End If
End Sub
Private Sub seg_KeyPress(KeyAscii As Integer)
KeyAscii = tiempoxx(KeyAscii)
End Sub
Private Sub stex_Click()
If MFT.Enabled = True Then
MFT.Enabled = False
selec.Enabled = True
ONs.Enabled = True
ElseIf selec.Enabled = True Then
selec.Enabled = False
MFT.Enabled = True
ONs.Enabled = True
End If
End Sub
Private Sub Text4_DblClick()
Text4.Text = ""
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub Text5_DblClick()
Text5.Text = ""
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
Private Sub Timer1_Timer()
If V.Text = 0 Then
C.Enabled = True
V.Text = 0
V.Enabled = True
R.Caption = 1
Timer1.Enabled = False
fx2.ForeColor = &HFFC0C0
stex.Enabled = True
Else
C.Text = R.Caption
Timer2.Enabled = True
V.Text = V.Text - 1
Timer1.Enabled = False
End If
End Sub
Private Sub Timer2_Timer()
If C.Text = 1 Then
SendKeys (buffer.Text) + ("{enter}")
     If V.Text >= 0 Then
      Timer1.Enabled = True
      End If
Timer2.Enabled = False
Else
SendKeys (buffer.Text) + ("{enter}")
C.Text = C.Text - 1
End If
End Sub
Private Sub V_DblClick()
V.Text = ""
End Sub
Private Sub V_KeyPress(KeyAscii As Integer)
KeyAscii = numero(KeyAscii)
End Sub
