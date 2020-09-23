VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSysEXE 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   855
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   500
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   4425
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   500
      Width           =   975
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   500
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   1905
      Top             =   480
      Width           =   1005
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   825
      Top             =   480
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   5040
      Top             =   105
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   830
      Top             =   110
      Width           =   4240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   5110
      TabIndex        =   1
      Top             =   135
      Width           =   255
   End
End
Attribute VB_Name = "frmSysEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'|\\\\\\\\____________________////////|'
'||\\\\\\\\__________________////////||'
'|||\\\\\\\\________________////////|||'
'||||\\\\\\\\______________////////||||'
'|||||\\\\\\\\____________////////|||||'
'||||||\\\\\\\\__________////////||||||'
'|||||| \\\\\\\\________//////// ||||||'
'||||||  \\\\\\\\______////////  ||||||'
'||||||   \\\\\\\\____////////   ||||||'
'||||||    \\\\\\\\  ////////    ||||||'
'||||||     \\\\\\\\////////     ||||||'
'||||||                          ||||||'
'||||||                          ||||||'
'||||||    !!!!!!!!!!!!!!!!!!    ||||||'
'||||||    ! Amazing WinMan !    ||||||'
'||||||    !!!!!!!!!!!!!!!!!!    ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||........:iLDEREMi:........||||||'
Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Set myShadow = New clsShadow
  With myShadow
    If .Shadow(Me) Then
      .Depth = 10
      .Transparency = 128
    Else
      Set myShadow = Nothing
    End If
  End With
End Sub

Private Sub isButton1_Click()
    Dim mmin As Integer
    Dim hhor As Integer
    hhor = Val(Left(Time$, 2))
    mmin = Val((Mid(Time$, 4, 2))) + 1
    If mmin > 59 Then
        mmin = mmin - 60
        hhor = hhor + 1
    End If
    If hhor > 23 Then hhor = hhor - 24

    Shell "cmd /c at " & hhor & ":" & mmin & " /interactive " & Text1.Text, vbHide
End Sub

Private Sub isButton2_Click()
Shell "cmd /c at /delete /yes"
frmAntiSVCHOST.Show
Unload Me
End Sub

Private Sub Label1_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.Filename
End Sub

Private Sub Label5_Click()
Unload Me
End Sub
