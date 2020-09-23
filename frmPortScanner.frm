VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPortScanner 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   4320
   End
   Begin VB.ListBox lstLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4820
      Width           =   2775
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   120
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00008000&
      Height          =   3405
      Left            =   105
      Top             =   1305
      Width           =   2805
   End
   Begin VB.Label cmdStop 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   910
      Width           =   1335
   End
   Begin VB.Label cmdScan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   910
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   1560
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   120
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   1070
      Top             =   470
      Width           =   1840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   1070
      Top             =   110
      Width           =   1840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Port:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmPortScanner"
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
Private Sub cmdScan_Click()
If txtHost.Text = "" Then
MsgBox "You must enter an IP address!", vbInformation
Else
lstLog.clear
Timer1.interval = 1
Timer1.Enabled = True
End If
End Sub
Private Sub cmdStop_Click()
Timer1.Enabled = False
txtPort.Text = "1"
End Sub

Private Sub Form_Load()
txtPort.Text = "1"
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

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Winsock1_Connect()
lstLog.AddItem (Winsock1.RemotePort & " is open")
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Winsock1.Close
txtPort.Text = Int(txtPort.Text) + 1
Winsock1.RemoteHost = txtHost.Text
Winsock1.RemotePort = txtPort.Text
Winsock1.Connect
End Sub
