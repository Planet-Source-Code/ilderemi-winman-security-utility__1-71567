VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmHTTPScan 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Text            =   "80"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Text            =   "46"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Text            =   "163"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   3
      Left            =   2400
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H00C0FFC0&
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   120
      Top             =   480
   End
   Begin VB.TextBox TxtTimeOut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   2400
      TabIndex        =   0
      Text            =   "2"
      Top             =   480
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   7
      Left            =   105
      Top             =   5520
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5565
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   6
      Left            =   2990
      Top             =   470
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   5
      Left            =   2990
      Top             =   110
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   4600
      Left            =   110
      Top             =   830
      Width           =   4480
   End
   Begin VB.Label CmdStop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   505
      Width           =   1575
   End
   Begin VB.Label CmdGo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   145
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   4
      Left            =   2390
      Top             =   470
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   3
      Left            =   2390
      Top             =   110
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   2
      Left            =   1790
      Top             =   110
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Index           =   1
      Left            =   1190
      Top             =   110
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   295
      Index           =   0
      Left            =   590
      Top             =   110
      Width           =   520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out (In Seconds):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmHTTPScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub CmdGo_Click()
If TxtIP(3).Text = 255 Then CmdStop_Click 'make sure it does not get past 255



TxtIP(3).Text = TxtIP(3).Text + 1 'change the ip

Winsock1.RemoteHost = TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text 'set the ip
Winsock1.RemotePort = 80
Winsock1.Connect 'try to connect to host

TimeOut.Enabled = False 'stop the time out (should already be stopped, but lets just make sure!)
TimeOut.interval = TxtTimeOut.Text * 1000 'set the time out

TimeOut.Enabled = True 'enable the timeout timer

End Sub

Private Sub CmdStop_Click()
Winsock1.Close 'close winsock
TimeOut.Enabled = False 'stop the time out

End Sub

Private Sub TimeOut_Timer()
ConnectionClose 'goto connectionclose sub
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next 'if winsock connects...

TimeOut.Enabled = False 'disable the time out
site = Inet1.OpenURL("http://" + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text, icByteArray) 'set the site

servers = Inet1.GetHeader("server") 'grab the server header (if there is one)

'you can grab any header you want, as long as it's there!, normal headers include, Content-type, Content-length, and Expires
'if you want to grab all the headers, then just use servers = Inet1.GetHeader()
'the headers you can get from kazaa are,
'X-Kazaa-Username
'X-Kazaa-Network
'X-Kazaa-IP
'X-Kazaa-SupernodeIP



If servers = "" Then 'if there isn't a server header, chances are it's going to be kazaa

    site = Inet1.OpenURL("http://" + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text, icByteArray)
    
    User = Inet1.GetHeader("X-Kazaa-Username")
    'so we try and get the kazaa username, (nothing important, just somthing i thought could be fun?!)
    TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + "Kazaa Username: " + User + vbNewLine

Else
    'show the http server type
    TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + servers + vbNewLine
End If


TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down the txtinfo box

Winsock1.Close 'close winsock
CmdGo_Click 'start again


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ConnectionClose 'call connection close sub

End Sub


Private Sub ConnectionClose()

TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + "NO Server" + vbNewLine 'obvisuly there was no server
TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down
TimeOut.Enabled = False 'stop time out
Winsock1().Close 'close winsock (if not already done)
CmdGo_Click 'start again
End Sub


