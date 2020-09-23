VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIPScan 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Port Sweep"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2775
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   2160
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock wnsConnection 
      Left            =   2040
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox StopIP4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "1"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "27374"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox StopIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "127"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox StopIP2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox StopIP3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox StartIP2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   2
      Text            =   " 0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox StartIP3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox StartIP4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox StartIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "127"
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3315
      Width           =   2535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   1370
      Left            =   110
      Top             =   1310
      Width           =   2570
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   8
      Left            =   1185
      Top             =   825
      Width           =   1125
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   7
      Left            =   2270
      Top             =   470
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   6
      Left            =   2270
      Top             =   110
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   5
      Left            =   1790
      Top             =   470
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   1790
      Top             =   110
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   1310
      Top             =   470
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   1310
      Top             =   110
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   830
      Top             =   470
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   830
      Top             =   110
      Width           =   400
   End
   Begin VB.Label CmdAction 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2830
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Port to Scan:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   855
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop:"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Start:"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmIPScan"
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
Dim Action As Integer
Dim StartIP As String, EndIP As String
Dim Seconds

Private Sub CmdAction_Click()
If Action = 1 Then
    CmdAction.Caption = "Start"
    wnsConnection.Close
    Action = 0
    Exit Sub
Else
    Action = 1
    CmdAction.Caption = "Stop"
    Call ScanPorts
End If
End Sub

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
Action = 0
CmdAction.Caption = "Start"
txtMessage.Locked = True
End Sub

Private Sub ScanPorts()
Dim Start As Integer
On Error Resume Next
Start = 0
EndIP = StopIP1 & "." & StopIP2 & "." & StopIP3 & "." & StopIP4
Do While Not StartIP = EndIP
    If Action = 0 Then
        Exit Do
    End If
    StartIP = Trim(StartIP1.Text) & "." & Trim(StartIP2.Text) & "." & Trim(StartIP3.Text) & "." & Trim(StartIP4.Text)
    If Start = 0 Then
        Open App.Path & "\" & "Scan.txt" For Append As #1
        Write #1, "Scan started with "; StartIP
        Close #1
    End If
    wnsConnection.Close
    wnsConnection.Connect StartIP, txtPort
    Seconds = 0
    Do While Seconds = 0
        DoEvents
    Loop
    If wnsConnection.State = 7 Then
        txtMessage.Text = txtMessage.Text & StartIP & vbNewLine
'///////////////////////////////////////////////////
        'Open App.Path & "\" & "Scan.txt" For Append As #1
        '   Write #1, StartIP, txtPort.Text
        'Close #1
'///////////////////////////////////////////////////
    End If
    StartIP4.Text = StartIP4.Text + 1
        If StartIP4.Text = "256" Then
            StartIP4.Text = "1"
            StartIP3.Text = StartIP3.Text + 1
            If StartIP3.Text = "256" Then
                StartIP3.Text = "1"
                StartIP2.Text = StartIP2.Text + 1
                If StartIP2.Text = "256" Then
                    StartIP2.Text = "1"
                    StartIP1.Text = StartIP1.Text + 1
                    If StartIP1.Text = "256" Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    Start = 1
Loop
Open App.Path & "\" & "Scan.txt" For Append As #1
    Write #1, "The last ip scanned was "; StartIP
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

'///////////////////////////////////////////////////
Private Sub StartIP1_GotFocus()
StartIP1.SelStart = 0
StartIP1.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP2_GotFocus()
StartIP2.SelStart = 0
StartIP2.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP3_GotFocus()
StartIP3.SelStart = 0
StartIP3.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP4_GotFocus()
StartIP4.SelStart = 0
StartIP4.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StopIP1_gotfocus()
StopIP1.SelStart = 0
StopIP1.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP2_gotfocus()
StopIP2.SelStart = 0
StopIP2.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP3_gotfocus()
StopIP3.SelStart = 0
StopIP3.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP4_gotfocus()
StopIP4.SelStart = 0
StopIP4.SelLength = Len(StopIP1.Text)
End Sub

Private Sub txtPort_gotfocus()
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
End Sub

Private Sub Timer_Timer()
Seconds = 1
End Sub



