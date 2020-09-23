VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGetPost 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
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
   ScaleHeight     =   7215
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeVariables 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Submission Variables"
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   135
      TabIndex        =   20
      ToolTipText     =   "Use this space to submit some custom vairables."
      Top             =   480
      Width           =   9495
      Begin VB.PictureBox pbxOVariables 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   10
         ScaleHeight     =   1500
         ScaleWidth      =   8175
         TabIndex        =   21
         Top             =   10
         Width           =   8175
         Begin VB.VScrollBar vsbVariables 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   7920
            Max             =   0
            TabIndex        =   35
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox pbxVariables 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1455
            ScaleWidth      =   7935
            TabIndex        =   22
            Top             =   0
            Width           =   7935
            Begin VB.TextBox txtVariableValue 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   3
               Left            =   4680
               TabIndex        =   41
               Top             =   1150
               Width           =   3015
            End
            Begin VB.TextBox txtVariableName 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   3
               Left            =   840
               TabIndex        =   40
               Top             =   1150
               Width           =   3015
            End
            Begin VB.TextBox txtVariableValue 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   0
               Left            =   4680
               TabIndex        =   28
               Top             =   70
               Width           =   3015
            End
            Begin VB.TextBox txtVariableValue 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   1
               Left            =   4680
               TabIndex        =   27
               Top             =   430
               Width           =   3015
            End
            Begin VB.TextBox txtVariableValue 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   2
               Left            =   4680
               TabIndex        =   26
               Top             =   790
               Width           =   3015
            End
            Begin VB.TextBox txtVariableName 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   25
               Top             =   70
               Width           =   3015
            End
            Begin VB.TextBox txtVariableName 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   24
               Top             =   430
               Width           =   3015
            End
            Begin VB.TextBox txtVariableName 
               BackColor       =   &H00181818&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   2
               Left            =   840
               TabIndex        =   23
               Top             =   790
               Width           =   3015
            End
            Begin VB.Shape ShapeB2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   3
               Left            =   4665
               Top             =   1140
               Width           =   3045
            End
            Begin VB.Shape ShapeB1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   3
               Left            =   830
               Top             =   1140
               Width           =   3045
            End
            Begin VB.Shape ShapeB2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   2
               Left            =   4665
               Top             =   780
               Width           =   3045
            End
            Begin VB.Shape ShapeB1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   2
               Left            =   830
               Top             =   780
               Width           =   3045
            End
            Begin VB.Shape ShapeB1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   1
               Left            =   825
               Top             =   420
               Width           =   3045
            End
            Begin VB.Shape ShapeB2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   1
               Left            =   4660
               Top             =   420
               Width           =   3045
            End
            Begin VB.Shape ShapeB2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   0
               Left            =   4670
               Top             =   60
               Width           =   3045
            End
            Begin VB.Shape ShapeB1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   0
               Left            =   830
               Top             =   60
               Width           =   3045
            End
            Begin VB.Label lblVariableValue 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   3
               Left            =   4080
               TabIndex        =   43
               Top             =   1150
               Width           =   615
            End
            Begin VB.Label lblVariableName 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   42
               Top             =   1150
               Width           =   615
            End
            Begin VB.Label lblVariableValue 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   0
               Left            =   4080
               TabIndex        =   34
               Top             =   70
               Width           =   615
            End
            Begin VB.Label lblVariableValue 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   1
               Left            =   4080
               TabIndex        =   33
               Top             =   430
               Width           =   615
            End
            Begin VB.Label lblVariableValue 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   2
               Left            =   4080
               TabIndex        =   32
               Top             =   790
               Width           =   615
            End
            Begin VB.Label lblVariableName 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   70
               Width           =   615
            End
            Begin VB.Label lblVariableName 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   30
               Top             =   430
               Width           =   615
            End
            Begin VB.Label lblVariableName 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   790
               Width           =   615
            End
         End
      End
      Begin VB.Label cmdMoreVariables 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "More"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   8280
         TabIndex        =   50
         Top             =   195
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   8280
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FF00&
         Height          =   1620
         Left            =   0
         Top             =   0
         Width           =   9495
      End
   End
   Begin VB.TextBox txtResponse 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5520
      Width           =   7335
   End
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      TabIndex        =   18
      ToolTipText     =   "Enter the URL here to the resource you want to request."
      Top             =   120
      Width           =   5415
   End
   Begin VB.ComboBox cboRequestMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmGetPost.frx":0000
      Left            =   7920
      List            =   "frmGetPost.frx":000A
      TabIndex        =   17
      Top             =   80
      Width           =   1695
   End
   Begin VB.TextBox txtRequest 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Frame fmeHeaders 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Additional Headers"
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "Use this space to add some custom HTTP headers of your own."
      Top             =   2205
      Width           =   9495
      Begin VB.PictureBox pbxOHeaders 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   10
         ScaleHeight     =   1455
         ScaleWidth      =   8175
         TabIndex        =   1
         Top             =   10
         Width           =   8175
         Begin VB.PictureBox pbxHeaders 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1470
            Left            =   0
            ScaleHeight     =   1470
            ScaleWidth      =   7815
            TabIndex        =   3
            Top             =   0
            Width           =   7815
            Begin VB.TextBox txtHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   3
               Left            =   840
               TabIndex        =   45
               Top             =   1060
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   3
               Left            =   4680
               TabIndex        =   44
               Top             =   1060
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   2
               Left            =   840
               TabIndex        =   9
               Top             =   730
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   8
               Top             =   400
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   7
               Top             =   70
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   2
               Left            =   4680
               TabIndex        =   6
               Top             =   730
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   1
               Left            =   4680
               TabIndex        =   5
               Top             =   400
               Width           =   3015
            End
            Begin VB.TextBox txtHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00181818&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Index           =   0
               Left            =   4680
               TabIndex        =   4
               Top             =   70
               Width           =   3015
            End
            Begin VB.Shape ShapeA2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   3
               Left            =   4660
               Top             =   1050
               Width           =   3045
            End
            Begin VB.Shape ShapeA1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   3
               Left            =   830
               Top             =   1050
               Width           =   3045
            End
            Begin VB.Shape ShapeA2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   2
               Left            =   4670
               Top             =   720
               Width           =   3045
            End
            Begin VB.Shape ShapeA1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   2
               Left            =   830
               Top             =   720
               Width           =   3045
            End
            Begin VB.Shape ShapeA1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   1
               Left            =   830
               Top             =   390
               Width           =   3045
            End
            Begin VB.Shape ShapeA2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   1
               Left            =   4670
               Top             =   390
               Width           =   3045
            End
            Begin VB.Shape ShapeA2 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   0
               Left            =   4670
               Top             =   60
               Width           =   3045
            End
            Begin VB.Shape ShapeA1 
               BorderColor     =   &H0000FF00&
               Height          =   315
               Index           =   0
               Left            =   830
               Top             =   60
               Width           =   3045
            End
            Begin VB.Label lblHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   47
               Top             =   1060
               Width           =   495
            End
            Begin VB.Label lblHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   3
               Left            =   4080
               TabIndex        =   46
               Top             =   1065
               Width           =   615
            End
            Begin VB.Label lblHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   730
               Width           =   495
            End
            Begin VB.Label lblHeaderName 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Top             =   400
               Width           =   615
            End
            Begin VB.Label lblHeaderName 
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   70
               Width           =   615
            End
            Begin VB.Label lblHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   2
               Left            =   4080
               TabIndex        =   12
               Top             =   735
               Width           =   615
            End
            Begin VB.Label lblHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   1
               Left            =   4080
               TabIndex        =   11
               Top             =   405
               Width           =   615
            End
            Begin VB.Label lblHeaderValue 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Value"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   0
               Left            =   4080
               TabIndex        =   10
               Top             =   75
               Width           =   615
            End
         End
         Begin VB.VScrollBar vsbHeaders 
            Enabled         =   0   'False
            Height          =   1290
            Left            =   7920
            Max             =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   80
            Width           =   255
         End
      End
      Begin VB.Label cmdMoreHeaders 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "More"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   8280
         TabIndex        =   49
         Top             =   195
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Index           =   0
         Left            =   8280
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         Height          =   1500
         Left            =   0
         Top             =   0
         Width           =   9495
      End
   End
   Begin MSWinsockLib.Winsock winsock 
      Left            =   8760
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   1605
      Index           =   1
      Left            =   950
      Top             =   5500
      Width           =   7365
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   1600
      Index           =   0
      Left            =   950
      Top             =   3830
      Width           =   7365
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   950
      Top             =   110
      Width           =   5440
   End
   Begin VB.Label cmdSend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8400
      TabIndex        =   51
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   8400
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   9735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8400
      TabIndex        =   48
      Top             =   6795
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   8400
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP Reponse:"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label lblUrl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblRequestMethod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Request method:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   37
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP Request:"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "frmGetPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' we set this to true whil a connection is established
Private blnConnected As Boolean

' sends HTTP req
Private Sub cmdSend_Click()
    Dim eUrl As URL
    
    Dim strMethod As String
    Dim strData As String
    Dim strPostData As String
    Dim strHeaders As String
    
    Dim strHTTP As String
    Dim x As Integer
    
    strPostData = ""
    strHeaders = ""
    strMethod = cboRequestMethod.List(cboRequestMethod.ListIndex)
    
    If blnConnected Then Exit Sub
    
    ' get url
    eUrl = ExtractUrl(txtUrl.Text)
    
    If eUrl.Host = vbNullString Then
        MsgBox "Invalid Host", vbCritical, "ERROR"
    
        Exit Sub
    End If
    
    ' cfg winsock
    winsock.Protocol = sckTCPProtocol
    winsock.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            winsock.RemotePort = eUrl.Port
        Else
            winsock.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        winsock.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If

    strData = ""
    For x = 0 To txtVariableName.Count - 1
        If txtVariableName(x).Text <> vbNullString Then
        
            strData = strData & URLEncode(txtVariableName(x).Text) & "=" & _
                            URLEncode(txtVariableValue(x).Text) & "&"
        End If
    Next x
    
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If

    If strData <> vbNullString Then
        strData = Left(strData, Len(strData) - 1)
        
        
        If strMethod = "GET" Then

            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
        Else

            strPostData = strData
            strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                         "Content-Length: " & Len(strPostData) & vbCrLf
                         
        End If
    End If

    For x = 0 To txtHeaderName.Count - 1
        If txtHeaderName(x).Text <> vbNullString Then
        
            strHeaders = strHeaders & txtHeaderName(x).Text & ": " & _
                            txtHeaderValue(x).Text & vbCrLf
        End If
    Next x

    txtResponse.Text = ""

    strHTTP = strMethod & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData

    txtRequest.Text = strHTTP
    
    winsock.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send HTTP req
    winsock.SendData strHTTP
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub winsock_Connect()
    blnConnected = True
End Sub

Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String

    winsock.GetData strResponse, vbString, bytesTotal
    
    strResponse = FormatLineEndings(strResponse)

    txtResponse.Text = txtResponse.Text & strResponse
    
End Sub

Private Sub winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "ERROR"
    
    winsock.Close
End Sub

Private Sub winsock_Close()
    blnConnected = False
    
    winsock.Close
End Sub

Private Function FormatLineEndings(ByVal str As String) As String
    Dim prevChar As String
    Dim nextChar As String
    Dim curChar As String
    
    Dim strRet As String
    
    Dim x As Long
    
    prevChar = ""
    nextChar = ""
    curChar = ""
    strRet = ""
    
    For x = 1 To Len(str)
        prevChar = curChar
        curChar = Mid$(str, x, 1)
                
        If nextChar <> vbNullString And curChar <> nextChar Then
            curChar = curChar & nextChar
            nextChar = ""
        ElseIf curChar = vbLf Then
            If prevChar <> vbCr Then
                curChar = vbCrLf
            End If
            
            nextChar = ""
        ElseIf curChar = vbCr Then
            nextChar = vbLf
        End If
        
        strRet = strRet & curChar
    Next x
    
    FormatLineEndings = strRet
End Function

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
    cboRequestMethod.ListIndex = 0
    blnConnected = False
End Sub

Private Sub cmdMoreHeaders_Click()
    Dim intNext As Integer
    Dim lngTop As Long
    
    ' find next control
    intNext = txtHeaderName.Count
    
    ' find next top
    lngTop = txtHeaderName(intNext - 1).Top + txtHeaderName(intNext - 1).Height + 70
    
    ' add new controls
    Load lblHeaderName(intNext)
    Load txtHeaderName(intNext)
    Load lblHeaderValue(intNext)
    Load txtHeaderValue(intNext)
    
    With lblHeaderName(intNext)
        .Top = lngTop
        .Left = lblHeaderName(intNext - 1).Left
        .Visible = True
    End With
    
    With txtHeaderName(intNext)
        .Top = lngTop
        .Left = txtHeaderName(intNext - 1).Left
        .Visible = True
        .Text = ""
    End With
        
    With lblHeaderValue(intNext)
        .Top = lngTop
        .Left = lblHeaderValue(intNext - 1).Left
        .Visible = True
    End With
    
    With txtHeaderValue(intNext)
        .Top = lngTop
        .Left = txtHeaderValue(intNext - 1).Left
        .Visible = True
        .Text = ""
    End With
    
    
    '================================ZzZzZzZz==============================
    Load ShapeA2(intNext)
    Load ShapeA1(intNext)
    
    With ShapeA1(intNext)
        .Top = lngTop - 15 'ShapeA1(intNext - 1).Top + ShapeA1(intNext - 1).Height + 15
        .Left = txtHeaderName(intNext - 1).Left - 15 'ShapeA1(0).Left
        .Visible = True
    End With
    With ShapeA2(intNext)
        .Top = lngTop - 15 'ShapeA2(intNext - 1).Top + ShapeA2(intNext - 1).Height + 15
        .Left = txtHeaderValue(intNext - 1).Left - 15 'ShapeA2(0).Left
        .Visible = True
    End With
    '======================================================================
    
    
    ' set the new height of the controls
    pbxHeaders.Height = txtHeaderName(intNext).Top + txtHeaderName(intNext).Height + 80
    
    If pbxHeaders.Height > pbxOHeaders.Height Then
        With vsbHeaders
            .Enabled = True
            .SmallChange = txtHeaderName(intNext).Height
            .LargeChange = pbxOHeaders.Height
            .Min = 0
            .Max = pbxHeaders.Height - pbxOHeaders.Height
            .value = .Max
        End With
    End If
End Sub

Private Sub cmdMoreVariables_Click()
    Dim intNext As Integer
    Dim lngTop As Long
    
    ' find next control
    intNext = txtVariableName.Count
    
    ' find the next top
    lngTop = txtVariableName(intNext - 1).Top + txtVariableName(intNext - 1).Height + 80
    
    ' add new controls
    Load lblVariableName(intNext)
    Load txtVariableName(intNext)
    Load lblVariableValue(intNext)
    Load txtVariableValue(intNext)
    
    With lblVariableName(intNext)
        .Top = lngTop
        .Left = lblVariableName(intNext - 1).Left
        .Visible = True
    End With
    
    With txtVariableName(intNext)
        .Top = lngTop
        .Left = txtVariableName(intNext - 1).Left
        .Visible = True
        .TabIndex = txtVariableName(intNext - 1).TabIndex + 2
        .Text = ""
    End With
        
    With lblVariableValue(intNext)
        .Top = lngTop
        .Left = lblVariableValue(intNext - 1).Left
        .Visible = True
    End With
    
    With txtVariableValue(intNext)
        .Top = lngTop
        .Left = txtVariableValue(intNext - 1).Left
        .TabIndex = txtVariableValue(intNext - 1).TabIndex + 2
        .Visible = True
        .Text = ""
    End With
    
       '================================ZzZzZzZz==============================
    Load ShapeB1(intNext)
    Load ShapeB2(intNext)
    
    With ShapeB1(intNext)
        .Top = lngTop - 15 'ShapeA1(intNext - 1).Top + ShapeA1(intNext - 1).Height + 15
        .Left = txtVariableName(intNext - 1).Left - 15 'ShapeA1(0).Left
        .Visible = True
    End With
    With ShapeB2(intNext)
        .Top = lngTop - 15 'ShapeA2(intNext - 1).Top + ShapeA2(intNext - 1).Height + 15
        .Left = txtVariableValue(intNext - 1).Left - 15 'ShapeA2(0).Left
        .Visible = True
    End With
    '======================================================================
 
    
    pbxVariables.Height = txtVariableName(intNext).Top + txtVariableName(intNext).Height + 80
    
    If pbxVariables.Height > pbxOVariables.Height Then
        With vsbVariables
            .Enabled = True
            .SmallChange = txtVariableName(intNext).Height
            .LargeChange = pbxOVariables.Height
            .Min = 0
            .Max = pbxVariables.Height - pbxOVariables.Height
            .value = .Max
        End With
    End If
End Sub

Private Sub vsbHeaders_Change()
    pbxHeaders.Top = 0 - vsbHeaders.value
End Sub

Private Sub vsbHeaders_Scroll()
    pbxHeaders.Top = 0 - vsbHeaders.value
End Sub

Private Sub vsbVariables_Change()
    pbxVariables.Top = 0 - vsbVariables.value
End Sub

Private Sub vsbVariables_Scroll()
    pbxVariables.Top = 0 - vsbVariables.value
End Sub

