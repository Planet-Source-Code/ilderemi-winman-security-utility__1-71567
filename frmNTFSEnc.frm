VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNTFSEnc 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
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
   ScaleHeight     =   1215
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00191919&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00191919&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   710
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   710
      TabIndex        =   7
      Top             =   855
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   4200
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   855
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   705
      Top             =   465
      Width           =   4245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   705
      Top             =   105
      Width           =   4245
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4995
      TabIndex        =   5
      Top             =   500
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   4920
      Top             =   465
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   4995
      TabIndex        =   4
      Top             =   140
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   4920
      Top             =   105
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Folder:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmNTFSenc"
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

Private Sub Label2_Click()
On Error Resume Next
Shell "cipher /e /s:""" & Text2.text & """"
Shell "cipher /e /a """ & Text1.text & """"
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
CommonDialog1.ShowOpen
Text1.text = CommonDialog1.Filename
End Sub

