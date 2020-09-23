VERSION 5.00
Begin VB.Form frmHash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4125
   ClientLeft      =   525
   ClientTop       =   180
   ClientWidth     =   6375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   3525
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHA-1024"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2300
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   120
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHA-512"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1940
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   120
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   4120
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   120
      Top             =   3760
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   3555
      Index           =   1
      Left            =   1305
      Top             =   465
      Width           =   4965
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   1310
      Top             =   110
      Width           =   4960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   120
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   120
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHA-384"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1580
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHA-256"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1220
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHA-160"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   860
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MD5"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   500
      Width           =   1095
   End
End
Attribute VB_Name = "frmHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MD5 As clsMD5

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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Label1_Click()
    Set MD5 = New clsMD5
    Text2 = MD5.CalculateMD5(Text1)
End Sub

Private Sub Label2_Click()
    Text2 = Sha1(Text1)
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label6_Click()
Unload Me
End Sub

Private Sub Label7_Click()
    Text2 = SHA512(Text1)
End Sub

Private Sub Label8_Click()
    Text2 = SHA1024(Text1)
End Sub
