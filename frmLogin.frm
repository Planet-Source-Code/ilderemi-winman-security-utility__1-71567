VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ê—Êœ"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   RightToLeft     =   -1  'True
   ScaleHeight     =   1215
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   1200
      Top             =   840
      Width           =   975
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   860
      Width           =   975
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Height          =   255
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   860
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
MyLang = Combo1.Text
Call ChLang(MyLang)
End Sub

Private Sub Form_Load()
Combo1 = "English"
MakeFlat Text1.hwnd
MakeFlat Text2.hwnd

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

Private Sub Form_Terminate()
End
End Sub

Private Sub isButton1_Click()
Unload Me
Unload frmScrSvr
End Sub

Private Sub isButton2_Click()
If Text1 = "iLDEREMi" And Text2 = "2020176912136" Then
frmMain.Show
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Timer1.Enabled = False
Me.Hide
Else
MsgBox "Username or Password is not valid."
Text1 = ""
Text2 = ""
End If
End Sub

Private Sub ChLang(Language As String)
Select Case Language
    
    Case "English"
    
    Case "Å«—”Ì"

End Select
End Sub
