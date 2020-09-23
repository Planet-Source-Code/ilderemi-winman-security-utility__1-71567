VERSION 5.00
Begin VB.Form frmCodeAscii 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEnText 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2880
      Width           =   6375
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   3720
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5295
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   5160
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5295
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5295
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2325
      Index           =   1
      Left            =   105
      Top             =   105
      Width           =   6405
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2325
      Index           =   0
      Left            =   105
      Top             =   2865
      Width           =   6405
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ASCII Code:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ASCII Char:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Key Code:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   6375
   End
End
Attribute VB_Name = "frmCodeAscii"
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

Private Sub Label9_Click()
Unload Me
End Sub

Private Sub txtEnText_Change()
Label2.Caption = Len(txtEnText)
End Sub

Private Sub txtText_Change()
Label1.Caption = Len(txtText)
If txtText = "" Then txtEnText = ""
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
Label4 = KeyCode
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
'For k = 1 To 26
    'If KeyAscii = k Then Exit Sub
'Next
Label7 = KeyAscii
On Error Resume Next
Label6 = Chr(KeyAscii)
If KeyAscii = 8 Then
    For i = 1 To 12
        If Mid(txtEnText, Len(txtEnText) - i, 1) = "&" Then
            txtEnText = Left(txtEnText, Len(txtEnText) - i - 2)
            GoTo 1
        End If
    Next
End If
If txtEnText.Text = "" Then
    txtEnText.Text = "chr$(" & KeyAscii & ")"
Else
    txtEnText.Text = txtEnText.Text & " & " & "chr$(" & KeyAscii & ")"
End If
1:
End Sub
