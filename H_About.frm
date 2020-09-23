VERSION 5.00
Begin VB.Form H_About 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ÏÑÈÇÑå"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      Picture         =   "H_About.frx":0000
      RightToLeft     =   -1  'True
      ScaleHeight     =   2505
      ScaleWidth      =   2985
      TabIndex        =   3
      Top             =   360
      Width           =   3015
      Begin VB.Timer Timer3 
         Interval        =   50
         Left            =   1320
         Top             =   120
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3360
      Top             =   960
   End
   Begin VB.Label DMSXpButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4035
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   120
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   1560
      Picture         =   "H_About.frx":63A02
      Top             =   3240
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   1455
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000FF00&
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   7580
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ØÑÇÍí ÈÑäÇãå åÇí ÇãäíÊí¡  ÓíÓÊã åÇí ãÏíÑíÊí¡ ˜äÊÑá åæÔãäÏ¡ ÔÈ˜å æ ..."
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
      Left            =   -5520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   375
      Left            =   8
      Top             =   15
      Width           =   7565
   End
End
Attribute VB_Name = "H_About"
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
Dim mp As Integer
Dim io As Integer
Private Sub DMSXpButton1_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Unload Me
Unload Form1
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer2.Enabled = True
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

Private Sub Timer1_Timer()
If Label3.Left > 7560 Then
    Label3.Left = -5520
Else
    Label3.Left = Label3.Left + 10
End If
End Sub

Private Sub Timer2_Timer()
io = io + 1
If io < 300 Then
    ort = "This is an open source software in utility and security category. For more information:" & vbCrLf & "mailderemi@gmail.com - info@ilderemi.com - +989151092841"
    Label2.Caption = Left(ort, io)
Else
io = 0
End If
End Sub

Private Sub Timer3_Timer()
mp = mp + 1
If mp = 1 Then
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Picture.Width, Picture1.Picture.Height, 0, 2550, Picture1.Picture.Width, Picture1.Picture.Height
ElseIf mp = 2 Then
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Picture.Width, Picture1.Picture.Height, 0, 2550 * 2, Picture1.Picture.Width, Picture1.Picture.Height
ElseIf mp = 3 Then
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Picture.Width, Picture1.Picture.Height, 0, 2550 * 3, Picture1.Picture.Width, Picture1.Picture.Height
ElseIf mp = 4 Then
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Picture.Width, Picture1.Picture.Height, 0, 2500 * 4, Picture1.Picture.Width, Picture1.Picture.Height
    mp = 0
End If
End Sub
