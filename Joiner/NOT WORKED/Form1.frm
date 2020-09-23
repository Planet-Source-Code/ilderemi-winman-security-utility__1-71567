VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
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
   ScaleHeight     =   7215
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   6735
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   6840
      Width           =   3000
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompany"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblPhone"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblName"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblWebsite 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblWebsite"
      ForeColor       =   &H00A47733&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblMailto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblMailto"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
lblWebsite.FontUnderline = True
Dim s As String
LoadMyFile (101)

    Me.Caption = Right(LoadCommand(1), Len(LoadCommand(1)) - Len("_title::"))
    Text1.Text = Right(LoadCommand(2), Len(LoadCommand(2)) - Len("_body::"))
    'Me.Caption = Right(LoadCommand(3), Len(LoadCommand(3)) - Len("_img::"))
    If Right(LoadCommand(4), Len(LoadCommand(4)) - Len("_msg::")) <> "" Then MsgBox Right(LoadCommand(4), Len(LoadCommand(4)) - Len("_msg::")), , Me.Caption
    'Me.Left = Val(Right(LoadCommand(5), Len(LoadCommand(5)) - Len("_frmLocationX::")))
    'Me.Top = Val(Right(LoadCommand(6), Len(LoadCommand(6)) - Len("_frmLocationY::")))
    lblMailto.Caption = Right(LoadCommand(7), Len(LoadCommand(7)) - Len("_mailto::"))
    lblWebsite.Caption = Right(LoadCommand(8), Len(LoadCommand(8)) - Len("_website::"))
    lblName.Caption = Right(LoadCommand(9), Len(LoadCommand(9)) - Len("_name::"))
    lblPhone.Caption = Right(LoadCommand(10), Len(LoadCommand(10)) - Len("_phone::"))
    lblCompany.Caption = Right(LoadCommand(11), Len(LoadCommand(11)) - Len("_company::"))

For i = 0 To LenB(LoadResData(101, "txt")) - 1
    s = s & Chr(MyFile(i))
Next
'Text1 = s

'================================
'Open "c:\a.txt" For Binary As #1
'    Put #1, 1, MyFile
'Close
'================================
End Sub
