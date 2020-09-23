VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMkCab 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
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
   ScaleHeight     =   1335
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2610
      TabIndex        =   7
      Top             =   915
      Width           =   1500
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1070
      TabIndex        =   6
      Top             =   920
      Width           =   1495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   2610
      Top             =   840
      Width           =   1495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   1070
      Top             =   840
      Width           =   1495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   3730
      Top             =   470
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   3735
      Top             =   105
      Width           =   380
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   1070
      Top             =   470
      Width           =   2685
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   1070
      Top             =   110
      Width           =   2680
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Source File:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   510
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   135
      Width           =   255
   End
End
Attribute VB_Name = "frmMkCab"
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

Private Sub isButton1_Click()
Shell "Makecab " & Chr(34) & Text1 & Chr(34) & " " & Chr(34) & Text2 & Chr(34) & " /V1"
End Sub

Private Sub Label1_Click()
CommonDialog1.ShowOpen
Text1 = CommonDialog1.Filename
End Sub

Private Sub Label2_Click()
CommonDialog1.ShowSave
Text2 = CommonDialog1.Filename
End Sub

Private Sub Label6_Click()
Unload Me
End Sub
