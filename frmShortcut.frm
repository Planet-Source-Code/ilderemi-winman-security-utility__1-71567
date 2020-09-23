VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmShortcut 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ÓÇÎÊ ãíÇäÈÑ"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
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
   ScaleHeight     =   2100
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   6975
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmShortcut.frx":0000
      Left            =   4340
      List            =   "frmShortcut.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   830
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmShortcut.frx":013F
      Left            =   4340
      List            =   "frmShortcut.frx":015B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   230
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   120
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   225
      Top             =   1185
      Width           =   7000
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   2100
      Left            =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label isButton22 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6010
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   885
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   6010
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   855
      Left            =   120
      Top             =   720
      Width           =   7215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   320
      Left            =   3120
      Top             =   830
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   1070
      Top             =   830
      Width           =   2085
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   860
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   860
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   3130
      TabIndex        =   8
      Top             =   860
      Width           =   375
   End
   Begin VB.Label isButton21 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6010
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   270
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   6010
      Top             =   225
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   3120
      Top             =   230
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   1070
      Top             =   230
      Width           =   2080
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   255
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   260
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   ">>"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   3130
      TabIndex        =   2
      Top             =   270
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   530
      Left            =   120
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmShortcut"
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
Private Sub Combo2_Click()
Dim ext As String
Select Case LCase(Combo2.Text)
    Case "my computer"
    ext = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    
    Case "history"
    ext = "{FF393560-C2A7-11CF-BFF4-444553540000}"
    
    Case "recycle bin"
    ext = "{645FF040-5081-101B-9F08-00AA002F954E}"
    
    Case "start menu"
    ext = "{48e7caab-b918-4e58-a94d-505519c795dc}"
    
    Case "folder(internet)"
    ext = "{3DC7A020-0ACD-11CF-A9BB-00AA004AE837}"
    
    Case "folder(application)"
    ext = "{3050f4d8-98B5-11CF-BB82-00AA00BDCE0B}"
    
    Case "set program access and defaults"
    ext = "{2559a1f7-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "e-mail"
    ext = "{2559a1f5-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "internet"
    ext = "{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "run"
    ext = "{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "windows security"
    ext = "{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "help"
    ext = "{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "search"
    ext = "{2559a1f0-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Case "printers and faxes"
    ext = "{2227A280-3AEA-1069-A2DE-08002B30309D}"
    
    Case "control panel"
    ext = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    
    Case "my network places"
    ext = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    
    Case "computer search results folder"
    ext = "{1f4de370-d627-11d1-ba4f-00a0c91eedba}"
    
    Case ""
    ext = ""
    
    Case ""
    ext = ""
End Select
Text8.Text = ext
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo2.ListIndex = 0
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

Private Sub isButton21_Click()
On Error Resume Next
Open Text5 & "\" & Combo1.Text & ".lnk" For Output As #1
Print #1, Chr(76) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(20) & Chr(2) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(192) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(70) & Chr(129) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(22) & Chr(0) & Chr(20) & Chr(0) & Chr(31) & Chr(128) & Chr(240 + Combo1.ListIndex) & Chr(161) & Chr(89) & Chr(37) & Chr(215) & Chr(33) & Chr(212) & Chr(17) & Chr(189) & Chr(175) & Chr(0) & Chr(192) & Chr(79) & Chr(96) & Chr(185) & Chr(240) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
Close #1
End Sub

Private Sub isButton22_Click()
On Error Resume Next
MkDir Text6 & "." & Text8.Text
End Sub

Private Sub Label1_Click()
CommonDialog1.ShowSave
Text6.Text = CommonDialog1.Filename
End Sub

Private Sub Label2_Click()
CommonDialog1.ShowSave
Text6.Text = CommonDialog1.Filename
End Sub

Private Sub Label7_Click()
Unload Me
End Sub
