VERSION 5.00
Begin VB.Form frmDir 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1110
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox Listsearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0080C0FF&
      Height          =   1590
      ItemData        =   "frmSearch.frx":0000
      Left            =   120
      List            =   "frmSearch.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   4785
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   510
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   3720
      Top             =   470
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   5030
   End
   Begin VB.Label cmdsearch 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   1090
      Top             =   470
      Width           =   2565
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   1100
      Top             =   110
      Width           =   2560
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   1620
      Left            =   100
      Top             =   830
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   3705
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   165
      Width           =   795
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dir"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3945
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmDir"
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
Private Sub cmdsearch_Click()
If txtname.Text <> "" And txtpath.Text <> "" Then
    Listsearch.clear
    'search a folder
    fileSearch txtpath.Text, txtname.Text
End If
End Sub
Private Sub fileSearch(PathName As String, Filename As String)
Dim rec As WIN32_FIND_DATA
Dim Mypath As String
Dim hResult As Long
Mypath = PathName
hResult = FindFirstFile(Mypath + Filename, rec)
If hResult <> INVALID_HANDLE_VALUE Then
    Do While FindNextFile(hResult, rec) = 1
        Listsearch.AddItem Mypath & rec.cFileName
    Loop
    FindClose (hResult)
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
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
