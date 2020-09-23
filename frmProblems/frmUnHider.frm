VERSION 5.00
Begin VB.Form frmUnHider 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "iLDEREMi Unhider"
   ClientHeight    =   6615
   ClientLeft      =   1095
   ClientTop       =   165
   ClientWidth     =   8535
   Icon            =   "frmUnHider.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
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
      Height          =   6075
      Left            =   120
      TabIndex        =   4
      Top             =   435
      Width           =   8325
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
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
      Left            =   2280
      TabIndex        =   0
      Top             =   70
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   6120
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6135
      TabIndex        =   5
      Top             =   105
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7335
      TabIndex        =   3
      Top             =   105
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   7320
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unhide All"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   4905
      Top             =   60
      Width           =   1125
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   6620
      Left            =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   6100
      Left            =   105
      Top             =   420
      Width           =   8350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   2270
      Top             =   60
      Width           =   2560
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Enter your Drive or Folder:"
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
      TabIndex        =   1
      Top             =   100
      Width           =   2055
   End
End
Attribute VB_Name = "frmUnHider"
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
Private Sub Command1_Click()
List1.Visible = False
Call search(Text1.Text, List1)
List1.Visible = True
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

Private Sub search(Path As String, ByRef lstBox As ListBox)
Dim result
Dim FsoObject, FolderName, Member, SubFolderName
Dim strFileList, intCounter

Set FsoObject = CreateObject("Scripting.FileSystemObject")
Set FolderName = FsoObject.GetFolder(Path)
intCounter = 0

For Each Member In FolderName.SubFolders
    intCounter = intCounter + 1
    
    FsoObject.GetFolder(FolderName.Path & "\" & Member.Name).Attributes = 0
    
    lstBox.AddItem FolderName.Path & "\" & Member.Name
    search FolderName.Path & "\" & Member.Name, lstBox
Next

For Each Member In FolderName.Files
    intCounter = intCounter + 1
    
    FsoObject.GetFile(FolderName.Path & "\" & Member.Name).Attributes = 0
    
    lstBox.AddItem FolderName.Path & "\" & Member.Name
Next
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
