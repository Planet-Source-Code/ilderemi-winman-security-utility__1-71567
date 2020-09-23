VERSION 5.00
Begin VB.Form frmAttrib 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
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
   ScaleHeight     =   1095
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Script"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton optS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CMD"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox optAttrib 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hidden"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox optAttrib 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Archive"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox optAttrib 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "System"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox optAttrib 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Read-Only"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtAttrib 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   3960
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   150
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   3960
      Top             =   105
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   945
      Top             =   105
      Width           =   2925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   140
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   1095
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmAttrib"
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
Dim strAttrib As String
Dim numAttrib As Integer

Private Sub cmdOK_Click()
On Error Resume Next
numAttrib = 0
'======================AttribPars=====================
For j = 0 To 3
    If optS(0).value = True Then
        Select Case optAttrib(j).value
            Case 1
            Select Case optAttrib(j).Index
                Case 0
                strAttrib = strAttrib & "+R "
                Case 1
                strAttrib = strAttrib & "+S "
                Case 2
                strAttrib = strAttrib & "+A "
                Case 3
                strAttrib = strAttrib & "+H "
            End Select
            
            Case 0
            Select Case optAttrib(j).Index
                Case 0
                strAttrib = strAttrib & "-R "
                Case 1
                strAttrib = strAttrib & "-S "
                Case 2
                strAttrib = strAttrib & "-A "
                Case 3
                strAttrib = strAttrib & "-H "
            End Select
        End Select
    Else: optS(1).value = True
        Select Case optAttrib(j).value
            Case 1 'CHECKED
            Select Case optAttrib(j).Index
                Case 0
                numAttrib = numAttrib Or 1  '+000001 ReadOnly
                Case 1
                numAttrib = numAttrib Or 4  '+000100 System
                Case 2
                numAttrib = numAttrib Or 32 '+100000 Archive
                Case 3
                numAttrib = numAttrib Or 2  '+000010 Hidden
            End Select
            Dim FS, F
            Set FS = CreateObject("Scripting.FileSystemObject")
            Set F = FS.GetFile(txtAttrib.Text)
            F.Attributes = numAttrib
            Set F = FS.GetFolder(txtAttrib.Text)
            F.Attributes = numAttrib
        End Select
    End If
Next
'=====================================================
If optS(0).value = True Then Shell "attrib " & strAttrib & Chr(34) & txtAttrib & Chr(34), vbHide
strAttrib = ""
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
