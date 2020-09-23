VERSION 5.00
Begin VB.Form frmCMD 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
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
   ScaleHeight     =   4215
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Trace Route"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   7560
      MaskColor       =   &H0000FF00&
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "MAC"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   1
      Left            =   7560
      MaskColor       =   &H0000FF00&
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CMD"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   7560
      MaskColor       =   &H0000FF00&
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame framCMD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   15
      Width           =   7215
      Begin VB.TextBox txtCMDCom 
         Appearance      =   0  'Flat
         BackColor       =   &H00181818&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1005
         TabIndex        =   8
         Top             =   240
         Width           =   4400
      End
      Begin VB.TextBox txtCMD 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   600
         Width           =   6975
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0000FF00&
         Height          =   3975
         Left            =   0
         Top             =   120
         Width           =   7215
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0000FF00&
         Height          =   3400
         Left            =   110
         Top             =   590
         Width           =   7000
      End
      Begin VB.Label isButton2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   270
         Width           =   1725
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FF00&
         Height          =   315
         Left            =   5400
         Top             =   225
         Width           =   1725
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FF00&
         Height          =   315
         Left            =   990
         Top             =   225
         Width           =   4420
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Command:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frmMAC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   15
      Width           =   7215
      Begin VB.TextBox txtmac 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox txtcname 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00181818&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   4800
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FF00&
         Height          =   3400
         Left            =   110
         Top             =   590
         Width           =   7000
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         Height          =   3975
         Left            =   0
         Top             =   120
         Width           =   7215
      End
      Begin VB.Label isButton1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   270
         Width           =   1720
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   5400
         Top             =   230
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   585
         Top             =   225
         Width           =   4830
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label isButton3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   3795
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   7560
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   7440
      X2              =   7440
      Y1              =   0
      Y2              =   4200
   End
End
Attribute VB_Name = "frmCMD"
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
On Error GoTo mace
If (txtcname.Text = "") Then
MsgBox "Enter The IP or Computer Name", vbInformation
txtcname.SetFocus
End If
Select Case Me.Caption
    Case "MAC"
    txtmac.Text = getm.GetCommandOutput("nbtstat -a " & txtcname.Text, True, False, True)
    
    Case "Trace Route"
    txtmac.Text = getm.GetCommandOutput("tracert " & txtcname.Text, True, False, True)
End Select
Exit Sub
mace:
MsgBox Err.Description, vbCritical
End Sub

Private Sub isButton2_Click()
txtCMD.Text = ""
txtCMD.Text = getm.GetCommandOutput("cmd /c " & txtCMDCom, True, True, True)
End Sub

Private Sub isButton3_Click()
Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
frmMAC.Visible = False
framCMD.Visible = False
'frmUser.Visible = False
Select Case Option1(Index).Caption
    
    Case "CMD"
    framCMD.Visible = True
    
    Case "MAC"
    frmMAC.Visible = True
    txtmac.Text = ""
    txtCMDCom.Text = ""
    
    Case "Trace Route"
    frmMAC.Visible = True
    txtmac.Text = ""
    txtCMDCom.Text = ""
    
End Select
Me.Caption = Option1(Index).Caption
End Sub

Private Sub txtCMDCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then isButton2_Click
End Sub
