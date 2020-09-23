VERSION 5.00
Begin VB.Form frmAntiSVCHOST 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "äÇã ˜ÇÑÈÑ æíäÏæÒ"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   975
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   830
      Top             =   110
      Width           =   1840
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   825
      TabIndex        =   2
      Top             =   620
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   825
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WinUser:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAntiSVCHOST"
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
On Error Resume Next
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im svchost.exe", vbHide
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im smss.exe", vbHide
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im csrss.exe", vbHide
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im winlogon.exe", vbHide
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im services.exe", vbHide
Shell "cmd /c taskkill /f /fi " & Chr(34) & "USERNAME eq " & Text1 & Chr(34) & " /im lsass.exe", vbHide
Unload Me
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
