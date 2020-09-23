VERSION 5.00
Begin VB.Form frmOutMail 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
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
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox Text5 
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
      Height          =   2205
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00154D00&
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
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00103C00&
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
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H000B2800&
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
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00061700&
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   2575
      TabIndex        =   13
      Top             =   4270
      Width           =   1890
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   710
      TabIndex        =   12
      Top             =   4270
      Width           =   1885
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   2580
      Top             =   4200
      Width           =   1885
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   710
      Top             =   4200
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2235
      Index           =   5
      Left            =   705
      Top             =   1545
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   710
      Top             =   3830
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   710
      Top             =   1190
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   710
      Top             =   830
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   710
      Top             =   470
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   710
      Top             =   110
      Width           =   3760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "frmOutMail"
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
On Error Resume Next
Shell "cmd /c %SYSTEMDRIVE%\" & Chr(34) & "Program Files" & Chr(34) & "\" & Chr(34) & "Outlook Express" & Chr(34) & "\msimn.exe /mailurl:mailto:" & Text1 & "?Subject=" & Text2 & "&Cc=" & Text3 & "&Bcc=" & Text4 & "&Attach=" & Text6 & "&Body=" & Text5 & Chr(34)
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
