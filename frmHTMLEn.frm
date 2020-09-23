VERSION 5.00
Begin VB.Form frmHTMLEn 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtEnc 
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
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2880
      Width           =   6855
   End
   Begin VB.TextBox TxtDec 
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
      ForeColor       =   &H00008000&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   4680
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5720
      Width           =   2295
   End
   Begin VB.Label isButton3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   5715
      Width           =   1095
   End
   Begin VB.Label isButton4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   5715
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   3480
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   2400
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Decode"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   5715
      Width           =   1095
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encode"
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
      TabIndex        =   2
      Top             =   5720
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   1200
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   120
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   2680
      Left            =   110
      Top             =   110
      Width           =   6880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2680
      Left            =   110
      Top             =   2870
      Width           =   6880
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmHTMLEn"
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
        Dim i, Dec
        For i = 1 To Len(TxtDec.Text)
            If Len(Hex(Asc(Mid(TxtDec.Text, i, 1)))) = 2 Then Dec = Dec & "%" & Hex(Asc(Mid(TxtDec.Text, i, 1)))
            If Len(Hex(Asc(Mid(TxtDec.Text, i, 1)))) = 1 Then Dec = Dec & "%0" & Hex(Asc(Mid(TxtDec.Text, i, 1)))
        Next
        TxtEnc.Text = "<SCRIPT LANGUAGE=" & Chr(34) & "JavaScript" & Chr(34) & ">document.write(unescape(" & Chr(34) & Dec & Chr(34) & "));</SCRIPT>"
End Sub

Private Sub isButton2_Click()
        Dim j, i As Integer
        Dim Enc As String
        On Error GoTo 2
        If Mid(TxtDec.Text, Len(TxtDec.Text), 1) <> "%" Then TxtDec.Text = TxtDec.Text & "%"
        For i = 1 To Len(TxtDec.Text)
            If Mid(TxtDec.Text, i, 1) = "%" Then
                For j = 1 To 4
                    If Mid(TxtDec.Text, i + j, 1) = "%" Then
                        Enc = Enc & Chr(Val("&H" & Mid(TxtDec.Text, i + 1, j - 1)))
                        GoTo 1
                    End If
                Next
            End If
1:
        Next
        TxtEnc.Text = Enc
        Exit Sub
2:
        TxtEnc.Text = "{Error : ASCII Code is Larger than &HFF or not defined}"
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
