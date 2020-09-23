VERSION 5.00
Begin VB.Form frmWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3120
      Top             =   3000
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   4080
      Max             =   255
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3675
      Width           =   4455
   End
   Begin VB.TextBox txthWnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   240
      Width           =   2115
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   600
      Width           =   2115
   End
   Begin VB.TextBox txtClass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   960
      Width           =   2115
   End
   Begin VB.TextBox txtParent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2400
      Width           =   2115
   End
   Begin VB.TextBox txtStyle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   645
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   1320
      Width           =   2115
   End
   Begin VB.TextBox txtParentText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2760
      Width           =   2115
   End
   Begin VB.TextBox txtParentClass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3120
      Width           =   2115
   End
   Begin VB.TextBox txtRect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2040
      Width           =   2115
   End
   Begin VB.TextBox txtCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtCW 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2385
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2385
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Size"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Move"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   2385
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "DrawFrame"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   12
      Top             =   3165
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "HideWindow"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   11
      Top             =   2925
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "FrameChanged"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   2925
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "NoZOrder"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   9
      Top             =   2685
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "NoCopyBits"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   8
      Top             =   2685
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "NoActivate"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   7
      Top             =   2685
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ShowWindow"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   6
      Top             =   3165
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "NoRedraw"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   5
      Top             =   2925
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmWindow.frx":0000
      Left            =   3840
      List            =   "frmWindow.frx":0010
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3165
      Width           =   1455
   End
   Begin VB.TextBox txtWin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox picCrossHair 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      MouseIcon       =   "frmWindow.frx":0035
      Picture         =   "frmWindow.frx":033F
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Back to Main"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   3860
      Width           =   3495
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   120
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   7070
      Top             =   2370
      Width           =   765
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   5630
      Top             =   2370
      Width           =   765
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   7070
      Top             =   2030
      Width           =   765
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   5630
      Top             =   2030
      Width           =   760
   End
   Begin VB.Label isButton7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Go"
      ForeColor       =   &H0000FF00&
      Height          =   200
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   270
      Width           =   495
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   7680
      Top             =   225
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   5025
      Top             =   225
      Width           =   2565
   End
   Begin VB.Label isButton13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label isButton10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Disable"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label isButton9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Not On Top"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label isButton8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label isButton12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label isButton6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enable"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label isButton5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "On Top"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label isButton4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label isButton11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximize"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label isButton3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Flash"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Set Title"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   795
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   11
      Left            =   7680
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   10
      Left            =   7680
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   9
      Left            =   7680
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   6360
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   7
      Left            =   6360
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   6
      Left            =   6360
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   5040
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   5040
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   5040
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   3720
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   3720
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   3720
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label isButton14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7920
      TabIndex        =   42
      Top             =   2110
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   7920
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   8595
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3675
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   2175
      Left            =   3720
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   3375
      Left            =   120
      Top             =   130
      Width           =   3495
   End
   Begin VB.Label lblHwnd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "hWnd:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   300
      Width           =   675
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   660
      Width           =   555
   End
   Begin VB.Label lblClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   1020
      Width           =   555
   End
   Begin VB.Label lblParent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parent:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   35
      Top             =   2460
      Width           =   675
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label lblParentText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Text:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   33
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Class:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rectangle:"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   3
      Left            =   5115
      TabIndex        =   22
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   2
      Left            =   6390
      TabIndex        =   21
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   20
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   19
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label lblCordi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X: 1043  Y: 0032"
      ForeColor       =   &H005E7B26&
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   3600
      Width           =   3435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Image imgCursor 
      Height          =   375
      Left            =   8400
      MouseIcon       =   "frmWindow.frx":0F81
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmWindow"
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
Option Explicit

Dim strBuffer As String
Dim lngReturn As Long
Dim strWindowsDirectory As String
Dim winst As Long
' Dragging window?
Private m_bDragging As Boolean

Private Sub Check2_Click()
If Check2.value = 1 Then
    txtCH.Enabled = True
    txtCW.Enabled = True
    txtCH.BackColor = &H4000&
    txtCW.BackColor = &H4000&
    Label1(2).Enabled = True
    Label1(3).Enabled = True
Else
    txtCH.Enabled = False
    txtCW.Enabled = False
    txtCH.BackColor = &H404040
    txtCW.BackColor = &H404040
    Label1(2).Enabled = False
    Label1(3).Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.value = 1 Then
    txtX.Enabled = True
    txtY.Enabled = True
    txtX.BackColor = &H4000&
    txtY.BackColor = &H4000&
    Label1(0).Enabled = True
    Label1(1).Enabled = True
Else
    txtX.Enabled = False
    txtY.Enabled = False
    txtX.BackColor = &H404040
    txtY.BackColor = &H404040
    Label1(0).Enabled = False
    Label1(1).Enabled = False
End If
End Sub

Private Sub Combo1_Click()
Select Case Combo1.Text

    Case "TopMost"
    winst = HWND_TOPMOST
    
    Case "NoTopMost"
    winst = HWND_NOTOPMOST
    
    Case "Top"
    winst = HWND_TOP
    
    Case "Bottom"
    winst = HWND_BOTTOM
    
End Select
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
HScroll1.value = 255
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

Private Sub HScroll1_Change()
TranslucentForm HScroll1.value
End Sub

Private Sub isButton1_Click()
    ' Hide window
    ShowWindow txthWnd.Text, SW_HIDE
End Sub

Private Sub isButton10_Click()
    ' Disable window
    EnableWindow txthWnd.Text, 0
End Sub

Private Sub isButton11_Click()
    ' Maximize window
    ShowWindow txthWnd.Text, SW_MAXIMIZE
End Sub

Private Sub isButton12_Click()
    ' Minimize window
    ShowWindow txthWnd.Text, SW_MINIMIZE
End Sub

Private Sub isButton13_Click()
    ' Show window
    ShowWindow txthWnd.Text, SW_NORMAL
End Sub

Private Sub isButton14_Click()
Dim temp As Long
temp = 0
If Check1(0).value = 1 Then temp = temp Or SWP_NOZORDER
If Check1(1).value = 1 Then temp = temp Or SWP_FRAMECHANGED
If Check1(2).value = 1 Then temp = temp Or SWP_HIDEWINDOW
If Check1(3).value = 1 Then temp = temp Or SWP_DRAWFARME
If Check1(4).value = 1 Then temp = temp Or SWP_NOCOPYBITS
If Check1(5).value = 1 Then temp = temp Or SWP_NOACTIVATE
If Check1(6).value = 1 Then temp = temp Or SWP_SHOWWINDOW
If Check2.value = 0 Then temp = temp Or SWP_NOSIZE
If Check3.value = 0 Then temp = temp Or SWP_NOMOVE
SetWindowPos txthWnd, winst, Val(txtX), Val(txtY), Val(txtCW), Val(txtCH), temp
End Sub

Private Sub isButton2_Click()
    Dim sTitle As String
    ' Ask user for new window title
    sTitle = InputBox("Enter new window title:", "iLDEREMiSpy ")
    ' Set new window title
    SetWindowText txthWnd.Text, sTitle
End Sub

Private Sub isButton3_Click()
    ' Flash window
    FlashWindow txthWnd.Text, 3
End Sub

Private Sub isButton4_Click()
    ' Show window
    ShowWindow txthWnd.Text, SW_SHOW
End Sub

Private Sub isButton5_Click()
    ' Put window on top of all others
    SetWindowPos txthWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub isButton6_Click()
    ' Enable window
    EnableWindow txthWnd.Text, 1
End Sub

Private Sub isButton7_Click()
        Dim lhWnd As Long
        Dim sTitle As String * 255
        Dim sClass As String * 255
        Dim tRC As RECT
        Dim sParentTitle As String * 255
        Dim sParentClass As String * 255
        Dim lhWndParent As Long
        Dim sStyle As String
        Dim lRetVal As Long
        
        ' Get window handle from point
        'lhWnd = WindowFromPoint(tPA.X, tPA.Y)
        ' Get window caption
        lRetVal = GetWindowText(FindWindow(vbNullString, txtWin.Text), sTitle, 255)
        ' Get window class name
        lRetVal = GetClassName(FindWindow(vbNullString, txtWin.Text), sClass, 255)
        ' Get window style
        sStyle = GetWindowStyle(FindWindow(vbNullString, txtWin.Text))
        ' Get window rect
        GetWindowRect lhWnd, tRC
        ' Get window parent
        lhWndParent = GetParent(FindWindow(vbNullString, txtWin.Text))
        ' Get parent window caption
        lRetVal = GetWindowText(lhWndParent, sParentTitle, 255)
        ' Get parent window class name
        lRetVal = GetClassName(lhWndParent, sParentClass, 255)
        
        ' Set values to textboxes
        txthWnd.Text = FindWindow(vbNullString, txtWin.Text)
        txtTitle.Text = sTitle
        txtClass.Text = sClass
        txtStyle.Text = sStyle
        'txtRect.Text = "(" & tRC.Left & ", " & tRC.Top & ") - (" & tRC.Right & ", " & tRC.bottom & ")"
        txtParent.Text = lhWndParent
        txtParentText.Text = sParentTitle
        txtParentClass.Text = sParentClass
        If txthWnd <> "" Then txtTitle = txtWin
End Sub

Private Sub isButton8_Click()
    ' Close window
    SendMessage txthWnd.Text, WM_CLOSE, 0, 0
End Sub

Private Sub isButton9_Click()
    ' Remove window from top
    SetWindowPos txthWnd.Text, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

'////////////////////////////////////////////////////////////////////
'//// CROSSHAIR EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub picCrossHair_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ShowWindow Me.hwnd, 6
ShowWindow frmMain.hwnd, 0
ShowWindow frmScrSvr.hwnd, 0
    ' If user pressed left mouse button and we are not dragging
    If Button = vbLeftButton And Not m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = True
        ' Set mouse pointer
        Me.MouseIcon = imgCursor.MouseIcon
        Me.MousePointer = 99
        ' Erase picture from picCrossHair
        picCrossHair.Picture = Nothing
    End If
End Sub

Private Sub picCrossHair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        Dim tPA As POINTAPI
        Dim lhWnd As Long
        Dim sTitle As String * 255
        Dim sClass As String * 255
        Dim tRC As RECT
        Dim sParentTitle As String * 255
        Dim sParentClass As String * 255
        Dim lhWndParent As Long
        Dim sStyle As String
        Dim lRetVal As Long
        
        ' Get cursor position
        GetCursorPos tPA
        ' Get window handle from point
        lhWnd = WindowFromPoint(tPA.X, tPA.Y)
        ' Get window caption
        lRetVal = GetWindowText(lhWnd, sTitle, 255)
        ' Get window class name
        lRetVal = GetClassName(lhWnd, sClass, 255)
        ' Get window style
        sStyle = GetWindowStyle(lhWnd)
        ' Get window rect
        GetWindowRect lhWnd, tRC
        ' Get window parent
        lhWndParent = GetParent(lhWnd)
        ' Get parent window caption
        lRetVal = GetWindowText(lhWndParent, sParentTitle, 255)
        ' Get parent window class name
        lRetVal = GetClassName(lhWndParent, sParentClass, 255)
        
        ' Set values to textboxes
        txthWnd.Text = lhWnd
        txtTitle.Text = sTitle
        txtClass.Text = sClass
        txtStyle.Text = sStyle
        txtRect.Text = "(" & tRC.Left & ", " & tRC.Top & ") - (" & tRC.Right & ", " & tRC.bottom & ")"
        txtParent.Text = lhWndParent
        txtParentText.Text = sParentTitle
        txtParentClass.Text = sParentClass
    End If
End Sub

Private Sub picCrossHair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = False
        ' Restore mouse pointer to normal (arrow)
        Me.MousePointer = vbNormal
        ' Load picture into picCrossHair
        picCrossHair.Picture = imgCursor.MouseIcon
    End If
ShowWindow frmScrSvr.hwnd, 5
ShowWindow frmMain.hwnd, 5
'ShowWindow Me.hwnd, 5
End Sub

' Get window styles
Private Function GetWindowStyle(ByVal lhWnd As Long) As String
    Dim lStyle As Long
        
    ' Get window styles
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    
    ' Get window styles
    If lStyle And WS_BORDER Then GetWindowStyle = GetWindowStyle & "WS_BORDER "
    If lStyle And WS_CAPTION Then GetWindowStyle = GetWindowStyle & "WS_CAPTION "
    If lStyle And WS_CHILD Then GetWindowStyle = GetWindowStyle & "WS_CHILD "
    If lStyle And WS_CLIPCHILDREN Then GetWindowStyle = GetWindowStyle & "WS_CLIPCHILDREN "
    If lStyle And WS_CLIPSIBLINGS Then GetWindowStyle = GetWindowStyle & "WS_CLIPSIBLINGS "
    If lStyle And WS_DLGFRAME Then GetWindowStyle = GetWindowStyle & "WS_DLGFRAME "
    If lStyle And WS_GROUP Then GetWindowStyle = GetWindowStyle & "WS_GROUP "
    If lStyle And WS_HSCROLL Then GetWindowStyle = GetWindowStyle & "WS_HSCROLL "
    If lStyle And WS_MAXIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MAXIMIZEBOX "
    If lStyle And WS_MINIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MINIMIZEBOX "
    If lStyle And WS_SYSMENU Then GetWindowStyle = GetWindowStyle & "WS_SYSMENU "
    If lStyle And WS_POPUPWINDOW Then GetWindowStyle = GetWindowStyle & "WS_POPUPWINDOW "
    If lStyle And WS_TABSTOP Then GetWindowStyle = GetWindowStyle & "WS_TABSTOP "
    If lStyle And WS_THICKFRAME Then GetWindowStyle = GetWindowStyle & "WS_THICKFRAME "
    If lStyle And WS_VISIBLE Then GetWindowStyle = GetWindowStyle & "WS_VISIBLE "
    If lStyle And WS_VSCROLL Then GetWindowStyle = GetWindowStyle & "WS_VSCROLL "

End Function
' Make textboxes flat
Public Sub MakeFlat(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    RemoveBorder lhWnd
End Sub
Private Sub RemoveBorder(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

' Get current mouse cordinates
Private Sub Timer1_Timer()
    Dim tPA As POINTAPI
    
    GetCursorPos tPA
    lblCordi.Caption = "X: " & tPA.X & "  Y: " & tPA.Y
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub txtTitle_Change()
txtWin.Text = txtTitle.Text
End Sub

Private Function TranslucentForm(TranslucenceLevel As Byte) As Boolean
Dim hwnd
SetWindowLong Val(txthWnd), GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Val(txthWnd), 0, TranslucenceLevel, LWA_ALPHA
TranslucentForm = Err.LastDllError = 0
End Function
