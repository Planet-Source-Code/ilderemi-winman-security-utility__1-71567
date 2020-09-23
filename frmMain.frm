VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14895
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   RightToLeft     =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1560
      Picture         =   "frmMain.frx":00D2
      RightToLeft     =   -1  'True
      ScaleHeight     =   690
      ScaleWidth      =   4230
      TabIndex        =   88
      Top             =   4320
      Width           =   4260
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   480
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   1560
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   49
      Top             =   735
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hide"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   48
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "hh.exe"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   47
      Top             =   1455
      Width           =   1335
   End
   Begin VB.OptionButton Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Explorer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   46
      Top             =   1455
      Width           =   1455
   End
   Begin VB.OptionButton Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Direct"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   45
      Top             =   1455
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   42
      Top             =   255
      Width           =   4575
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0015160C&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   11160
      ScaleHeight     =   103.279
      ScaleMode       =   0  'User
      ScaleWidth      =   100.844
      TabIndex        =   4
      Top             =   8160
      Width           =   3615
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H001E6C00&
         Height          =   945
         Left            =   0
         Top             =   0
         Width           =   3585
      End
   End
   Begin VB.PictureBox picUsage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   10080
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   8160
      Width           =   990
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10920
      Top             =   8040
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7560
      Top             =   7920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4035
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5160
      Width           =   7335
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   480
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3825
         ScaleWidth      =   7065
         TabIndex        =   62
         Top             =   120
         Width           =   7095
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1455
            TabIndex        =   64
            Top             =   1770
            Width           =   3945
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00008000&
            Height          =   280
            Left            =   1380
            Top             =   1760
            Visible         =   0   'False
            Width           =   4095
         End
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H0000FF00&
         Height          =   4040
         Left            =   0
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "KazmeGheyz2"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   2955
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   3720
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   1320
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "BhoNew 1"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   120
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "WINFILE"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   85
      Top             =   2955
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   2520
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "KazmeGheyz1"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   2955
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ghorveh"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   54
      Left            =   13080
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "NTFS Encryption"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   82
      Top             =   5955
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   53
      Left            =   11280
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP IP Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   81
      Top             =   5955
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   52
      Left            =   9480
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hash"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   80
      Top             =   5955
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   51
      Left            =   7680
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MDB2000"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   79
      Top             =   5955
      Width           =   1695
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   2520
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IR-NewFolder[Ali Sadeghi-Vir]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   78
      Top             =   2475
      Width           =   2295
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   4920
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldost"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Important"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   4920
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IR-Thumbs"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   3720
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Output is Disable"
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   74
      Top             =   7395
      Width           =   3495
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000FF&
      Height          =   375
      Index           =   50
      Left            =   9480
      Top             =   7320
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Output is Disable"
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   73
      Top             =   7395
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GeoClock"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   72
      Top             =   5475
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   49
      Left            =   7680
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Remove Messenger"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   71
      Top             =   5475
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   48
      Left            =   9480
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "NoLogOff Update"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   70
      Top             =   5475
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   47
      Left            =   11280
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TurnOff Monitor"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   69
      Top             =   5475
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   37
      Left            =   13080
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   120
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   1320
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   2520
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   6360
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label isButton26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "NTDETECT"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label isButton30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "SVCHOST"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label isButton28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MSFUN80"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label isButton35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Anti Hider"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Print URL"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   63
      Top             =   3075
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   61
      Top             =   7395
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00008000&
      Height          =   375
      Index           =   46
      Left            =   7680
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   45
      Left            =   13080
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   44
      Left            =   7680
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   43
      Left            =   9480
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP Get/Post"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   60
      Top             =   4995
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   42
      Left            =   11280
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Device Manager"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   59
      Top             =   4995
      Width           =   1695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Shell"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   58
      Top             =   4995
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Floppy"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   57
      Top             =   4995
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IP Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   56
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   41
      Left            =   7680
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   40
      Left            =   9480
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   39
      Left            =   11280
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   38
      Left            =   13080
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Saver"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   55
      Top             =   3075
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Attribute"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   54
      Top             =   3075
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   13080
      TabIndex        =   53
      Top             =   7395
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   675
      Left            =   1545
      Top             =   720
      Width           =   4605
   End
   Begin VB.Label isButton3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   51
      Top             =   780
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   6240
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Run:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   855
      Width           =   1095
   End
   Begin VB.Label isButton20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   44
      Top             =   280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Del File:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   255
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   1545
      Top             =   240
      Width           =   4605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Left            =   6240
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label isButton9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Map Network Drive"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   41
      Top             =   1635
      Width           =   1695
   End
   Begin VB.Label isButton25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tasks Optimizer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   40
      Top             =   1635
      Width           =   1695
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Text"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   39
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MMC"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   38
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Label isButton38 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Keylogger"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   13080
      TabIndex        =   37
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Label isButton36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Compressed Folder"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   36
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Label isButton240 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Pass"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   35
      Top             =   2115
      Width           =   1695
   End
   Begin VB.Label isButton21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Pass"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   34
      Top             =   2115
      Width           =   1695
   End
   Begin VB.Label isButton24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Favorites Org"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   2115
      Width           =   1695
   End
   Begin VB.Label isButton16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "About Windows"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   32
      Top             =   2115
      Width           =   1695
   End
   Begin VB.Label isButton10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Shares"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   31
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label isButton15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "FDD Format"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   30
      Top             =   1635
      Width           =   1695
   End
   Begin VB.Label isButton18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Test Page"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   29
      Top             =   1635
      Width           =   1695
   End
   Begin VB.Label isButton19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Restart"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   28
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label isButton13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CPL Files"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   27
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label isButton14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Connect to Printer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label isButton7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Windows"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   25
      Top             =   675
      Width           =   1695
   End
   Begin VB.Label isButton8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Share"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   24
      Top             =   675
      Width           =   1695
   End
   Begin VB.Label isButton12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Open as..."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   23
      Top             =   675
      Width           =   1695
   End
   Begin VB.Label isButton5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Printer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   22
      Top             =   675
      Width           =   1695
   End
   Begin VB.Label isButton11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Standby"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   21
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label isButton6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Add Net Place Wiz"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   20
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label isButton17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Eject Hardware"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   19
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label isButton4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TCP/IP printer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   195
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   23
      Left            =   13080
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   22
      Left            =   11280
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   21
      Left            =   9480
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   20
      Left            =   7680
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   19
      Left            =   13080
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   18
      Left            =   11280
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   17
      Left            =   9480
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   16
      Left            =   7680
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   15
      Left            =   13080
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   14
      Left            =   11280
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   13
      Left            =   9480
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   12
      Left            =   7680
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   11
      Left            =   13080
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   10
      Left            =   11280
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   9
      Left            =   9480
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   7680
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   7
      Left            =   13080
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   6
      Left            =   11280
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   9480
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   7680
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   13080
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   11280
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   9480
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   7680
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label isButton37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Make CAB"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   17
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   35
      Left            =   11280
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label isButton23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Directory"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   16
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   34
      Left            =   9480
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   15
      Top             =   3075
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   33
      Left            =   7680
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Group Policy"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   14
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   32
      Left            =   13080
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label isButton27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "System Execution"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   31
      Left            =   11280
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Windows"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   12
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   30
      Left            =   9480
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label isButton31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Port Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   29
      Left            =   7680
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label isButton34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Outlook E-Mail"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   10
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   28
      Left            =   13080
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label isButton22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Encoder"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13080
      TabIndex        =   9
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   27
      Left            =   13080
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label isButton29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Key ASCII/Code"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   8
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   26
      Left            =   11280
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label isButton33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Process"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   7
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   25
      Left            =   9480
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label isButton32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CMD"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   24
      Left            =   7680
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label lblDate2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0000FF00&
      Height          =   9330
      Left            =   0
      Top             =   0
      Width           =   14895
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00008000&
      Height          =   375
      Index           =   36
      Left            =   13080
      Top             =   7320
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
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
Dim TC As Byte
Dim g As Double
Dim cx, cy, sv, sa As Integer
Dim mystring(2) As String
Private QueryObject As Object

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Sub Form_Load()
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA
mystring(0) = "Created by Masoud iLDEREMi"
On Error Resume Next
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

'====Farsi====
MyLang = Combo1.text
'ChLang
'=============

  Set myShadow = New clsShadow
  With myShadow
    If .Shadow(Me) Then
      .Depth = 10
      .Transparency = 128
    Else
      Set myShadow = Nothing
    End If
  End With
    
   
    '===================================================================
    MyLang = "English"
    'ChLang
    '===================================================================
   
    '===================================================================
    'set the Priority of this process to 'High'
    'this makes sure our program gets updated, even when
    'another process is consuming lots of CPU cycles
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize
    'start the timer
    tmrRefresh.Enabled = True
    'don't wait for the first interval to elapse
    tmrRefresh_Timer
    '===================================================================
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ReleaseCapture
    'SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'stop the timer
    tmrRefresh.Enabled = False
    'clean up
    QueryObject.Terminate
    Set QueryObject = Nothing
    End
End Sub

Private Sub isButton1_Click()
frmShortcut.Show 1
End Sub

Private Sub isButton10_Click()
If Priv Then Text3 = "RUNDLL32 ntlanui.dll,ShareManage - Shares"
Shell "RUNDLL32 ntlanui.dll,ShareManage - Shares"
End Sub

Private Sub isButton11_Click()
If Priv Then Text3 = "RUNDLL32 PowrProf.dll, SetSuspendState"
Shell "RUNDLL32 PowrProf.dll, SetSuspendState"
End Sub

Private Sub isButton12_Click()
Text3 = "RUNDLL32 SHELL32.DLL,OpenAs_RunDLL filename"
End Sub

Private Sub isButton13_Click()
frmCPL.Show 1
End Sub

Private Sub isButton14_Click()
If Priv Then Text3 = "RUNDLL32 WINSPOOL.DRV,ConnectToPrinterDlg"
Shell "RUNDLL32 WINSPOOL.DRV,ConnectToPrinterDlg"
End Sub

Private Sub isButton15_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,SHFormatDrive"
Shell "RUNDLL32 SHELL32.DLL,SHFormatDrive"
End Sub

Private Sub isButton16_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,ShellAboutW"
Shell "RUNDLL32 SHELL32.DLL,ShellAboutW"
End Sub

Private Sub isButton17_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,Control_RunDLL hotplug.dll"
Shell "RUNDLL32 SHELL32.DLL,Control_RunDLL hotplug.dll"
End Sub

Private Sub isButton18_Click()
If Priv Then Text3 = "RUNDLL32.EXE SHELL32.DLL,SHHelpShortcuts_RunDLL PrintTestPage"
Shell "RUNDLL32.EXE SHELL32.DLL,SHHelpShortcuts_RunDLL PrintTestPage"
End Sub

Private Sub isButton19_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,RestartDialog"
Shell "RUNDLL32 SHELL32.DLL,RestartDialog"
End Sub

Private Sub isButton2_Click()
On Error GoTo 10
'frmWindow.Show 1
'Exit Sub
10: frmWindow.Show , Me
End Sub

Private Sub isButton20_Click()
On Error Resume Next
Kill Text4
End Sub

Private Sub isButton21_Click()
    Dim strfile As String, strPath As String, strEXE As String
    strfile = SPACE(255)
    GetSystemDirectory strfile, 255&
    
    strfile = Replace(Trim(strfile), Chr$(0), "")
    strPath = Trim(strfile) & "\logonui.exe"
    FileCopy strPath, "c:/logonui.exe"
    Close
    Open strPath For Binary As #1
    strEXE = SPACE(LOF(1))
    Get #1, , strEXE
    Close
    
    strEXE = Replace(strEXE, "edit [id=atom(password)]", "edit [id=atom(keyboard)]")

    Kill strPath
    Open strPath For Binary As #2
    Put #2, , strEXE
    If InStr(strEXE, "edit [id=atom(keyboard)]") > 0 Then MsgBox "Changed Successfully!!!" & vbLf & "Restart your computer to see the change." Else MsgBox "An error occured!!"
    Close
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
End Sub

Private Sub isButton22_Click()
frmHTMLEn.Show , Me
End Sub

Private Sub isButton23_Click()
frmDir.Show , Me
End Sub

Private Sub isButton24_Click()
If Priv Then Text3 = "RUNDLL32.EXE shdocvw.dll,DoOrganizeFavDlg"
Shell "RUNDLL32.EXE shdocvw.dll,DoOrganizeFavDlg"
End Sub

Private Sub isButton240_Click()
    Dim strfile As String, strPath As String, strEXE As String
    strfile = SPACE(255)
    GetSystemDirectory strfile, 255&
    
    strfile = Replace(Trim(strfile), Chr$(0), "")
    strPath = Trim(strfile) & "\logonui.exe"
    FileCopy strPath, "c:/logonui.exe"
    Close
    Open strPath For Binary As #1
    strEXE = SPACE(LOF(1))
    Get #1, , strEXE
    Close

    strEXE = Replace(strEXE, "edit [id=atom(keyboard)]", "edit [id=atom(password)]")

    Kill strPath
    Open strPath For Binary As #2
    Put #2, , strEXE
    If InStr(strEXE, "edit [id=atom(password)]") > 0 Then MsgBox "Changed Successfully!!!" & vbLf & "Restart your computer to see the change." Else MsgBox "An error occured!!"
    Close
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
End Sub

Private Sub isButton25_Click()
If Priv Then Text3 = "rundll32.exe advapi32.dll,ProcessIdleTasks"
Shell "rundll32.exe advapi32.dll,ProcessIdleTasks"
End Sub

Private Sub isButton26_Click()
Shell "taskkill /f /im """"UpDateWinc.exe""", vbHide
Shell "taskkill /f /im """"NTDETECT.exe""", vbHide
Shell "taskkill /f /im """"a1.exe""", vbHide
Shell "taskkill /f /im """"tem.exe""", vbHide
Shell "attrib -h -s -a -r %Systemroot%\system32\UpDateWinc.exe", vbHide
Shell "attrib -h -s -a -r %Systemroot%\system32\UpDateWind.exe", vbHide
Shell "attrib -h -s -a -r %Systemroot%\LogBoy.log", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\a1.exe", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\pass1.txt", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\tem.exe", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\temp1.bat", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\NTDETECT.exe", vbHide
Shell "attrib -h -s -a -r %SystemDrive%\autorun.inf", vbHide
sv = 0
Timer3.Enabled = True
End Sub

Private Sub isButton27_Click()
frmSysEXE.Show , Me
End Sub

Private Sub isButton28_Click()
Shell "taskkill -f -im " & Chr(34) & "fun.xls.exe" & Chr(34), vbHide
Shell "taskkill -f -im " & Chr(34) & "Autorun.exe" & Chr(34), vbHide
sv = 2
Timer3.Enabled = True
End Sub

Private Sub isButton29_Click()
frmCodeAscii.Show , Me
End Sub

Private Sub isButton3_Click()
On Error Resume Next
If Check2(1).value Then
    If Check1.value Then
    Shell "explorer " & Text3, vbHide
    Else
    Shell "explorer " & Text3
    End If
    Exit Sub
End If
If Check2(0).value Then
    If Check1.value Then
    Shell "hh " & Text3, vbHide
    Else
    Shell "hh " & Text3
    End If
    Exit Sub
End If
If Check2(2).value Then
    If Check1.value Then
    Shell Text3, vbHide
    Else
    Shell Text3
    End If
    Exit Sub
End If
End Sub

Private Sub isButton30_Click()
frmAntiSVCHOST.Show , Me
End Sub

Private Sub isButton31_Click()
frmPortScanner.Show , Me
End Sub

Private Sub isButton32_Click()
frmCMD.Show 1
End Sub

Private Sub isButton33_Click()
frmProcess.Show 1
End Sub

Private Sub isButton34_Click()
frmOutMail.Show 1
End Sub

Private Sub isButton35_Click()
frmUnHider.Show , Me
End Sub

Private Sub isButton36_Click()
Text3 = "RUNDLL32 zipfldr.dll,RouteTheCall [file_address]"
End Sub

Private Sub isButton37_Click()
frmMkCab.Show 1
End Sub

Private Sub isButton38_Click()
Dim KLPath As String
CommonDialog1.ShowSave
KLPath = CommonDialog1.Filename
If KLPath <> "" Then
    LoadEXE (102)
    Open KLPath For Binary As #5
        Put #5, 1, EXEFile
    Close
    Shell "cmd /c " & Chr(34) & CommonDialog1.Filename & Chr(34)
End If
End Sub

Private Sub isButton4_Click()
If Priv Then Text3 = "RUNDLL32 tcpmonui.dll,LocalAddPortUI"
Shell "RUNDLL32 tcpmonui.dll,LocalAddPortUI"
End Sub

Private Sub isButton5_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,SHHelpShortcuts_RunDLL AddPrinter"
Shell "RUNDLL32 SHELL32.DLL,SHHelpShortcuts_RunDLL AddPrinter"
End Sub

Private Sub isButton6_Click()
If Priv Then Text3 = "RUNDLL32 netplwiz.dll,AddNetPlaceRunDll"
Shell "RUNDLL32 netplwiz.dll,AddNetPlaceRunDll"
End Sub

Private Sub isButton7_Click()
If Priv Then Text3 = "RUNDLL32 USER32.DLL,LockWorkStation"
Shell "RUNDLL32 USER32.DLL,LockWorkStation"
End Sub

Private Sub isButton8_Click()
If Priv Then Text3 = "RUNDLL32 ntlanui.dll,ShareCreate"
Shell "RUNDLL32 ntlanui.dll,ShareCreate"
End Sub

Private Sub isButton9_Click()
If Priv Then Text3 = "RUNDLL32 SHELL32.DLL,SHHelpShortcuts_RunDLL Connect"
Shell "RUNDLL32 SHELL32.DLL,SHHelpShortcuts_RunDLL Connect"
End Sub

Private Sub Label11_Click()
Close
On Error Resume Next
Unload Me
End
End Sub

Private Sub Label12_Click()
    '===================================================================
    H_About.Show , Me
    Form1.Left = H_About.Left
    Form1.Top = H_About.Top - Form1.Height
    Form1.Show , Me
    '===================================================================
End Sub

Private Sub Label13_Click()
frmPrintURL.Show , Me
End Sub

Private Sub Label14_Click()
frmAttrib.Show , Me
End Sub

Private Sub Label15_Click()
SendMessage Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
End Sub

Private Sub Label16_Click()
frmIPScan.Show , Me
End Sub

Private Sub Label17_Click()
If Priv Then Text3 = "RUNDLL32 DISKCOPY.DLL,DiskCopyRunDll"
Shell "RUNDLL32 DISKCOPY.DLL,DiskCopyRunDll"
End Sub

Private Sub Label18_Click()
frmShell.Show , Me
End Sub

Private Sub Label19_Click()
If Priv Then Text3 = "RUNDLL32 devmgr.dll DeviceManager_Execute"
Shell "RUNDLL32 devmgr.dll DeviceManager_Execute"
End Sub

Private Sub Label20_Click()
frmGetPost.Show , Me
End Sub

Private Sub Label22_Click()
Shell "cmd /c taskkill /f /im """"Thumbs.exe"""
Shell "attrib -s -h -a -r %SystemDrive%\Thumbs.exe", vbHide
Shell "attrib -s -h -a -r %SystemDrive%\Autorun.inf", vbHide
sv = 1
Timer3.Enabled = True
End Sub

Private Sub Label24_Click()
Shell "mmc"
End Sub

Private Sub Label25_Click()
frmPrintText.Show , Me
End Sub

Private Sub Label26_Click()
Shell "cmd /c taskkill /f /im """"autoply.exe"""
Shell "cmd /c taskkill /f /im """"SoundMax.exe"""
Shell "cmd /c taskkill /f /im """"OfficeUpdate.exe"""
Shell "cmd /c taskkill /f /im """"MSshare.exe"""
Shell "cmd /c taskkill /f /im """"Sex_Game.exe"""
Shell "cmd /c taskkill /f /im """"Sex_ScreenSaver.scr"""

Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Common Files\MicrosoftShared\MSshare.exe""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Documents and Settings\AllUsers\StartMenu\Programs\Startup\OfficeUpdate\Important.htm""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\eMule\Incoming\Sex_ScreenSaver.scr""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\eMule\Incoming\Sex_Game.exe""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Kazaa\MySearchAgents\Sex_ScreenSaver.scr""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Kazaa\MySharedFolderS\Sex_ScreenSaver.scr""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Kazaa\MySharedFolderS\Sex_Game.exe""", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Kazaa\MySearchAgents\Sex_Game.exe""", vbHide

Shell "attrib -s -h -a -r %SystemDrive%\autoply.exe", vbHide
Shell "attrib -s -h -a -r """"%SystemDrive%\Program Files\Sound Utility\SoundMax.exe""", vbHide
Shell "attrib -s -h -a -r %SystemDrive%\Autorun.inf", vbHide
sv = 4
Timer3.Enabled = True
End Sub

Private Sub Label27_Click()
Shell "cmd /c taskkill /f /im """"Windows Explorer.exe"""
Shell "attrib -s -h -a -r """"%windir%\Windows Explorer.exe""", vbHide
sv = 3
Timer3.Enabled = True
End Sub

Private Sub Label28_Click()
frmMDB2000.Show , Me
End Sub

Private Sub Label29_Click()
frmHash.Show , Me
End Sub

Private Sub Label30_Click()
frmHTTPScan.Show , Me
End Sub

Private Sub Label31_Click()
frmNTFSenc.Show , Me
End Sub

Private Sub Label4_Click()
If Priv Then Text3 = "RUNDLL32 PowrProf.dll, SetSuspendState"
Shell "RUNDLL32 PowrProf.dll, SetSuspendState"
End Sub

Private Sub Label5_Click()
If Priv Then Text3 = "RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True"
Shell "RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True"
End Sub

Private Sub Label6_Click()
If Priv Then Text3 = "RUNDLL32 advpack.dll,LaunchINFSection %windir%\INF\msmsgs.inf,BLC.Remove"
Shell "RUNDLL32 advpack.dll,LaunchINFSection %windir%\INF\msmsgs.inf,BLC.Remove"
End Sub

Private Sub Label7_Click()
frmGeoCalc.Show , Me
End Sub

Private Sub Picture1_Paint()
'Picture1.Cls
End Sub

Private Sub Timer1_Timer()
MTX (True)
MTX (False)
End Sub

Private Sub Timer2_Timer()
Label9.Caption = Time$
'Call rfd
lblDate2 = WeekdayName(Weekday(Date), False) & ", " & Day(Date) & ", " & MonthName(Month(Date), False) & ", " & Year(Date)
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
sa = sa + 3
If sv = 0 Then
    If sa >= 3 Then
        Shell "cmd /c del /f %Systemroot%\system32\UpDateWinc.exe", vbHide
        Shell "cmd /c del /f %Systemroot%\system32\UpDateWind.exe", vbHide
        Shell "cmd /c del /f %Systemroot%\LogBoy.log", vbHide
        Shell "cmd /c del /f %SystemDrive%\a1.exe", vbHide
        Shell "cmd /c del /f %SystemDrive%\pass1.txt", vbHide
        Shell "cmd /c del /f %SystemDrive%\tem.exe", vbHide
        Shell "cmd /c del /f %SystemDrive%\temp1.bat", vbHide
        Shell "cmd /c del /f %SystemDrive%\NTDETECT.exe", vbHide
        Shell "cmd /c del /f %SystemDrive%\autorun.inf", vbHide
        Shell "cmd /c reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\run /v " & Chr(34) & "UpDateWinc" & Chr(34), vbHide
        Shell "cmd /c reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\run /v " & Chr(34) & "UpDateWinc.exe" & Chr(34), vbHide
        Shell "cmd /c reg delete HKLM\Software\Microsoft\Windows\CurrentVersion\run /v " & Chr(34) & "UpDateWinc" & Chr(34), vbHide
        Shell "cmd /c reg delete HKLM\Software\Microsoft\Windows\CurrentVersion\run /v " & Chr(34) & "UpDateWinc.exe" & Chr(34), vbHide
        For i = 0 To Drive1.ListCount - 1
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\NTDETECT.exe", vbHide
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\Autorun.inf", vbHide
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\NTDETECT.exe", vbHide
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\Autorun.inf", vbHide
        Next
        Timer3.Enabled = False
    End If
ElseIf sv = 1 Then
    If sa >= 3 Then
        Shell "cmd /c del /f %SystemDrive%\Thumbs.exe", vbHide
        Shell "cmd /c del /f %SystemDrive%\Autorun.inf", vbHide
        For i = 0 To Drive1.ListCount - 1
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\Thumbs.exe", vbHide
            Shell "attrib -s -h -a -r " & Drive1.List(i) & "\Autorun.inf", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Thumbs.exe", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.inf", vbHide
        Next
        Timer3.Enabled = False
    End If
ElseIf sv = 2 Then
    If sa >= 3 Then
        For i = 0 To Drive1.ListCount - 1
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.~ex", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.bat", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.bin", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.exe", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.ico", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\AUTORUN.INF", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.ini", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.reg", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.srm", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.txt", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.vbs", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\autorun.wsh", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\fun.xls.exe", vbHide
            Shell "cmd /c reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced /v ShowSuperHidden /t REG_DWORD /d 0x00000001 /f", vbHide
            Shell "cmd /c reg add " & Chr(34) & "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" & Chr(34) & " /v Userinit /t REG_SZ /d ^userinit.exe^ /f"
        Next
        Timer3.Enabled = False
    End If
ElseIf sv = 3 Then
    If sa >= 3 Then
        Shell "cmd /c del /f """"%windir%\Windows Explorer.exe""", vbHide
        Shell "cmd /c reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\run /v " & Chr(34) & "Windows Explorer" & Chr(34), vbHide
        Timer3.Enabled = False
    End If
ElseIf sv = 4 Then
    If sa >= 3 Then
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Sound Utility\SoundMax.exe""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Common Files\MicrosoftShared\MSshare.exe""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Documents and Settings\AllUsers\StartMenu\Programs\Startup\OfficeUpdate\Important.htm""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\eMule\Incoming\Sex_ScreenSaver.scr""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\eMule\Incoming\Sex_Game.exe""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Kazaa\MySearchAgents\Sex_ScreenSaver.scr""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Kazaa\MySharedFolderS\Sex_ScreenSaver.scr""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Kazaa\MySharedFolderS\Sex_Game.exe""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\Program Files\Kazaa\MySearchAgents\Sex_Game.exe""", vbHide
    Shell "cmd /c del /f """"%SystemDrive%\""", vbHide
    
        For i = 0 To Drive1.ListCount - 1
            Shell "cmd /c del /f " & Drive1.List(i) & "\autoply.exe", vbHide
            Shell "cmd /c del /f " & Drive1.List(i) & "\Autorun.inf", vbHide
            Shell "cmd /c reg delete HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v """"SoundMax""", vbHide
            Shell "cmd /c reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v """"SoundMax""", vbHide
            'Shell "cmd /c reg add " & Chr(34) & "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" & Chr(34) & " /v Userinit /t REG_SZ /d ^userinit.exe^ /f"
        Next
        Timer3.Enabled = False
    End If
End If
End Sub

Private Sub Timer4_Timer()
'Shape13.Refresh
If g <= Len(mystring(0)) + 60 Then
    g = g + 0.2
    Label3.Caption = Mid(mystring(0), 1, Int(g))
ElseIf g <= Len(mystring(0)) + 60 + Len(mystring(1)) + 60 Then
    g = 0
End If
Label3.Refresh
MTX (True)
MTX (False)
End Sub

Private Sub Timer5_Timer()
TC = TC + 5
If TC < 255 Then
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, TC, LWA_ALPHA
Else
    Me.Hide
    Me.Show
    Timer5.Enabled = False
End If
End Sub

Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    'query the CPU usage
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) + "%"
    End If
    
End Sub

Sub MTX(b As Boolean)
On Error Resume Next
Dim a As Integer
Dim C As String
a = Int(Rnd * 1.9)
Picture1.CurrentX = Int(Rnd * Picture1.Width / 200) * 200
cx = Picture1.CurrentX
Picture1.CurrentY = Int(Rnd * Picture1.Height / 200) * 200
cy = Picture1.CurrentY

'Picture1.Line (CX, CY)-Step(200, 200), , BF

If b = True Then
    Picture1.ForeColor = 0
    Picture1.Line (cx, cy)-Step(200, 200), , BF
    
    Picture1.CurrentX = cx
    Picture1.CurrentY = cy
    C = "&h" & Int(Rnd * &HFF00 + &H80) & "00"
    Picture1.ForeColor = C
    Picture1.Print a
End If
End Sub
Private Function Priv() As Boolean
If Label8.Caption = "Output is Disable" Then
    Priv = False
Else
    Priv = True
End If
End Function
