VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmCMDLine 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCMDLine.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0015160C&
      ForeColor       =   &H0000FF00&
      Height          =   5535
      Left            =   10
      Picture         =   "frmCMDLine.frx":A2E52
      ScaleHeight     =   5505
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   10
      Width           =   9015
      Begin VB.Frame fmMainPass 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   855
         Left            =   1080
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   6855
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   830
            Left            =   10
            ScaleHeight     =   795
            ScaleWidth      =   6795
            TabIndex        =   20
            Top             =   10
            Width           =   6830
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               IMEMode         =   3  'DISABLE
               Left            =   6600
               PasswordChar    =   "l"
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   -480
               Width           =   270
            End
            Begin VB.Timer Timer4 
               Enabled         =   0   'False
               Interval        =   30
               Left            =   6360
               Top             =   0
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00000080&
               Height          =   800
               Left            =   0
               Top             =   0
               Width           =   6800
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   6615
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   6495
            End
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6720
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   470
            Left            =   10
            ScaleHeight     =   435
            ScaleWidth      =   1875
            TabIndex        =   17
            Top             =   10
            Width           =   1910
            Begin VB.Shape Shape2 
               BorderColor     =   &H0000FF00&
               Height          =   435
               Left            =   0
               Top             =   0
               Width           =   1870
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "No Track"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   18
               Top             =   95
               Width           =   1935
            End
         End
      End
      Begin VB.Timer Timer3 
         Interval        =   2000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame fmDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   6720
         TabIndex        =   9
         Top             =   60
         Width           =   1935
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   10
            Picture         =   "frmCMDLine.frx":145CA4
            ScaleHeight     =   1635
            ScaleWidth      =   1875
            TabIndex        =   10
            Top             =   10
            Width           =   1910
            Begin MSWinsockLib.Winsock Winsock2 
               Left            =   0
               Top             =   480
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.Timer Timer2 
               Interval        =   1000
               Left            =   0
               Top             =   0
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   15
               Top             =   1080
               Width           =   1875
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   14
               Top             =   120
               Width           =   1875
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   13
               Top             =   360
               Width           =   1875
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   12
               Top             =   600
               Width           =   1875
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   11
               Top             =   840
               Width           =   1875
            End
         End
      End
      Begin VB.Frame fmPortScan 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1530
         Left            =   6720
         TabIndex        =   7
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
         Begin MSWinsockLib.Winsock Winsock1 
            Left            =   240
            Top             =   960
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   240
            Top             =   480
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1500
            Left            =   10
            TabIndex        =   8
            Top             =   10
            Width           =   1905
         End
      End
      Begin VB.Frame fmDebug 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   6720
         TabIndex        =   5
         Top             =   2340
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1545
            Left            =   15
            TabIndex        =   6
            Top             =   15
            Width           =   1905
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   8820
         TabIndex        =   4
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "^"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   8805
         TabIndex        =   3
         Top             =   60
         Width           =   255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0015160C&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   5490
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8685
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00443002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         Height          =   5535
         Left            =   8715
         Top             =   -15
         Width           =   330
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   5565
      Left            =   0
      Top             =   0
      Width           =   9045
   End
End
Attribute VB_Name = "frmCMDLine"
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

Dim bp, SCL, strPort, sPass As Integer
Dim MyPData As String


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
Dim ar As Variant
Dim StdStr, strTemp As String
Dim StdIndex As Integer
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long





Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'-------------------------------------------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
  (ByVal lpBuffer As String, nSize As Long) As Long
'-------------------------------------------------------------------------------------
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
  (ByVal lpBuffer As String, nSize As Long) As Long
'-------------------------------------------------------------------------------------
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Windows version constants
Private Const WIN_VER_MAJ_9XNT4 = 4 'Windows 95/98/ME/NT4
Private Const WIN_VER_MAJ_NT3 = 3 'Windows NT3
Private Const WIN_VER_MAJ_2KXP = 5 'Windows NT5

Private Const WIN_VER_MIN_95 = 0    'Win95 minor
Private Const WIN_VER_MIN_98 = 10   'Win98 minor
Private Const WIN_VER_MIN_ME = 90   'WinME minor
Private Const WIN_VER_MIN_NT3 = 51  'WinNT3.51 minor
Private Const WIN_VER_MIN_NT4 = 0   'WinNT4 minor
Private Const WIN_VER_MIN_2K = 0    'Win2k minor
Private Const WIN_VER_MIN_XP = 1    'WinXP(Whistler) minor

'Platform ID
Private Const VER_PLATFORM_WIN32s = 0 'Win32s
Private Const VER_PLATFORM_WIN32_WINDOWS = 1 'Windows 9x
Private Const VER_PLATFORM_WIN32_NT = 2 'Windows NT

Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias _
  "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
  (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
  lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Enum enuDriveType
    DRIVE_UNKNOWN = 0
    DRIVE_NO_ROOT_DIR = 1
    DRIVE_REMOVABLE = 2     'floppy
    DRIVE_FIXED = 3         ' hard disk
    DRIVE_REMOTE = 4        ' network drive
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum enuStartup
    Normal = 0
    Safe = 1
    SafeWithNetwork = 2
End Enum

Private Const SM_CXSCREEN = 0        ' Width of screen
Private Const SM_CYSCREEN = 1        ' Height of screen
Private Const SM_CXFULLSCREEN = 16   ' Width of window client area
Private Const SM_CYFULLSCREEN = 17   ' Height of window client area
Private Const SM_CYMENU = 15         ' Height of menu
Private Const SM_CYCAPTION = 4       ' Height of caption or title
Private Const SM_CXFRAME = 32        ' Width of window frame
Private Const SM_CYFRAME = 33        ' Height of window frame
Private Const SM_CXHSCROLL = 21      ' Width of arrow bitmap on horizontal scroll bar
Private Const SM_CYHSCROLL = 3       ' Height of arrow bitmap on horizontal scroll bar
Private Const SM_CXVSCROLL = 2       ' Width of arrow bitmap on vertical scroll bar
Private Const SM_CYVSCROLL = 20      ' Height of arrow bitmap on vertical scroll bar
Private Const SM_CXSIZE = 30         ' Width of bitmaps in title bar
Private Const SM_CYSIZE = 31         ' Height of bitmaps in title bar
Private Const SM_CXCURSOR = 13       ' Width of cursor
Private Const SM_CYCURSOR = 14       ' Height of cursor
Private Const SM_CXBORDER = 5        ' Width of window frame that cannot be sized
Private Const SM_CYBORDER = 6        ' Height of window frame that cannot be sized
Private Const SM_CXDOUBLECLICK = 36  ' Width of rectangle around the location of the first click. The
                                     '  second click must occur in the same rectangular location.
Private Const SM_CYDOUBLECLICK = 37  ' Height of rectangle around the location of the first click. The
                                     '  second click must occur in the same rectangular location.
Private Const SM_CXDLGFRAME = 7      ' Width of dialog frame window
Private Const SM_CYDLGFRAME = 8      ' Height of dialog frame window
Private Const SM_CXICON = 11         ' Width of icon
Private Const SM_CYICON = 12         ' Height of icon
Private Const SM_CXICONSPACING = 38  ' Width of rectangles the system uses to position tiled icons
Private Const SM_CYICONSPACING = 39  ' Height of rectangles the system uses to position tiled icons
Private Const SM_CXMIN = 28          ' Minimum width of window
Private Const SM_CYMIN = 29          ' Minimum height of window
Private Const SM_CXMINTRACK = 34     ' Minimum tracking width of window
Private Const SM_CYMINTRACK = 35     ' Minimum tracking height of window
Private Const SM_CXHTHUMB = 10       ' Width of scroll box (thumb) on horizontal scroll bar
Private Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on vertical scroll bar
Private Const SM_DBCSENABLED = 42    ' Returns a non-zero if the current Windows version uses double-byte
                                     '  characters, otherwise returns zero
Private Const SM_DEBUG = 22          ' Returns non-zero if the Windows version is a debugging version
Private Const SM_MENUDROPALIGNMENT = 40
                                     ' Alignment of pop-up menus. If zero, left side is aligned with
                                     '  corresponding left side of menu-bar item. If non-zero, left side
                                     '  is aligned with right side of corresponding menu bar item
Private Const SM_MOUSEPRESENT = 19   ' Non-zero if mouse hardware is installed
Private Const SM_PENWINDOWS = 41     ' Handle of Pen Windows dynamic link library if Pen Windows is
                                     '  installed
Private Const SM_SWAPBUTTON = 23     ' Non-zero if the left and right mouse buttons are swapped
Private Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons
Private Const SM_CLEANBOOT = 67     'How did machine boot
Private Const SM_MOUSEWHEELPRESENT = 75
                                    'Is there a mousewheel?
Private Const SM_SHOWSOUNDS = 70    'Show visual feedback for sounds?
Private Const SM_NETWORK = 63       'Network present if LSB <> 0

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'-------------------------------------------------------------------------------------

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'-------------------------------------------------------------------------------------

Private Function GetNetworked() As String

    Dim lngL As Long

    lngL = GetSystemMetrics(SM_NETWORK)

    If lngL <> 0 Then
        GetNetworked = "Network present = yes" + vbCrLf
    Else
        GetNetworked = "Network present = nos" + vbCrLf
    End If

End Function

Private Function GetLastBootState() As String

    Dim lngL As Long

    lngL = GetSystemMetrics(SM_CLEANBOOT)

    Select Case lngL
        Case Normal
            GetLastBootState = "Started in normal mode" + vbCrLf
        Case Safe
            GetLastBootState = "Started in safe mode" + vbCrLf
        Case SafeWithNetwork
            GetLastBootState = "Started in safe mode with network" + vbCrLf
        Case Else
            GetLastBootState = "Started in unknown operating mode" + vbCrLf
    End Select

End Function

Private Function GetWinVer() As String

    Dim strTemp As String
    Dim OSInfo As OSVERSIONINFO
    Dim lngL As Long
    
    'Preset the size of the structure
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    lngL = GetVersionEx(OSInfo)
    
    If lngL <> 0 Then
        Select Case OSInfo.dwMajorVersion
            Case WIN_VER_MAJ_9XNT4
                If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
                    'Windows 9x Kernel, figure out which edition
                    If OSInfo.dwMinorVersion = WIN_VER_MIN_95 Then
                        strTemp = "Windows95 "
                    ElseIf OSInfo.dwMinorVersion = WIN_VER_MIN_98 Then
                        strTemp = "Windows98 "
                    ElseIf OSInfo.dwMinorVersion = WIN_VER_MIN_ME Then
                        strTemp = "Windows ME "
                    Else
                        strTemp = "Unknown Windows 9x system "
                    End If
                Else
                    'NT4 kernel
                    If OSInfo.dwMinorVersion = WIN_VER_MIN_NT4 Then
                        strTemp = "Windows NT 4 "
                    Else
                        strTemp = "Unknown NT 4-based version "
                    End If
               End If
               
            Case WIN_VER_MAJ_NT3
                    strTemp = "Windows NT 3." & OSInfo.dwMinorVersion & " "
                
            Case WIN_VER_MAJ_2KXP
                If OSInfo.dwMinorVersion = WIN_VER_MIN_2K Then
                    strTemp = "Windows 2000 "
                ElseIf OSInfo.dwMinorVersion = WIN_VER_MIN_XP Then
                    strTemp = "Windows XP (Whistler) "
                Else
                    strTemp = "Unknown Windows NT 5 system "
                End If
            
            Case Else
                strTemp = "Unknown Windows system"
        End Select
            
        'Get service pack level information
        strTemp = strTemp + StripNullTerminator(OSInfo.szCSDVersion) + vbCrLf
        strTemp = strTemp + "Windows Version Number = " + CStr(OSInfo.dwMajorVersion) + "." _
        + CStr(OSInfo.dwMinorVersion) + "." + CStr(OSInfo.dwBuildNumber) + vbCrLf
    Else
        strTemp = "Unable to get version information. GetVersionEx returned " + lngL + vbCrLf
    End If
    
    GetWinVer = strTemp

End Function

Private Function GetWinDir() As String

    Dim boolRetVal As Boolean
    Dim lpBuffer As String
    Dim nSize As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    boolRetVal = GetWindowsDirectory(lpBuffer, nSize)
    
    GetWinDir = "Windows Directory = " + StripNullTerminator(lpBuffer) + vbCrLf

End Function

Private Function GetSysDir() As String

    Dim boolRetVal As Boolean
    Dim lpBuffer As String
    Dim nSize As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    boolRetVal = GetSystemDirectory(lpBuffer, nSize)
    
    GetSysDir = "System Directory = " + StripNullTerminator(lpBuffer) + vbCrLf

End Function

Private Function GetCompName() As String

    Dim boolRetVal As Boolean
    Dim lpBuffer As String
    Dim nSize As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    boolRetVal = GetComputerName(lpBuffer, nSize)
    
    GetCompName = "Computer Name = " + StripNullTerminator(lpBuffer) + vbCrLf
    
End Function

Private Function GetDomainName() As String

    Dim lpBuffer As String
    Dim nSize As Long
    Dim lngRetVal As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    lngRetVal = GetEnvironmentVariable("USERDOMAIN", lpBuffer, nSize)
    
    GetDomainName = "Domain Name = " + StripNullTerminator(lpBuffer) + vbCrLf
    
End Function

Private Function GetDriveInfo(strDrive As String) As String

    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    Dim lpRetVal As Long
    Dim strDriveType As String
    Dim lpBuffer As String
    Dim nSize As Long
    Dim lngL As Long
    
    Dim lpBytesPerCluster As Long
    Dim lpDriveSize As Long
    Dim lpDriveFreeSpace As Long
      
    lpRetVal = GetDriveType(strDrive)
    Select Case lpRetVal
        Case DRIVE_UNKNOWN
            strDriveType = "Drive type unknown"
        Case DRIVE_NO_ROOT_DIR
            strDriveType = "Drive has no root directory"
        Case DRIVE_REMOVABLE
            strDriveType = "Floppy / removable drive"
        Case DRIVE_FIXED
            strDriveType = "Fixed hard drive"
        Case DRIVE_REMOTE
            strDriveType = "Network drive: "
            'now get the mapped network drive
            lpBuffer = SPACE(255)
            nSize = 254
            lngL = WNetGetConnection(Left(strDrive, 2), lpBuffer, nSize)
            strDriveType = strDriveType + StripNullTerminator(lpBuffer)
        Case DRIVE_CDROM
            strDriveType = "CD-ROM drive"
        Case DRIVE_RAMDISK
            strDriveType = "RAM disk"
        Case Else
            strDriveType = "Unknown device"
    End Select

    lpRetVal = GetDiskFreeSpace(strDrive, lpSectorsPerCluster, lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
    lpBytesPerCluster = lpBytesPerSector * lpSectorsPerCluster
    lpDriveSize = lpBytesPerCluster * (lpTotalNumberOfClusters / 1024) / 1024
    lpDriveFreeSpace = lpBytesPerCluster * (lpNumberOfFreeClusters / 1024) / 1024

    If lpRetVal = 1 And lpDriveSize > 0 Then
        GetDriveInfo = strDrive + " drive - " + strDriveType + vbCrLf _
                    + vbTab + "Drive Size = " + CStr(lpDriveSize) + " MB" + vbCrLf _
                    + vbTab + "Free Space = " + CStr(lpDriveFreeSpace) + " MB" + vbCrLf + vbCrLf
    Else
        GetDriveInfo = ""
    End If

End Function

Private Function GetLogonServer() As String

    Dim lpBuffer As String
    Dim nSize As Long
    Dim lngRetVal As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    lngRetVal = GetEnvironmentVariable("LOGONSERVER", lpBuffer, nSize)
    
    GetLogonServer = "Logon Server Name = " + StripNullTerminator(lpBuffer) + vbCrLf
    
End Function


Private Function GetMemoryInfo() As String

    Dim msMemory As MEMORYSTATUS
    Dim lngTotalPhys As Long
    Dim lngAvailPhys As Long
    Dim lngTotalPageFile As Long
    Dim lngAvailPageFile As Long
    Dim lngTotalVirtual As Long
    Dim lngAvailVirtual As Long
    
    'GlobalMemoryStatus msMemory
    lngTotalPhys = msMemory.dwTotalPhys / 1024 * (1 / 1024)
    lngAvailPhys = msMemory.dwAvailPhys / 1024 * (1 / 1024)
    lngTotalPageFile = msMemory.dwTotalPageFile / 1024 * (1 / 1024)
    lngAvailPageFile = msMemory.dwAvailPageFile / 1024 * (1 / 1024)
    lngTotalVirtual = msMemory.dwTotalPageFile / 1024 * (1 / 1024)
    lngAvailVirtual = msMemory.dwAvailPageFile / 1024 * (1 / 1024)
    
    GetMemoryInfo = "Memory Status:" + vbCrLf _
                    + vbTab + "Total RAM = " + CStr(lngTotalPhys) + "MB" + vbCrLf _
                    + vbTab + "Available RAM = " + CStr(lngAvailPhys) + "MB" + vbCrLf _
                    + vbTab + "Total PageFile = " + CStr(lngTotalPageFile) + "MB" + vbCrLf _
                    + vbTab + "Available PageFile = " + CStr(lngAvailPageFile) + "MB" + vbCrLf _
                    + vbTab + "Total Virtual Memory = " + CStr(lngTotalVirtual) + "MB" + vbCrLf _
                    + vbTab + "Available Virtual Memory = " + CStr(lngAvailVirtual) + "MB" + vbCrLf

End Function

Private Function GetTimeSinceReboot()

    'Returns the time since the machine was last restarted in format h:m:s
    Dim h As Long
    Dim M As Long
    Dim S As Long
    Dim l As Long
        
    l = GetTickCount()
    'GetTickCount returns number of milliseconds since last restart. divide by 1000 for seconds and convert to hours
    l = l / 1000
    'use integer division so we don't get rounding problems
    h = l \ 3600
    'Number of minutes over the hour
    M = (l - (h * 3600)) \ 60
    'Number of seconds over the minute
    S = l - (h * 3600 + M * 60)
    GetTimeSinceReboot = "Hours since reboot = " + Format(h, "00") + ":" & Format(M, "00") _
                        + ":" + Format(S, "00") + vbCrLf

End Function

Private Function GetUName() As String

    Dim lngRetVal As Long
    Dim lpBuffer As String
    Dim nSize As Long
    
    lpBuffer = SPACE(255)
    nSize = 254
    lngRetVal = GetUserName(lpBuffer, nSize)

    GetUName = "User Name = " + StripNullTerminator(lpBuffer) + vbCrLf

End Function

Private Function StripNullTerminator(lpBuffer As String) As String

    Dim i As Integer

    For i = 1 To 255
        If Asc(Mid(lpBuffer, i, 1)) = 0 Then
            lpBuffer = Left(lpBuffer, i - 1)
            Exit For
        End If
    Next i
    
    StripNullTerminator = lpBuffer

End Function

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function








Private Sub Form_Load()
SCL = 180
Label1.Caption = "Welcome to ILDEREMI Console 2008-2009." & vbCrLf & "Use HELP command for more details..." & vbCrLf & vbCrLf & ">_"
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


Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Top < 10 Then
    Label1.Top = Label1.Top + SCL
    Label1.Height = Label1.Height - SCL
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Top = Label1.Top - SCL
Label1.Height = Label1.Height + SCL
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 34 Then
    Label1.Top = Label1.Top - SCL
    Label1.Height = Label1.Height + SCL
ElseIf KeyCode = 33 And Label1.Top < 10 Then
    Label1.Top = Label1.Top + SCL
    Label1.Height = Label1.Height - SCL
End If
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
Dim a, d, t, jk, mData1, mData2
HM:
On Error Resume Next
Dim MyCommand, MyData, strTemp As String
'================='>'======================
a = 0
Do
    d = a
    a = InStr(a + 1, Label1.Caption, ">")
Loop While (a > 0)
'==========================================

'==========================================
If KeyAscii = 13 Then
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1) & Chr(KeyAscii) & ">_"
ElseIf KeyAscii = 8 And Right(Label1.Caption, 2) <> ">_" Then
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & "_"
ElseIf KeyAscii = 27 And Right(Label1.Caption, 2) <> ">_" Then
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1) & "_"
ElseIf KeyAscii <> 8 And KeyAscii <> 27 Then
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1) & Chr(KeyAscii) & "_"
End If
'==========================================

'==========================================
MyCommand = Right(Label1.Caption, Len(Label1.Caption) - d)
For t = 1 To Len(MyCommand)
    If Mid(MyCommand, t, 1) = " " Then
        MyData = Mid(MyCommand, t + 1, Len(MyCommand) - t - 3)
        MyCommand = Left(MyCommand, t - 1)
        Exit For
    End If
Next
MyPData = MyData
If Asc(Right(MyData, 1)) = 13 Then
    MyData = Left(MyData, Len(MyData) - 1)
End If
'==========================================
Label2.Caption = "Command:" & UCase(MyCommand) & vbCrLf & "Data:" & MyPData
'==========================================
If KeyAscii = 13 Then
Select Case UCase(MyCommand)
   
    Case "ABOUT" & Chr(13) & ">_"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & _
                    "   This product is created by Masoud iLDEREMi." & vbCrLf & _
                    "       mailderemi@gmail.com - +98 915 109 2841" & vbCrLf & _
                    "================================================" & vbCrLf & _
                    "Version " & App.Major & App.Minor & App.Revision & vbCrLf & _
                    "                      2008-2009" & vbCrLf & vbCrLf & ">_"
    Case "EXIT" & Chr(13) & ">_"
    Close
    End
    Case "QUIT" & Chr(13) & ">_"
    Close
    End
    
    Case "CLS" & Chr(13) & ">_"
    Label1.Caption = ">_"
    Label1.Top = 10
    Label1.Height = 5490
    
    Case "SCRSVR" & Chr(13) & ">_"
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
    
    Case "TASKOPT" & Chr(13) & ">_"
    Shell "rundll32.exe advapi32.dll,ProcessIdleTasks"
    Label2.Caption = Label2.Caption & vbCrLf & "Open:" & "rundll32.exe advapi32.dll,ProcessIdleTasks"
    
    Case "CMD"
    strTemp = getm.GetCommandOutput("cmd /c " & MyData, True, True, True)
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & _
    strTemp & vbCrLf & vbCrLf & ">_"
    Label2.Caption = Label2.Caption & vbCrLf & "CMD:" & MyData
    
    Case "DEBUGIT" & Chr(13) & ">_"
    fmDebug.Visible = True
    
    Case "-DEBUGIT" & Chr(13) & ">_"
    fmDebug.Visible = False
    
    '+++++++++++++++++++++
    Case "SCANPORT"
    If Trim(MyData) = "" Then
        Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & "You must enter an IP address!" & vbCrLf & ">_"
    Else
        fmPortScan.Visible = True
        List1.clear
        Timer1.interval = 1
        Timer1.Enabled = True
    End If
    
    Case "SCANPORT" & Chr(13) & ">_"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & "You must enter an IP address!" & vbCrLf & ">_"
    
    Case "-SCANPORT" & Chr(13) & ">_"
    If Timer1.Enabled = True Then
        strPort = 0
        Timer1.Enabled = False
        fmPortScan.Visible = False
    End If
    '+++++++++++++++++++++++
    
    Case "WINTALPHA"
    For jk = 1 To Len(MyData)
        If Mid(MyData, jk, 1) = " " Then
            mData1 = Left(MyData, jk - 1)
            mData2 = Right(MyData, Len(MyData) - Len(mData1) - 1)
            Exit For
        End If
    Next
    Label2.Caption = Label2.Caption & vbCrLf & "Data1:" & mData1 & vbCrLf & "Data2:" & mData2
    SetWindowLong FindWindow(vbNullString, mData1), GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FindWindow(vbNullString, mData1), 0, mData2, LWA_ALPHA

    Case "MSGBOX"
    For jk = Len(MyData) To 1 Step -1
        If Mid(MyData, jk, 1) = Chr(34) Then
            mData1 = Right(MyData, Len(MyData) - jk - 1)
            Exit For
        End If
    Next
    Label2.Caption = Label2.Caption & vbCrLf & "Data1:" & mData1
    MsgBox Parser(MyData, Chr(34), 1), , mData1

    Case "CPL"
    Shell "RUNDLL32 SHELL32.DLL,Control_RunDLL " & MyData & ",@0,0"
    Label2.Caption = Label2.Caption & vbCrLf & "Open:" & "RUNDLL32 SHELL32.DLL,Control_RunDLL " & MyData & ",@0,0"
    
    Case "DETAILS" & Chr(13) & ">_"
    fmDetails.Visible = True
    Timer2.Enabled = True
    
    Case "-DETAILS" & Chr(13) & ">_"
    fmDetails.Visible = False
    Timer2.Enabled = False


    Case "DET.PIC"
    Picture2.Picture = LoadPicture(MyData)
    
    Case "DET.LHOST" & Chr(13) & ">_"
    Label9.Visible = True
    
    Case "DET.LIP" & Chr(13) & ">_"
    Label8.Visible = True
    
    Case "DET.DATE" & Chr(13) & ">_"
    Label6.Visible = True
    
    Case "DET.TIME" & Chr(13) & ">_"
    Label5.Visible = True
    
    Case "DET.SCREEN" & Chr(13) & ">_"
    Label7.Visible = True
    
    
    Case "-DET.PIC" & Chr(13) & ">_"
    Picture2.Picture = Nothing
    
    Case "-DET.LHOST" & Chr(13) & ">_"
    Label9.Visible = False
    
    Case "-DET.LIP" & Chr(13) & ">_"
    Label8.Visible = False
    
    Case "-DET.DATE" & Chr(13) & ">_"
    Label6.Visible = False
    
    Case "-DET.TIME" & Chr(13) & ">_"
    Label5.Visible = False
    
    Case "-DET.SCREEN" & Chr(13) & ">_"
    Label7.Visible = False
    
    Case "BACKPIC"
    Picture1.Picture = LoadPicture(MyData)
    
    Case "-BACKPIC" & Chr(13) & ">_"
    Picture1.Picture = Nothing
    
    Case "CFILE"
    Close
    Dim stu As Byte
    strTemp = "unsigned char data[" & FileLen(MyData) & "] = {" & vbCrLf
    Open MyData For Binary As #1
        For jk = 1 To FileLen(MyData)
            Get #1, jk, stu
            strTemp = strTemp & "0x" & Hex(stu) & ", "
        Next
    Close
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & _
                    Left(strTemp, Len(strTemp) - 2) & vbCrLf & "};" & vbCrLf & ">_"
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB

    Case "PASFILE"
    Close
    Dim stu2 As Byte
    strTemp = "data: array[0.." & FileLen(MyData) & "] of byte = (" & vbCrLf
    Open MyData For Binary As #1
        For jk = 1 To FileLen(MyData)
            Get #1, jk, stu2
            strTemp = strTemp & "$" & Hex(stu2) & ", "
        Next
    Close
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & _
                    Left(strTemp, Len(strTemp) - 2) & vbCrLf & ");" & vbCrLf & ">_"
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
    
'---------------------------
    Case "VBFILE"
    Close
    Dim stu3 As Byte
    strTemp = "MData = Array("
    Open MyData For Binary As #1
        For jk = 1 To FileLen(MyData)
            Get #1, jk, stu3
            strTemp = strTemp & "&&" & "H" & Hex(stu3) & ", "
        Next
    Close
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & _
                    Left(strTemp, Len(strTemp) - 2) & ")" & vbCrLf & ">_"
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
'---------------------------
                    
    Case "SYSEXE"
    Dim mmin As Integer
    Dim hhor As Integer
    hhor = Val(Left(Time$, 2))
    mmin = Val((Mid(Time$, 4, 2))) + 1
    If mmin > 59 Then
        mmin = mmin - 60
        hhor = hhor + 1
    End If
    If hhor > 23 Then hhor = hhor - 24
    Shell "cmd /c at " & hhor & ":" & mmin & " /interactive " & MyData, vbHide
    
    Case "-SYSEXE" & Chr(13) & ">_"
    Shell "cmd /c at /delete /yes"
    
    Case "ASCII"
    For jk = 1 To Len(MyData)
        strTemp = strTemp & Asc(Mid(MyData, jk, 1)) & ","
    Next
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & Left(strTemp, Len(strTemp) - 1) & vbCrLf & ">_"

    Case "PASTE" & Chr(13) & ">_"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & Clipboard.GetText & vbCrLf & ">_"
    
    Case "PASTECMD" & Chr(13) & ">_"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & ">" & Clipboard.GetText & vbCrLf
    GoTo HM
    
    Case "PCINFO" & Chr(13) & ">_"
    Dim i As Integer

    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetUName
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetCompName
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetNetworked
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetDomainName
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetLogonServer
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetTimeSinceReboot
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetLastBootState
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & vbCrLf
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetWinVer
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetWinDir
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetSysDir
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetMemoryInfo
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & vbCrLf
    For i = 1 To 26
        Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetDriveInfo(Chr(Asc("A") + i - 1) + ":\")
    Next i
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & ">_"
    
    Case "SYSINFO" & Chr(13) & ">_"
    strTemp = getm.GetCommandOutput("cmd /c systeminfo", True, True, True)
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & _
    strTemp & vbCrLf & vbCrLf & ">_"
    
    Case "BEEP" & Chr(13) & ">_"
    Beep
    
    Case "SAVECMD"
    Close
    Open MyData For Output As #10
        Print #10, Label1.Caption
    Close
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
    
    Case "CMDFONT"
    Label1.Font = MyData
    
    Case "CMDFONT.B" & Chr(13) & ">_"
    Label1.FontBold = True
    
    Case "CMDFONT.I" & Chr(13) & ">_"
    Label1.FontItalic = True
    
    Case "CMDFONT.U" & Chr(13) & ">_"
    Label1.FontUnderline = True
    
    Case "-CMDFONT.B" & Chr(13) & ">_"
    Label1.FontBold = False
    
    Case "-CMDFONT.I" & Chr(13) & ">_"
    Label1.FontItalic = False
    
    Case "-CMDFONT.U" & Chr(13) & ">_"
    Label1.FontUnderline = False
    
    Case "CMD" & Chr(13) & ">_"
    Shell "cmd /k", vbNormalFocus
    
    Case "FILEDATE"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & FileDateTime(MyData) & vbCrLf & ">_"
    
    Case "FILELEN"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & FileLen(MyData) & vbCrLf & ">_"
    
    Case "FILEATTR"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & GetAttr(MyData) & vbCrLf & ">_"
    
    Case "REGMAIN" & Chr(13) & ">_"
    fmMainPass.Visible = True
    Timer4.Enabled = True
    Text1.SetFocus
    
    'Case "PLAY"
    'sndPlaySound (MyData), &H80
    
    Case "FASTSHUTDOWN" & Chr(13) & ">_"
    Shell "WMIC OS Where Primary=TRUE Call Shutdown"
    
    Case "FASTREBOOT" & Chr(13) & ">_"
    Shell "WMIC OS Where Primary=TRUE Call Reboot"
    
    Case "DEFAULT" & Chr(13) & ">_"
    Label1.Font = "Lucida Console"
    
    Case "HELP" & Chr(13) & ">_"
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2) & vbCrLf & _
                    "ABOUT" & vbCrLf & "EXIT/QUIT" & vbCrLf & "CLS" & vbCrLf & _
                    "SCRSVR" & vbCrLf & "TASKOPT" & vbCrLf & "CMD" & vbCrLf & _
                    "DEBUGIT" & vbCrLf & "SCANPORT" & vbCrLf & "WINTALPHA" & vbCrLf & _
                    "MSGBOX" & vbCrLf & "CPL" & vbCrLf & "DETAILS" & vbCrLf & _
                    "DET.PIC" & vbCrLf & "DET.DATE" & vbCrLf & "DET.TIME" & vbCrLf & _
                    "DET.SCREEN" & vbCrLf & "DET.LIP" & vbCrLf & "DET.LHOST" & vbCrLf & _
                    "BACKPIC" & vbCrLf & "CFILE" & vbCrLf & "PASFILE" & vbCrLf & "VBFILE" & vbCrLf & "SYSEXE" & vbCrLf & _
                    "PASTE" & vbCrLf & "PASTECMD" & vbCrLf & "PCINFO" & vbCrLf & "SYSINFO" & vbCrLf & _
                    "BEEP" & vbCrLf & "SAVECMD" & vbCrLf & "CMDFONT" & vbCrLf & _
                    "CMDFONT.B" & vbCrLf & "CMDFONT.I" & vbCrLf & "CMDFONT.U" & vbCrLf & ">_"
End Select
End If
'==========================================
End Sub

Private Sub Text1_Change()
Label12 = String(Len(Text1), "l")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1 = "masoudilderemi136" Then
    sPass = 0
    fmMainPass.Visible = False
    Timer4.Enabled = False
    frmMain.Label8.Caption = "Output is Enable"
    frmMain.Label21.Caption = "Output is Enable"
    Me.Hide
    frmMain.Show
ElseIf KeyAscii = vbKeyEscape Then
    fmMainPass.Visible = False
    Timer4.Enabled = False
    sPass = 0
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Winsock1.Close
strPort = strPort + 1
Winsock1.RemoteHost = MyPData
Winsock1.RemotePort = strPort
Winsock1.Connect
End Sub

Private Sub Timer2_Timer()
Label9.Caption = Winsock2.LocalHostName
Label8.Caption = Winsock2.LocalIP
Label7.Caption = Screen.Width & "x" & Screen.Height
Label6.Caption = Date
Label5.Caption = Time$
End Sub

Private Sub Timer3_Timer()
bp = bp + 1
If bp = 2 Then
    Picture1.Picture = Me.Picture
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
sPass = sPass + 1
If sPass <= Len("Enter password to enable output for Main form._") Then
    Label11.Caption = Mid("Enter password to enable output for Main form", 1, sPass) & "_"
Else
    Label11.Caption = "Enter password to enable output for Main form"
End If
End Sub

Private Sub Winsock1_Connect()
List1.AddItem (Winsock1.RemotePort & " is open")
End Sub

Function Parser(ByVal text As String, ByVal Toker As String, ByVal TokenNumber As Long) As String
    Dim ft, j
    'Declarations
    Dim token, TokerFirst, TokerLast As String
    Dim Count As Integer
    'Equalities
    TokerFirst = Left(Toker, 1)
    TokerLast = Right(Toker, 1)
    Count = 0
    token = ""
    'Finding the first char of token
    For ft = 1 To Len(text)
        If Mid(text, ft, 1) = TokerLast Then Count = Count + 1
        If Count = TokenNumber Then Exit For
    Next
    'Adding token to Token
    For j = ft + 1 To Len(text)
        If Mid(text, j, 1) = TokerFirst Then Exit For
        token = token & Mid(text, j, 1)
    Next
    'Returning Token
    Parser = token
End Function
