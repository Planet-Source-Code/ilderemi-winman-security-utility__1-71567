VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMDB2000 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
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
   ScaleHeight     =   1335
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   490
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browse"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   130
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   3840
      Top             =   105
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   830
      Top             =   480
      Width           =   3765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3610
      TabIndex        =   2
      Top             =   915
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   3610
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   825
      Top             =   105
      Width           =   2925
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMDB2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NullDate As Date
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const OFS_MAXPATHNAME = 128
Private Const OF_READ = &H0
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hfile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function GetFileDate(File As String) As Date
Dim fhi As BY_HANDLE_FILE_INFORMATION
Dim ctime As FILETIME, atime As FILETIME, wtime As FILETIME
Dim ftime As SYSTEMTIME
Dim buff As OFSTRUCT
Dim rval As Long, hfile As Long
    hfile = OpenFile(File, buff, OF_READ)
    If hfile = -1 Then
        GetFileDate = NullDate
    Else
        GetFileInformationByHandle hfile, fhi
        ctime = fhi.ftCreationTime

        rval = FileTimeToLocalFileTime(ctime, ctime)

        rval = FileTimeToSystemTime(ctime, ftime)
        GetFileDate = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
    End If
    CloseHandle hfile
End Function

Public Function GuessAccess2000Password(ProtectedFile As String) As String
Dim n As Long, s1 As String * 1, S2 As String * 1
Dim Password As String
Dim x1 As Byte, x2 As Byte
Dim TempFile As String
Dim DateFile As Date, PreviousDate As Date
Dim Handle1 As Long, Handle2 As Long

    DateFile = GetFileDate(ProtectedFile)
    If DateFile = NullDate Then
        GuessAccess2000Password = "Can't open database file. Maybe you have it open in exclusive mode"
        Exit Function
    End If
    
    TempFile = TempPath & "temp.mdb"

    If Dir(TempFile) <> "" Then
        Kill TempFile
    End If
    
    PreviousDate = Date
    Date = DateFile

    CreateDatabase TempFile, dbLangGeneral
    
    Date = PreviousDate
    
    Handle1 = FreeFile
    Open TempFile For Binary As #Handle1
    Handle2 = FreeFile
    Open ProtectedFile For Binary As #Handle2
    Password = ""
    Seek #Handle1, &H43
    Seek #Handle2, &H43

    For n = 0 To 19
        x1 = Asc(Input(1, Handle1))
        x2 = Asc(Input(1, Handle2))
        If x1 <> x2 Then
          Password = Password & Chr(x1 Xor x2)
        End If
        x1 = Asc(Input(1, Handle1))
        x2 = Asc(Input(1, Handle2))
    Next
    Close 1
    Close 2
    'Kill TempFile
    GuessAccess2000Password = Password
End Function
Private Function TempPath() As String

Dim Path As String
Dim TempFile As String
Dim slength As Long
Dim lastfour As Long

    Path = SPACE(255)
    slength = GetTempPath(255, Path)
    TempPath = Left(Path, slength)

End Function

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

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label5_Click()
    CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    Text2 = "The password is: " & GuessAccess2000Password(CommonDialog1.Filename)
End Sub
