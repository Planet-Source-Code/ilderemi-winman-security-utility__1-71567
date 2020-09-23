Attribute VB_Name = "getm"
Option Explicit


Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100


Private Const SW_HIDE = 0


Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2


Private Const ERROR_BROKEN_PIPE = 109



Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type



Private Declare Function CreatePipe Lib "kernel32" ( _
  phReadPipe As Long, _
  phWritePipe As Long, _
  lpPipeAttributes As Any, _
  ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpBuffer As Any, _
  ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, _
  lpOverlapped As Any) As Long

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpCommandLine As String, _
  lpProcessAttributes As Any, _
  lpThreadAttributes As Any, _
  ByVal bInheritHandles As Long, _
  ByVal dwCreationFlags As Long, _
  lpEnvironment As Any, _
  ByVal lpCurrentDriectory As String, _
  lpStartupInfo As STARTUPINFO, _
  lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function DuplicateHandle Lib "kernel32" ( _
  ByVal hSourceProcessHandle As Long, _
  ByVal hSourceHandle As Long, _
  ByVal hTargetProcessHandle As Long, _
  lpTargetHandle As Long, _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwOptions As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long

Private Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" ( _
  lpszSrc As Any, _
  ByVal lpszDst As String, _
  ByVal cchDstLength As Long) As Long


Public Function GetCommandOutput(sCommandLine As String, Optional fStdOut As Boolean = True, _
                                 Optional fStdErr As Boolean = False, Optional fOEMConvert As Boolean = True) As String

  Dim hPipeRead As Long, hPipeWrite1 As Long, hPipeWrite2 As Long
  Dim hCurProcess As Long
  Dim sa As SECURITY_ATTRIBUTES
  Dim si As STARTUPINFO
  Dim pi As PROCESS_INFORMATION
  Dim baOutput() As Byte
  Dim sNewOutput As String
  Dim lBytesRead As Long
  Dim fTwoHandles As Boolean
  
  Dim lRet As Long
  
  
  Const BUFSIZE = 1024
  
  
  If (Not fStdOut) And (Not fStdErr) Then Err.Raise 5         ' Invalid Procedure call or Argument
  
  
  fTwoHandles = fStdOut And fStdErr
  
  ReDim baOutput(BUFSIZE - 1) As Byte

  With sa
    .nLength = Len(sa)
    .bInheritHandle = 1    ' get inheritable pipe handles
  End With
  
  If CreatePipe(hPipeRead, hPipeWrite1, sa, BUFSIZE) = 0 Then Exit Function
  
  hCurProcess = GetCurrentProcess()
  
  
  Call DuplicateHandle(hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, _
                       0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
  
  If fTwoHandles Then
    Call DuplicateHandle(hCurProcess, hPipeWrite1, hCurProcess, hPipeWrite2, 0&, _
                         1&, DUPLICATE_SAME_ACCESS)
  End If
  
  With si
    .cb = Len(si)
    .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    .wShowWindow = SW_HIDE          ' hide the window
    
    If fTwoHandles Then
      .hStdOutput = hPipeWrite1
      .hStdError = hPipeWrite2
    ElseIf fStdOut Then
      .hStdOutput = hPipeWrite1
    Else
      .hStdError = hPipeWrite1
    End If
  End With
    
  If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, _
                   ByVal 0&, vbNullString, si, pi) Then
    
    
    Call CloseHandle(pi.hThread)
    
    
    Call CloseHandle(hPipeWrite1)
    hPipeWrite1 = 0
    If hPipeWrite2 Then
      Call CloseHandle(hPipeWrite2)
      hPipeWrite2 = 0
    End If

    Do
      
      DoEvents

      If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then Exit Do
      
      If fOEMConvert Then
        
        
        sNewOutput = String$(lBytesRead, 0)
        Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
      
      Else
        
        
        sNewOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
      End If
      
      GetCommandOutput = GetCommandOutput & sNewOutput
      
    Loop
    
    
    Call CloseHandle(pi.hProcess)
    
  End If


  Call CloseHandle(hPipeRead)
  If hPipeWrite1 Then Call CloseHandle(hPipeWrite1)
  If hPipeWrite2 Then Call CloseHandle(hPipeWrite2)
  
End Function





