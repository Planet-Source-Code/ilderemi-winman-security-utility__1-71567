VERSION 5.00
Begin VB.Form frmProcess 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7815
   ClientLeft      =   2220
   ClientTop       =   2970
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4200
      Width           =   7095
   End
   Begin VB.TextBox txtFullProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   6135
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9720
      TabIndex        =   2
      Top             =   6510
      Width           =   2655
   End
   Begin VB.ListBox lstProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3540
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   3480
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3150
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Back to Main"
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
      Height          =   195
      Left            =   7680
      TabIndex        =   12
      Top             =   7260
      Width           =   4695
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   7680
      Top             =   7250
      Width           =   4695
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Left            =   9705
      Top             =   6495
      Width           =   2685
   End
   Begin VB.Label isButton4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   8640
      TabIndex        =   11
      Top             =   6555
      Width           =   975
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   8640
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label isButton3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Show All"
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
      Left            =   7680
      TabIndex        =   10
      Top             =   6555
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   7680
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label isButton27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copy name"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   7280
      Width           =   1215
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
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
      Left            =   240
      TabIndex        =   8
      Top             =   7280
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   1560
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label isButton1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   3555
      Width           =   1575
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   1920
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label isButton2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Analyse"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3550
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   240
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Left            =   120
      Top             =   7080
      Width           =   7335
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7320
      Width           =   4455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   6855
      Left            =   7560
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   3855
      Index           =   1
      Left            =   3840
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   3855
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   7815
      Left            =   0
      Top             =   0
      Width           =   12615
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   2895
      Left            =   120
      Top             =   4080
      Width           =   7335
   End
End
Attribute VB_Name = "frmProcess"
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
Dim B As Integer 'b => timer
Dim Processing As Long
Dim jkj As Integer
Private Sub EnumProcess(Optional ByVal sExeName As String = "")
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    
    ' Create snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        ' Clear list
        lstProcess.clear
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        
        ' Find first process
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            ' Found specified process
            If sExeName = Left$(tPE.szExeFile, Len(sExeName)) And Len(sExeName) > 0 Then
                Dim lProcess As Long
                Dim lExitCode As Long
                ' Open process
                lProcess = OpenProcess(0, False, tPE.th32ProcessID)
                ' Terminate process
                TerminateProcess lProcess, lExitCode
                ' Close handle
                CloseHandle lProcess
            Else
                ' Add exe to list
                lstProcess.AddItem tPE.szExeFile
            End If
            ' Get next process
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        
        ' Close handle
        CloseHandle (lSnapShot)
        
    Else
        lstProcess.AddItem "Cannot enumerate running process!"
    End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Shell "taskkill /f /im " & Chr(34) & Text1 & Chr(34), vbHide
End Sub

Private Sub Form_Load()
isButton1_Click
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
jkj = jkj + 1
If jkj >= 3 Then
    jkj = 0
    Timer.Enabled = False
    Exit Sub
End If
EnumProcess
List1.clear
For a = 0 To lstProcess.ListCount - 1
    If UCase(lstProcess.List(a)) <> UCase("[System Process]") And UCase(lstProcess.List(a)) <> UCase("System") And UCase(lstProcess.List(a)) <> UCase("smss.exe") And UCase(lstProcess.List(a)) <> UCase("csrss.exe") And UCase(lstProcess.List(a)) <> UCase("winlogon.exe") And UCase(lstProcess.List(a)) <> UCase("services.exe") And UCase(lstProcess.List(a)) <> UCase("lsass.exe") And UCase(lstProcess.List(a)) <> UCase("svchost.exe") And UCase(lstProcess.List(a)) <> UCase("explander.exe") Then
        List1.AddItem lstProcess.List(a)
    End If
Next
spr = getm.GetCommandOutput("cmd /c tasklist /m", True, False, False)
txtFullProcess.Text = Right(spr, Len(spr) - 162)
End Sub

Private Sub isButton2_Click()
Text2 = "«ê— »—‰«„Â Â«ÌÌ òÂ œ— Å«ÌÌ‰ ·Ì”  „Ì ‘Ê‰œ œ— ”Ì” „ ‘„« ‰’» ‰»«‘‰œ «Ì‰ «Õ „«· ÊÃÊœ œ«—œ òÂ ‰Ê⁄Ì ›«Ì· „Œ—» »«‘‰œ. «·» Â œ— »⁄÷Ì «“ „Ê«—œ œ— ’Ê—  «Ã—« »Êœ‰ »—‰«„Â Ê €Ì— ﬁ«»· —ÊÌ  »Êœ‰ ¬‰ «Ì‰ «Õ „«· ÊÃÊœ œ«—œ òÂ ¬‰ ›«Ì· „‘òÊò »«‘œ..." & vbCrLf & "------------------------------------------------------"
Timer.Enabled = True
End Sub

Private Sub isButton27_Click()
Clipboard.SetText Text1
End Sub

Private Sub isButton3_Click()
On Error Resume Next
spr = getm.GetCommandOutput("tasklist /m", True, False, False)
txtFullProcess.Text = Right(spr, Len(spr) - 162)
End Sub

Private Sub isButton4_Click()
'winrnr.dll"
spr = getm.GetCommandOutput("tasklist /fi " & Chr(34) & "MODULES eq " & Text3 & Chr(34), True, False, False)
If spr <> "" Then txtFullProcess.Text = Right(spr, Len(spr) - 147)
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub List1_Click()
On Error Resume Next
Text1 = List1.List(List1.ListIndex)
End Sub

Private Sub lstProcess_Click()
Text1 = lstProcess.List(lstProcess.ListIndex)
End Sub

Private Sub Timer_Timer()
B = B + 1
Select Case LCase(lstProcess.List(B))
    Case "nod32kui.exe"
    Text2 = Text2 & vbCrLf & "nod32kui.exe „—»Êÿ »Â ‰—„ «›“«— «„‰Ì Ì NOD32 „Ì »«‘œ..."
    
    Case "mdm.exe"
    Text2 = Text2 & vbCrLf & "mdm.exe ÌòÌ «“ ›«Ì· Â«Ì ÊÌ‰œÊ“ (Machine Debug Manager)..."
    
    Case "alg.exe"
    Text2 = Text2 & vbCrLf & "alg.exe ÌòÌ «“ ›«Ì· Â«Ì ÊÌ‰œÊ“ (Application Layer Gateway Service)..."
    
    Case "nvsvc64.exe"
    Text2 = Text2 & vbCrLf & "nvsvc64.exe „—»Êÿ »Â œ—«ÌÊ— ò«—  ê—«›Ìò (NVIDIA Driver Helper Service)..."
    
    Case "cisvc.exe"
    Text2 = Text2 & vbCrLf & "cisvc.exe ÌòÌ «“ ›«Ì· Â«Ì ÊÌ‰œÊ“ (Content Index service)..."
    
    Case "wscntfy.exe"
    Text2 = Text2 & vbCrLf & "wscntfy.exe ÌòÌ «“ ›«Ì· Â«Ì..."
    
    Case "ctfmon.exe"
    Text2 = Text2 & vbCrLf & "ctfmon.exe ÌòÌ «“ ›«Ì· Â«Ì ÊÌ‰œÊ“ òÂ «»“«—  €ÌÌ— “»«‰ ‰Ì«“ »Â «Ì‰ ›«Ì· œ«—œ (CTF Loader)..."
    
    Case "csrss.exe"
    Text2 = Text2 & vbCrLf & "csrss.exe ÌòÌ «“ ›«Ì· Â«Ì «’·Ì ÊÌ‰œÊ“ (Client Server Runtime Process)..."
    
    Case "explorer.exe"
    Text2 = Text2 & vbCrLf & "explorer.exe ÌòÌ «“ ›«Ì· Â«Ì «’·Ì ÊÌ‰œÊ“..."
    
    Case "vb6.exe"
    Text2 = Text2 & vbCrLf & "VB6.exe „ÕÌÿ »—‰«„Â ‰ÊÌ”Ì Visual Basic 6 ..."
    
    Case "spoolsv.exe"
    Text2 = Text2 & vbCrLf & "spoolsv.exe ÌòÌ «“ »—‰«„Â Â«Ì ÊÌ‰œÊ“ òÂ ç«Åê— ‰Ì«“ »Â «Ì‰ ›«Ì· œ«—œ (Spooler SubSystem App)..."
    
    Case "inetinfo.exe"
    Text2 = Text2 & vbCrLf & "inetinfo.exe ”—ÊÌ” IIS œ— ÊÌ‰œÊ“ (Internet Information Services)..."
    
    Case "cidaemon.exe"
    Text2 = Text2 & vbCrLf & "cidaemon.exe ÌòÌ «“ ›«Ì· Â«Ì ÊÌ‰œÊ“ (Indexing Service filter daemon)..."
    
    Case "babylon.exe"
    Text2 = Text2 & vbCrLf & "babylon.exe „—»Êÿ »Â ‰—„ «›“«— œÌò‘‰—Ì Babylon..."
    
    Case "windows explorer.exe"
    Text2 = Text2 & vbCrLf & "›«Ì· windows explorer.exe „‘òÊò »Â ÊÌ—Ê” New Folder..."
    
    Case "eksplorer.exe"
    Text2 = Text2 & vbCrLf & "›«Ì· eksplorer.exe „‘òÊò »Â ÌòÌ «“ ÊÌ—Ê” Â«Ì —Ê”Ì «” ..."
    
    Case "fun.xls.exe"
    Text2 = Text2 & vbCrLf & "›«Ì· fun.xls.exe „‘òÊò »Â ‰ê«—‘ œÊ„ «“ ÊÌ—Ê” çÌ‰Ì MSFUN80 «” ..."
    
    Case "nwiz.exe"
    Text2 = Text2 & vbCrLf & "nwiz.exe „—»Êÿ »Â œ—«ÌÊ— ò«—  ê—«›Ìò (NVIDIA nView Wizard)..."
    
    Case "msmsgs.exe"
    Text2 = Text2 & vbCrLf & "msmsgs.exe „—»Êÿ »Â »—‰«„Â MSN Messenger..."
    
    Case "dap.exe"
    Text2 = Text2 & vbCrLf & "dap.exe „—»Êÿ »Â »—‰«„Â Download Accelerator Plus"
    
    Case "cmd.exe"
    Text2 = Text2 & vbCrLf & "cmd.exe „—»Êÿ »Â »—‰«„Â Command Prompt"
    
    Case "regedit.exe"
    Text2 = Text2 & vbCrLf & "regedit.exe ÌòÌ «“ ›«Ì· Â«Ì «’·Ì ÊÌ‰œÊ“ Regedit"
    
    Case "calc.exe"
    Text2 = Text2 & vbCrLf & "calc.exe „«‘Ì‰ Õ”«» ÊÌ‰œÊ“"
    
    Case "charmap.exe"
    Text2 = Text2 & vbCrLf & "charmap.exe ÌòÌ «“ »—‰«„Â Â«Ì ÊÌ‰œÊ“ (Character Map)"
    
    Case "ccapp.exe"
    Text2 = Text2 & vbCrLf & "ccapp.exe „—»Êÿ »Â ‰—„ «›“«— «„‰Ì Ì Norton"
    
    Case "egui.exe"
    Text2 = Text2 & vbCrLf & "egui.exe „—»Êÿ »Â ‰—„ «›“«— «„‰Ì Ì NOD32 „Ì »«‘œ..."
    
    Case "ekrn.exe"
    Text2 = Text2 & vbCrLf & "ekrn.exe „—»Êÿ »Â ‰—„ «›“«— «„‰Ì Ì NOD32 „Ì »«‘œ..."
    
    Case "iexplore.exe"
    Text2 = Text2 & vbCrLf & "iexplore.exe „—»Êÿ »Â „—Ê—ê— Internet Explorer „Ì »«‘œ..."
    
    Case "opera.exe"
    Text2 = Text2 & vbCrLf & "opera.exe „—»Êÿ »Â „—Ê—ê— Opera „Ì »«‘œ..."
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
    
    Case ""
    Text2 = Text2 & vbCrLf & ""
End Select
If B >= lstProcess.ListCount - 1 Then
Timer.Enabled = False
B = 0
End If
'text2
End Sub
