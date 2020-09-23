VERSION 5.00
Begin VB.Form frmScrSvr 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   LinkTopic       =   "frmScrSvr"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer7 
      Interval        =   57
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Timer Timer15 
      Interval        =   40
      Left            =   840
      Top             =   2040
   End
   Begin VB.PictureBox picCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   4680
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   6
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Timer Timer14 
      Interval        =   1000
      Left            =   960
      Top             =   3480
   End
   Begin VB.Timer Timer13 
      Interval        =   10
      Left            =   1440
      Top             =   3000
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00008000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   4
      Top             =   6600
      Width           =   4335
   End
   Begin VB.PictureBox picVB 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   3
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   3360
      Top             =   3000
   End
   Begin VB.Timer Timer11 
      Interval        =   50
      Left            =   2880
      Top             =   2520
   End
   Begin VB.Timer Timer10 
      Interval        =   39
      Left            =   2400
      Top             =   600
   End
   Begin VB.Timer Timer9 
      Interval        =   45
      Left            =   960
      Top             =   600
   End
   Begin VB.Timer Timer8 
      Interval        =   51
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Timer Timer5 
      Interval        =   33
      Left            =   3840
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Interval        =   27
      Left            =   3360
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   21
      Left            =   2880
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   2400
      Top             =   2040
   End
   Begin VB.PictureBox digits 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   6000
      Picture         =   "frmScrSvr.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   1560
   End
   Begin VB.PictureBox picASM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox picC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
      Begin VB.Timer Timer16 
         Interval        =   1
         Left            =   3720
         Top             =   1080
      End
   End
   Begin VB.PictureBox picWatch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FontTransparent =   0   'False
      ForeColor       =   &H00008000&
      Height          =   3735
      Left            =   4560
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   5
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmScrSvr"
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

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Private Const W_OF_CHAR As Long = 15
Private Const H_OF_CHAR As Long = 18

Private Type tetrixPos
    filled As Boolean
    Color As Long
End Type

Private Type tetrixPiece
    X As Long
    Y As Long
    Color As Long
End Type

Dim colors(3) As Long
Private msgString As String

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim txtSpy As String

Private Sub Form_Click()
On Error Resume Next
    Unload frmCMDLine
    Unload frmMain
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    End
End Sub


Private Sub Form_Load()
txtSpy = "0"
    'frmLogin.Show , Me
    Dim nV As Long
    nV = Int(((Screen.Height / Screen.TwipsPerPixelY) / H_OF_CHAR) / 3)
    nV = 10
    Move 0, 0, Screen.Width, Screen.Height
    
    
    picASM.Move 0, 0, W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picC.Move 0, H_OF_CHAR * (nV + 1), W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picVB.Move 0, H_OF_CHAR * (2 * nV + 2), W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picGame.Move 0, H_OF_CHAR * (3 * nV + 3), 3 * (Screen.Width / Screen.TwipsPerPixelX) / W_OF_CHAR, 3 * (Screen.Height / Screen.TwipsPerPixelY) / H_OF_CHAR
    picWatch.Move (Screen.Width / Screen.TwipsPerPixelX) - 196, 0, 196, 388
    picCommand.Move (Screen.Width / Screen.TwipsPerPixelX) - 196, 392 + H_OF_CHAR, 196, 300
    
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
    
    msgString = _
                " (C) BY Masoud iLDEREMi, 2008" & vbCrLf & _
                "          WinMan 1.0.1       " & vbCrLf & _
                " THIS SOFTWARE IS CREATED BY " & vbCrLf & _
                vbCrLf & _
                "           VISUAL BASIC 6.0" & vbCrLf & vbCrLf & _
                " TO LOGIN, TYPE:" & vbCrLf & _
                "                ILDEREMI CMD" & vbCrLf & _
                "                      OR" & vbCrLf & _
                "                ILDEREMI MAIN" & vbCrLf & _
                " CONTACT ME:" & vbCrLf & _
                "         +98 915 109 2841" & vbCrLf & _
                "       MAILDEREMI@GMAIL.COM"
                 
    colors(0) = RGB(0, 32, 0)
    colors(1) = RGB(0, 64, 0)
    colors(2) = RGB(0, 96, 0)
    colors(3) = RGB(0, 128, 0)
                
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Close ffASM
    Close ffC
    Close ffVB
    Close
End Sub

Private Sub picASM_Click()
    End
End Sub

Private Sub Timer1_Timer()
    Static started As Boolean
    Dim i As Long
    If Not started Then
        picVB.Visible = False
        picC.Visible = False
        picASM.Visible = False
        picWatch.Visible = False
        picGame.Visible = False
        picCommand.Visible = False
        For i = 0 To 3000
            printNumber Int(Rnd * 999999999)
        Next i
        started = True
        picVB.Visible = True
        picC.Visible = True
        picASM.Visible = True
        picWatch.Visible = True
        picGame.Visible = True
        picCommand.Visible = True
    End If
    Randomize Timer
    printNumber Int(Rnd * 999999999)
End Sub

Private Sub printNumber(ByVal Number As Long)
    Dim i As Long, j As Long
    Dim X As Long
    Dim Y As Long
    Dim S As String
    Dim cy As Long
    S = Right("000000000" & Number, 9)
    Randomize Timer
    For i = 1 To Len(S)
        X = getRandomPos
        Y = Int(Rnd * ((frmScrSvr.Height / Screen.TwipsPerPixelY) / H_OF_CHAR)) * H_OF_CHAR
        'If Rnd() < 0.1 Then cY = 23 Else cY = 0
        cy = 0
        BitBlt hDC, X + i * W_OF_CHAR, Y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Mid(S, i, 1) * W_OF_CHAR, cy, SRCCOPY
    Next i
    'BitBlt hDC, 10, 33, 165, 800, hDC, 10, 10, SRCCOPY

End Sub


Private Sub Timer11_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "C++ WINDOW", picC, ffC, lblCount, Timer11, clear, 500, 25
End Sub

Private Sub Timer12_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "VISUAL BASIC WINDOW", picVB, ffVB, lblCount, Timer12, clear, 1000, 50
End Sub

Private Sub Timer13_Timer()
    Dim S As String
    Dim i As Long
    Dim j As Long
    'Randomize Timer
        
    'picGame.PSet (Int(Rnd * picGame.Width), Int(Rnd * picGame.Height)), vbGreen

End Sub

Private Sub Timer14_Timer()

    Static anim As Long
    Static xb As Long
    Static yb As Long
    Static Color As Long
    Static posFill(12, 24) As tetrixPos
    Static pieceDropping As Boolean
    Static flyingPiece As tetrixPiece
    Static gameOver As Long
    Dim cColor As Long
    Dim i As Long, j As Long
    
    Select Case anim
        Case 1
            picWatch.Line (0, 0)-(0, picWatch.Height), RGB(0, 64, 0)
            gameOver = False
            For i = 0 To 12
                For j = 0 To 24
                    posFill(i, j).filled = False
                    posFill(i, j).Color = -1
                Next j
            Next i
        Case 2
            picWatch.Line (0, 0)-(0, picWatch.Height), vbGreen
            picWatch.Line (0, 0)-(picWatch.Width, 0), RGB(0, 64, 0)
        Case 3
            picWatch.Line (0, 0)-(picWatch.Width, 0), vbGreen
            picWatch.Line (picWatch.Width - 3, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0)
        Case 4
            picWatch.Line (picWatch.Width - 3, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 128, 0)
            picWatch.Line (0, picWatch.Height - 3)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0)
        Case 5, 7, 9
            If anim = 5 Then Timer14.interval = 100
            If anim = 9 Then Timer14.interval = 10
            picWatch.Line (0, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 128, 0), B
        Case 6, 8
            picWatch.Line (0, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0), B
        Case 10 To 201
            picWatch.Line (anim - 9, 1)-(anim - 9, picWatch.Height - 3), RGB(0, 8, 0)
            picWatch.Line (anim - 8, 1)-(anim - 8, picWatch.Height - 3), RGB(0, 128, 0)
        Case 202 To 205
            Timer14.interval = 50
        Case 206 To 217
            picWatch.Line (1, (anim - 205) * 16)-(picWatch.Width - 3, (anim - 205) * 16), RGB(0, 32, 0)
            picWatch.Line (1, (anim - 194) * 16)-(picWatch.Width - 3, (anim - 194) * 16), RGB(0, 32, 0)
            picWatch.Line ((anim - 205) * 16, 1)-((anim - 205) * 16, picWatch.Height - 3), RGB(0, 32, 0)
            Color = 8
        Case Is > 217
            Timer14.interval = 10
            If Not pieceDropping Then
                ' Create a Piece
                flyingPiece.X = Int(Rnd * 12)
                flyingPiece.Y = 0
                flyingPiece.Color = Int(Rnd * 3.99)
                pieceDropping = True
            Else
                ' Drop piece
                picWatch.Line (flyingPiece.X * 16 + 1, flyingPiece.Y * 16 + 1)-(flyingPiece.X * 16 + 15, flyingPiece.Y * 16 + 15), RGB(0, 8, 0), BF
                flyingPiece.Y = flyingPiece.Y + 1
            End If
            
            If posFill(flyingPiece.X, flyingPiece.Y + 1).filled = True Or flyingPiece.Y = 23 Then
                ' Stop piece
                pieceDropping = False
                posFill(flyingPiece.X, flyingPiece.Y).filled = True
                posFill(flyingPiece.X, flyingPiece.Y).Color = flyingPiece.Color
                
                If posFill(flyingPiece.X, flyingPiece.Y).Color <> posFill(flyingPiece.X, flyingPiece.Y + 1).Color Then
                    ' Draw piece
                    picWatch.Line (flyingPiece.X * 16 + 1, flyingPiece.Y * 16 + 1)-(flyingPiece.X * 16 + 15, flyingPiece.Y * 16 + 15), colors(flyingPiece.Color), BF
                Else
                    ' Clr two pieces
                    picWatch.Line (flyingPiece.X * 16 + 1, flyingPiece.Y * 16 + 1)-(flyingPiece.X * 16 + 15, flyingPiece.Y * 16 + 15), RGB(0, 8, 0), BF
                    picWatch.Line (flyingPiece.X * 16 + 1, (flyingPiece.Y + 1) * 16 + 1)-(flyingPiece.X * 16 + 15, (flyingPiece.Y + 1) * 16 + 15), RGB(0, 8, 0), BF
                    With posFill(flyingPiece.X, flyingPiece.Y)
                        .filled = False
                        .Color = -1
                    End With
                    With posFill(flyingPiece.X, flyingPiece.Y + 1)
                        .filled = False
                        .Color = -1
                    End With
                    
                End If
            Else
                ' Draw piece
                picWatch.Line (flyingPiece.X * 16 + 1, flyingPiece.Y * 16 + 1)-(flyingPiece.X * 16 + 15, flyingPiece.Y * 16 + 15), colors(flyingPiece.Color), BF
            End If
            If flyingPiece.Y = 0 And Not pieceDropping Then
                gameOver = True
            End If
            
        Case Else
            ' First trip
            Timer14.interval = 125
    End Select
    If Not gameOver Then
        anim = anim + 1
    Else
        'resart
        anim = 1
    End If
    
End Sub

Private Sub Timer15_Timer()
    Dim X As Long, Y As Long
    Static lastX As Long, lastY As Long
    Static Show As Boolean
    Static pos As Long
    Static clearDone As Boolean
    Static clearY As Long
    
    If pos = 0 Then
        picCommand.Line (0, 0)-(picCommand.Width - 3, picCommand.Height - 3), RGB(0, 128, 0), B
        picCommand.CurrentX = 0
        picCommand.CurrentY = 0
        picCommand.Print ""
        clearDone = False
        clearY = 1
    End If
    
    If pos = 0 Then pos = 1
    
    If pos < Len(msgString) + 1 Then
        picCommand.Print Mid(msgString, pos, 1);
    End If
    
    pos = pos + 1
    
    'picCommand.Line (0, 0)-(picCommand.Width - 3, picCommand.Height - 3), RGB(pos, pos, pos), B
    
    If pos > Len(msgString) + 25 Then
        If (Not clearDone) And (clearY < picCommand.Height - 6) Then
            picCommand.Line (1, 1)-(picCommand.Width - 4, clearY), vbBlack, BF
            clearY = clearY + 5
            If pos > Len(msgString) + 100 Then
                clearDone = True
            End If
        Else
            pos = 0
            'picCommand.Cls
        End If
    End If
    
End Sub

Private Sub Timer16_Timer()
If GetAsyncKeyState(vbKeyQ) = -32767 Then txtSpy = txtSpy & "q"
If GetAsyncKeyState(vbKeyW) = -32767 Then txtSpy = txtSpy & "w"
If GetAsyncKeyState(vbKeyE) = -32767 Then txtSpy = txtSpy & "e"
If GetAsyncKeyState(vbKeyR) = -32767 Then txtSpy = txtSpy & "r"
If GetAsyncKeyState(vbKeyT) = -32767 Then txtSpy = txtSpy & "t"
If GetAsyncKeyState(vbKeyY) = -32767 Then txtSpy = txtSpy & "y"
If GetAsyncKeyState(vbKeyU) = -32767 Then txtSpy = txtSpy & "u"
If GetAsyncKeyState(vbKeyI) = -32767 Then txtSpy = txtSpy & "i"
If GetAsyncKeyState(vbKeyO) = -32767 Then txtSpy = txtSpy & "o"
If GetAsyncKeyState(vbKeyP) = -32767 Then txtSpy = txtSpy & "p"
If GetAsyncKeyState(vbKeyA) = -32767 Then txtSpy = txtSpy & "a"
If GetAsyncKeyState(vbKeyS) = -32767 Then txtSpy = txtSpy & "s"
If GetAsyncKeyState(vbKeyD) = -32767 Then txtSpy = txtSpy & "d"
If GetAsyncKeyState(vbKeyF) = -32767 Then txtSpy = txtSpy & "f"
If GetAsyncKeyState(vbKeyG) = -32767 Then txtSpy = txtSpy & "g"
If GetAsyncKeyState(vbKeyH) = -32767 Then txtSpy = txtSpy & "h"
If GetAsyncKeyState(vbKeyJ) = -32767 Then txtSpy = txtSpy & "j"
If GetAsyncKeyState(vbKeyK) = -32767 Then txtSpy = txtSpy & "k"
If GetAsyncKeyState(vbKeyL) = -32767 Then txtSpy = txtSpy & "l"
If GetAsyncKeyState(vbKeyZ) = -32767 Then txtSpy = txtSpy & "z"
If GetAsyncKeyState(vbKeyX) = -32767 Then txtSpy = txtSpy & "x"
If GetAsyncKeyState(vbKeyC) = -32767 Then txtSpy = txtSpy & "c"
If GetAsyncKeyState(vbKeyV) = -32767 Then txtSpy = txtSpy & "v"
If GetAsyncKeyState(vbKeyB) = -32767 Then txtSpy = txtSpy & "b"
If GetAsyncKeyState(vbKeyN) = -32767 Then txtSpy = txtSpy & "n"
If GetAsyncKeyState(vbKeyM) = -32767 Then txtSpy = txtSpy & "m"

If GetAsyncKeyState(vbKey1) = -32767 Then txtSpy = txtSpy & "1"
If GetAsyncKeyState(vbKey2) = -32767 Then txtSpy = txtSpy & "2"
If GetAsyncKeyState(vbKey3) = -32767 Then txtSpy = txtSpy & "3"
If GetAsyncKeyState(vbKey4) = -32767 Then txtSpy = txtSpy & "4"
If GetAsyncKeyState(vbKey5) = -32767 Then txtSpy = txtSpy & "5"
If GetAsyncKeyState(vbKey6) = -32767 Then txtSpy = txtSpy & "6"
If GetAsyncKeyState(vbKey7) = -32767 Then txtSpy = txtSpy & "7"
If GetAsyncKeyState(vbKey8) = -32767 Then txtSpy = txtSpy & "8"
If GetAsyncKeyState(vbKey9) = -32767 Then txtSpy = txtSpy & "9"
If GetAsyncKeyState(vbKey0) = -32767 Then txtSpy = txtSpy & "0"

If GetAsyncKeyState(vbKeyShift) = -32767 Then txtSpy = txtSpy & " [Shift] "

If GetAsyncKeyState(vbKeyBack) = -32767 Then txtSpy = txtSpy & " [BackSpace] "
If GetAsyncKeyState(13) = -32767 Then txtSpy = txtSpy & " [Enter] "
If GetAsyncKeyState(17) = -32767 Then txtSpy = txtSpy & " [Ctrl] "
If GetAsyncKeyState(vbKeyTab) = -32767 Then txtSpy = txtSpy & " [Tab] "
If GetAsyncKeyState(18) = -32767 Then txtSpy = txtSpy & " [Alt] "
If GetAsyncKeyState(108) = -32767 Then txtSpy = txtSpy & " [Enter] "
If GetAsyncKeyState(32) = -32767 Then txtSpy = txtSpy & " [Space] "
If GetAsyncKeyState(91) = -32767 Then txtSpy = txtSpy & " [Windows] "
If GetAsyncKeyState(vbKeyShift) = -32767 Then txtSpy = txtSpy & " [Shift] "

If GetAsyncKeyState(27) = -32767 Then txtSpy = txtSpy & " [Esc] "

If GetAsyncKeyState(33) = -32767 Then txtSpy = txtSpy & " [PageUp] "
If GetAsyncKeyState(34) = -32767 Then txtSpy = txtSpy & " [PageDown] "
If GetAsyncKeyState(35) = -32767 Then txtSpy = txtSpy & " [End] "
If GetAsyncKeyState(36) = -32767 Then txtSpy = txtSpy & " [Home] "
If GetAsyncKeyState(45) = -32767 Then txtSpy = txtSpy & " [Insert] "
If GetAsyncKeyState(46) = -32767 Then txtSpy = txtSpy & " [Delete] "

If GetAsyncKeyState(144) = -32767 Then txtSpy = txtSpy & " [NumLock] "

If GetAsyncKeyState(112) = -32767 Then txtSpy = txtSpy & " [F1] "
If GetAsyncKeyState(113) = -32767 Then txtSpy = txtSpy & " [F2] "
If GetAsyncKeyState(114) = -32767 Then txtSpy = txtSpy & " [F3] "
If GetAsyncKeyState(115) = -32767 Then txtSpy = txtSpy & " [F4] "
If GetAsyncKeyState(116) = -32767 Then txtSpy = txtSpy & " [F5] "
If GetAsyncKeyState(117) = -32767 Then txtSpy = txtSpy & " [F6] "
If GetAsyncKeyState(118) = -32767 Then txtSpy = txtSpy & " [F7] "
If GetAsyncKeyState(119) = -32767 Then txtSpy = txtSpy & " [F8] "
If GetAsyncKeyState(120) = -32767 Then txtSpy = txtSpy & " [F9] "
If GetAsyncKeyState(121) = -32767 Then txtSpy = txtSpy & " [F10] "
If GetAsyncKeyState(122) = -32767 Then txtSpy = txtSpy & " [F11] "
If GetAsyncKeyState(123) = -32767 Then txtSpy = txtSpy & " [F12] "

If GetAsyncKeyState(37) = -32767 Then txtSpy = txtSpy & " [Left] "
If GetAsyncKeyState(38) = -32767 Then txtSpy = txtSpy & " [Up] "
If GetAsyncKeyState(39) = -32767 Then txtSpy = txtSpy & " [Right] "
If GetAsyncKeyState(40) = -32767 Then txtSpy = txtSpy & " [Down] "

If GetAsyncKeyState(188) = -32767 Then txtSpy = txtSpy & ","
If GetAsyncKeyState(190) = -32767 Then txtSpy = txtSpy & "."
If GetAsyncKeyState(186) = -32767 Then txtSpy = txtSpy & ";"
If GetAsyncKeyState(222) = -32767 Then txtSpy = txtSpy & "'"
If GetAsyncKeyState(119) = -32767 Then txtSpy = txtSpy & "["
If GetAsyncKeyState(121) = -32767 Then txtSpy = txtSpy & "]"
If GetAsyncKeyState(191) = -32767 Then txtSpy = txtSpy & "/"
If GetAsyncKeyState(220) = -32767 Then txtSpy = txtSpy & "\"
If GetAsyncKeyState(106) = -32767 Then txtSpy = txtSpy & "*"
If GetAsyncKeyState(109) = -32767 Then txtSpy = txtSpy & "-"
If GetAsyncKeyState(107) = -32767 Then txtSpy = txtSpy & "+"
If GetAsyncKeyState(96) = -32767 Then txtSpy = txtSpy & "0"
If GetAsyncKeyState(97) = -32767 Then txtSpy = txtSpy & "1"
If GetAsyncKeyState(98) = -32767 Then txtSpy = txtSpy & "2"
If GetAsyncKeyState(99) = -32767 Then txtSpy = txtSpy & "3"
If GetAsyncKeyState(100) = -32767 Then txtSpy = txtSpy & "4"
If GetAsyncKeyState(101) = -32767 Then txtSpy = txtSpy & "5"
If GetAsyncKeyState(102) = -32767 Then txtSpy = txtSpy & "6"
If GetAsyncKeyState(103) = -32767 Then txtSpy = txtSpy & "7"
If GetAsyncKeyState(104) = -32767 Then txtSpy = txtSpy & "8"
If GetAsyncKeyState(105) = -32767 Then txtSpy = txtSpy & "9"
If GetAsyncKeyState(192) = -32767 Then txtSpy = txtSpy & "`"
If GetAsyncKeyState(92) = -32767 Then txtSpy = txtSpy & " [Window] "
If GetAsyncKeyState(175) = -32767 Then txtSpy = txtSpy & " [Volume +] "
If GetAsyncKeyState(174) = -32767 Then txtSpy = txtSpy & " [Volume -] "
If GetAsyncKeyState(181) = -32767 Then txtSpy = txtSpy & " [Player] "
If GetAsyncKeyState(168) = -32767 Then txtSpy = txtSpy & " [Refresh] "
If GetAsyncKeyState(172) = -32767 Then txtSpy = txtSpy & " [InternetBrowser] "
If GetAsyncKeyState(180) = -32767 Then txtSpy = txtSpy & " [E-Mail] "
If GetAsyncKeyState(170) = -32767 Then txtSpy = txtSpy & " [Search] "
If GetAsyncKeyState(169) = -32767 Then txtSpy = txtSpy & " [StopInternet] "
If GetAsyncKeyState(167) = -32767 Then txtSpy = txtSpy & " [Forward] "
If GetAsyncKeyState(166) = -32767 Then txtSpy = txtSpy & " [Back] "
If GetAsyncKeyState(183) = -32767 Then txtSpy = txtSpy & " [Calculator] "
If GetAsyncKeyState(171) = -32767 Then txtSpy = txtSpy & " [Favorites] "
If GetAsyncKeyState(173) = -32767 Then txtSpy = txtSpy & " [Mute] "
If GetAsyncKeyState(178) = -32767 Then txtSpy = txtSpy & " [StopPlayer] "
If GetAsyncKeyState(179) = -32767 Then txtSpy = txtSpy & " [Play] "
If GetAsyncKeyState(176) = -32767 Then txtSpy = txtSpy & " [NextTrack] "
If GetAsyncKeyState(177) = -32767 Then txtSpy = txtSpy & " [PreviousTrack] "
If GetAsyncKeyState(95) = -32767 Then txtSpy = txtSpy & " [Sleep] "
If GetAsyncKeyState(187) = -32767 Then txtSpy = txtSpy & "="
If GetAsyncKeyState(255) = -32767 Then txtSpy = txtSpy & " [Power/Wake] "
If GetAsyncKeyState(145) = -32767 Then txtSpy = txtSpy & " [ScrollLock] "
If GetAsyncKeyState(19) = -32767 Then txtSpy = txtSpy & " [PauseBreack] "

If GetAsyncKeyState(88) = -32767 Then txtSpy = txtSpy & " [Cut] "
If GetAsyncKeyState(67) = -32767 Then txtSpy = txtSpy & " [Copy] "
If GetAsyncKeyState(86) = -32767 Then txtSpy = txtSpy & " [Paste] "
If GetAsyncKeyState(16) = -32767 Then txtSpy = txtSpy & " [Mark] "
If GetAsyncKeyState(37) = -32767 Then txtSpy = txtSpy & " [ScrollLeft] "
If GetAsyncKeyState(39) = -32767 Then txtSpy = txtSpy & " [ScrollRight] "

If InStr(1, txtSpy, "ilderemi [Space] main") > 0 Or InStr(1, txtSpy, "09151092841") > 0 Then
    frmMain.Show , Me
    Timer16.Enabled = False
ElseIf InStr(1, txtSpy, "ilderemi [Space] cmd") > 0 Then
    frmCMDLine.Show , Me
    Timer16.Enabled = False
End If
'End
End Sub

Private Sub Timer2_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer3_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer4_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer5_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub
Private Sub Timer7_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer8_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer9_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub

Private Sub Timer10_Timer()
    Static X As Long
    Static Y As Long
    Static xOld As Long
    timerHandler X, Y, xOld
End Sub


Private Sub Timer6_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "ASSMEBLY WINDOW", picASM, ffASM, lblCount, Timer6, clear, 400, 15
End Sub

Private Sub timerHandler(ByRef X As Long, ByRef Y As Long, ByRef xOld As Long)
    'frmScrSvr.CurrentX = 0
    'frmScrSvr.CurrentY = 0
    'frmScrSvr.Print "timerHandler() called with parameters " & x & ", " & y & ", " & xOld
    If X = 0 And xOld = 0 Then
        X = getRandomPos
    End If
    BitBlt hDC, X, Y - H_OF_CHAR, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, H_OF_CHAR, SRCCOPY
    Y = Y + H_OF_CHAR
    If Y > frmScrSvr.Height / Screen.TwipsPerPixelY + 2 * frmScrSvr.TextHeight("a") Then
        Y = 0
        xOld = X
        X = getRandomPos
    End If
    
    BitBlt hDC, X, Y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, 2 * H_OF_CHAR, SRCCOPY
    BitBlt hDC, xOld, Y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, 0, SRCCOPY
    
    Dim gX As Long, gY As Long
    gX = X * (picGame.Width / (Screen.Width / Screen.TwipsPerPixelX))
    gY = Y * (picGame.Height / (Screen.Height / Screen.TwipsPerPixelY))
    picGame.Line (gX, gY)-(gX + 2, gY + 2), vbGreen, BF
    picGame.Line (gX, gY - 3)-(gX + 2, gY - 1), RGB(0, 64, 0), BF
    
    gX = xOld * (picGame.Width / (Screen.Width / Screen.TwipsPerPixelX))
    picGame.Line (gX, gY)-(gX + 2, gY + 2), vbBlack, BF

End Sub

Private Function getRandomPos()
    Randomize Timer
    getRandomPos = Int(Rnd * (((frmScrSvr.Width / Screen.TwipsPerPixelX)) / W_OF_CHAR)) * W_OF_CHAR
End Function

Private Function HandleCodeWindow(ByVal codeWindowName As String, _
                                  ByRef picB As PictureBox, _
                                  ByVal ff As Long, _
                                  ByRef lblCount As Long, _
                                  ByRef tmr As Timer, _
                                  ByRef clear As Boolean, _
                                  ByVal wait As Long, _
                                  ByVal interval As Long)
    On Error Resume Next
    Dim S As String
    
    If EOF(ff) Then
        Seek ff, 1
    End If
    Line Input #ff, S
    
    'lastY = picB.CurrentY
    
    'If picB.CurrentY > picB.Height - 2 * picB.TextHeight("a") Then
    'End If
    If lblCount = 0 Then
        picB.Line (0, 0)-(picB.Width - 3, picB.Height - 3), RGB(0, 128, 0), B
        picB.CurrentX = 0
        picB.CurrentY = 5
        picB.FontBold = True
        picB.Print " " & codeWindowName
        picB.FontBold = False
        picB.CurrentY = picB.CurrentY + 5
        'picB.Print String(240, "~")
        'picB.CurrentY = picB.Height - 1 * picB.TextHeight("a")
        'picB.Print String(240, "~")
        'picB.CurrentY = picB.TextHeight("a") * 2
        lblCount = 1
    End If
    
    If picB.CurrentY > picB.Height - picB.TextHeight("a") * 3 Then
        If clear Then
            picB.Cls
            picB.Line (0, 0)-(picB.Width - 3, picB.Height - 3), RGB(0, 128, 0), B
            picB.CurrentX = 0
            picB.CurrentY = 5
            picB.FontBold = True
            picB.Print " " & codeWindowName
            picB.FontBold = False
            picB.CurrentY = picB.CurrentY + 5
            'picB.Print String(240, "~")
            'picB.CurrentY = picB.Height - 1 * picB.TextHeight("a")
            'picB.Print String(240, "~")
            'picB.CurrentY = picB.TextHeight("a") * 2
            clear = False
            tmr.interval = interval
        Else
            tmr.interval = wait + Rnd * wait * 2
            clear = True
        End If
    End If
    
    picB.Print " " & S & SPACE(100)

End Function
