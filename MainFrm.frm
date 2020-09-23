VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   7560
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   540
      Top             =   1425
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   0
      Top             =   -15
      Width           =   7560
   End
End
Attribute VB_Name = "Form1"
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
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Const BrightYellow = &HFFFF&
Const DarkYellow = &H6060&
Dim mDC As Long
Dim mBitmap As Long
Dim nDC As Long
Dim nBitmap As Long
Dim TotalStringLength As Long
Const CharWidth = 6
Const CharHeight = 8
Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Timer1.Enabled = True
  'Load font to buffer
  LoadFontContentToBuffer
  'Set the graphic mode to persistent
  Me.AutoRedraw = True
  'API uses pixels
  Me.ScaleMode = vbPixels
  'Create a device context, compatible with the screen
  mDC = CreateCompatibleDC(GetDC(0))
  'Create a bitmap, compatible with the screen
  mBitmap = CreateCompatibleBitmap(GetDC(0), FontContentSize * 3 - 1, CharHeight * 3 - 1)
  'Select the bitmap into the device context
  SelectObject mDC, mBitmap
  'Create a device context, compatible with the screen
  nDC = CreateCompatibleDC(GetDC(0))
  'Create a bitmap, compatible with the screen
  nBitmap = CreateCompatibleBitmap(GetDC(0), FontContentSize * 3 - 1, CharHeight * 3 - 1)
  'Select the bitmap into the device context
  SelectObject nDC, nBitmap
  Dim i As Long, j As Long, k As Long, l As Long
  For i = 0 To (FontContentSize - 1)
    For j = 0 To (CharHeight - 1)
      For k = 0 To 1
        For l = 0 To 1
          If FontContent(i) \ (2 ^ (7 - j)) Mod 2 = 1 Then
            SetPixel mDC, (i * 3) + k, (j * 3) + l, BrightYellow
          Else
            SetPixel mDC, (i * 3) + k, (j * 3) + l, DarkYellow
          End If
          SetPixel mDC, (i * 3) + k, (j * 3) + 2, &H0
          SetPixel mDC, (i * 3) + 2, (j * 3) + l, &H0
          SetPixel mDC, (i * 3) + 2, (j * 3) + 2, &H0
        Next l
      Next k
    Next j
  Next i
End Sub
Public Sub LedTableShow(str As String)
  Dim i As Long
  TotalStringLength = Len(str) * CharWidth * 3
  For i = 0 To Len(str) - 1
    Select Case Asc(Mid(str, i + 1, 1))
      Case 32 To 126
        BitBlt nDC, i * (CharWidth * 3), 0, (CharWidth * 3), CharHeight * 3 - 1, mDC, (Asc(Mid(str, i + 1, 1)) - 32) * (CharWidth * 3), 0, vbSrcCopy
      Case Else
        BitBlt nDC, i * (CharWidth * 3), 0, (CharWidth * 3), CharHeight * 3 - 1, mDC, 0, 0, vbSrcCopy
    End Select
  Next i
End Sub
Private Sub Form_Resize()
  If Me.Height <> 690 / 2 Then Me.Height = 690 / 2
End Sub
Private Sub Timer1_Timer()
  Static i As Long
  i = i - 3
  If TotalStringLength > 0 Then
    If i < TotalStringLength Then i = i + TotalStringLength
    i = i Mod TotalStringLength
  End If
  LedTableShow "                        Designed by Masoud iLDEREMi"
  Cls
  BitBlt Me.hDC, i, 0, TotalStringLength - i, CharHeight * 3 - 1, nDC, 0, 0, vbSrcCopy
  BitBlt Me.hDC, 0, 0, i, CharHeight * 3 - 1, nDC, TotalStringLength - i, 0, vbSrcCopy
  Me.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  DeleteDC mDC
  DeleteObject mBitmap
  DeleteDC nDC
  DeleteObject nBitmap
  Timer1.Enabled = False
End Sub

