VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   4380
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4380
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4380
      Width           =   1875
   End
   Begin VB.PictureBox Picture4 
      Height          =   255
      Left            =   3180
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3180
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   2475
      Left            =   4500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "Form1.frx":014A
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   4275
      Left            =   3360
      ScaleHeight     =   4215
      ScaleWidth      =   5895
      TabIndex        =   1
      Top             =   60
      Width           =   5955
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   255
         TabIndex        =   5
         Top             =   2520
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1020
         TabIndex        =   4
         Top             =   3480
         Width           =   4455
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2835
         Left            =   5460
         TabIndex        =   3
         Top             =   1020
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   2595
         Left            =   60
         ScaleHeight     =   2535
         ScaleWidth      =   4335
         TabIndex        =   2
         Top             =   1260
         Width           =   4395
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5953
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3420
      Width           =   3195
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0150
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0262
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0374
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0486
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentResType As String, CurrentResName As String

Private Sub Command1_Click() 'Open
   CurrentResType = "": CurrentResName = ""
   RefreshView
   ClearResource
   cdlg.Filter = "Executable (dll,exe)|*.dll;*.exe|All files (*.*)|*.*"
   cdlg.InitDir = App.Path
   cdlg.ShowOpen
   If cdlg.FileName <> "" Then
      Call FillResTypes(TreeView1, cdlg.FileName, cdlg.FileTitle)
   End If
End Sub

Private Sub Command2_Click() 'Save
   Dim srcPic As StdPicture
   Dim srcText As String, sTemp As String
   Dim srcArr() As Byte
   Dim nd As Node
   cdlg.FileName = ""
   cdlg.FilterIndex = 1
   cdlg.InitDir = App.Path
   If CurrentResType <> "6" And CurrentResName = "" Then
      MsgBox "No resource selected!", vbCritical, "Error"
      Exit Sub
   End If
   Select Case UCase(CurrentResType)
      Case "1", "12" 'Cursors
           cdlg.Filter = "Cursor files (*.cur)|*.cur|Bitmap files (*.bmp)|*.bmp"
           Set srcPic = GetPicture(CurrentResType, CurrentResName)
      Case "2" 'Bitmap
           cdlg.Filter = "Bitmap files (*.bmp)|*.bmp"
           Set srcPic = GetPicture(CurrentResType, CurrentResName)
      Case "3", "14" 'Icon
           cdlg.Filter = "Icon files (*.ico)|*.ico|Bitmap files (*.bmp)|*.bmp"
           Set srcPic = GetPicture(CurrentResType, CurrentResName)
      Case "4" 'Menu
           cdlg.Filter = "Save as text (*.txt)|*.txt|Save as data (*.*)|*.*"
           srcText = Text1.Text
      Case "6" 'String
           cdlg.Filter = "Save as text (*.txt)|*.txt"
           If CurrentResName <> "" Then
              srcText = Text1.Text
           Else
              TreeView1_Expand TreeView1.SelectedItem
              Set nd = TreeView1.SelectedItem.Child
              Do
                If nd Is Nothing Then Exit Do
                sTemp = nd.Text
                If IsNumeric(sTemp) Then sTemp = "#" & sTemp
                srcText = srcText & GetString(sTemp) & vbCrLf
                Set nd = nd.Next
              Loop
           End If
      Case "9" 'Accelerators Table
           cdlg.Filter = "Save as text (*.txt)|*.txt|Save as data (*.*)|*.*"
           srcText = Text1.Text
      Case "11" 'Message Table
           cdlg.Filter = "Save as text (*.txt)|*.txt"
           srcText = Text1.Text
      Case "16" 'version info
           cdlg.Filter = "Save as text (*.txt)|*.txt|Save as data (*.*)|*.*"
           srcText = Text1.Text
      Case "23", "HTML"
           cdlg.Filter = "HTML files (*.html)|*.html"
      Case "AVI", "JPG", "JPEG", "GIF", "PNG", "TIF", "TIFF", "WMF", "EMF"
           cdlg.Filter = UCase(CurrentResType) & " files (*." & LCase(CurrentResType) & ")|*." & LCase(CurrentResType)
      Case Else
           cdlg.Filter = "Save as data (*.*)|*.*"
  End Select
  cdlg.ShowSave
  If cdlg.FileName = "" Then Exit Sub
  If Not srcPic Is Nothing Then
     If cdlg.FilterIndex = 1 Then
        SavePicture srcPic, cdlg.FileName
     Else
        Picture4.Picture = srcPic
        SavePicture Picture4.Image, cdlg.FileName
     End If
  ElseIf (srcText <> "") And (cdlg.FilterIndex = 1) Then
     SaveText cdlg.FileName, srcText
  Else
     srcArr = GetDataArray(CurrentResType, CurrentResName)
     SaveData cdlg.FileName, srcArr
  End If
ErrSave:
  If Err Then MsgBox "Unable to save resource", vbCritical, "Error"
  Set srcPic = Nothing
  Set nd = Nothing
End Sub

Private Sub Command3_Click() 'Exit
   Unload Me
End Sub

Private Sub Form_Load()
  Label1 = ""
  Caption = "Ark's resource Viewer/Extractor"
  Command1.Caption = "&Open file with resources"
  Command2.Caption = "&Save resource"
  Command3.Caption = "&Exit"
  With VScroll1
       .Move Picture2.Width - .Width - 60, 0, .Width, Picture2.Height - HScroll1.Height - 60
       .SmallChange = 1
       .LargeChange = 10
       .Enabled = False
  End With
  With HScroll1
       .Move 0, Picture2.Height - .Height - 60, Picture2.Width - VScroll1.Width - 60, .Height
       .SmallChange = 1
       .LargeChange = 10
       .Enabled = False
        Picture3.Move VScroll1.Left, .Top, VScroll1.Width, .Height
   End With
   With Picture1
      .BorderStyle = 0
      .BackColor = vbButtonFace
      .AutoRedraw = True
      .Move 0, 0, Picture2.Width - VScroll1.Width - 60, Picture2.Height - HScroll1.Height - 60
      picWidth = .Width
      picHeight = .Height
   End With
   With Text1
      .Move Picture2.Left, Picture2.Top, Picture2.Width, Picture2.Height
      .Visible = False
      .BackColor = vbButtonFace
      .FontName = "courier new"
   End With
   With Picture4
       .BorderStyle = 0
       .AutoRedraw = True
       .AutoSize = True
       .Visible = False
   End With
   Picture1_Resize
   Call FillResTypes(TreeView1, "Shell32.dll", "shell32.dll")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearResource
End Sub

Private Sub HScroll1_Change()
   Picture1.Left = -HScroll1.Value * Screen.TwipsPerPixelX
End Sub

Private Sub Picture1_Resize()
  HScroll1.Enabled = (Picture1.Width > picWidth)
  VScroll1.Enabled = (Picture1.Height > picHeight)
  If HScroll1.Enabled Then
     HScroll1.Value = 0
     HScroll1.Max = ((Picture1.Width - Picture2.Width) + 3 * Picture1.TextWidth("A")) \ Screen.TwipsPerPixelY
  End If
  If VScroll1.Enabled Then
     VScroll1.Value = 0
     VScroll1.Max = (Picture1.Height - Picture2.Height) \ Screen.TwipsPerPixelY
  End If
End Sub

Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
  RefreshView
End Sub

Private Sub TreeView1_Expand(ByVal Node As Node)
   If Node.Child.Text = "Dummy" Then
      TreeView1.Nodes.Remove Node.Child.Index
      Call FillResNames(TreeView1, Node)
   End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)
   Dim ResType As String, ResName As String
   Dim ret As Boolean
   Text1.Visible = False
   RefreshView
   CurrentResType = "": CurrentResName = ""
   If Node = Node.Root Then Exit Sub
   Label1 = "ResType: " & Node.Text
   If Node.key = "" Then
      CurrentResType = Node.Text
   Else
      CurrentResType = Mid(Node.key, 2)
   End If
   If Node.Parent = Node.Root Then Exit Sub
   MousePointer = vbHourglass
   If Node.Parent.key = "" Then
      ResType = Node.Parent.Text
   Else
      ResType = Mid(Node.Parent.key, 2)
   End If
   ResName = Node.Text
   If IsNumeric(ResName) Then ResName = "#" & ResName
   CurrentResType = ResType: CurrentResName = ResName
   Label1 = "ResType: " & Node.Parent.Text & vbCrLf & "ResName: " & ResName & vbCrLf & "ResSize: " & ResSize(ResType, ResName) & " bytes"
   Select Case UCase(ResType)
      Case "1", "2", "3", "12", "14"
           ret = ShowPicture(GetPicture(ResType, ResName), Picture1)
      Case "4" 'Menu
           Text1.Visible = True
           ret = ShowText(GetMenuText(ResName), Text1)
      Case "5", "17" 'Dialog
           ret = ShowDialog(ResName, Picture1)
      Case "6" 'String
           Text1.Visible = True
           ret = ShowText(GetString(ResName), Text1)
      Case "9" 'Accelerators Table
           Text1.Visible = True
           ret = ShowText(GetAccelerators(ResName), Text1)
      Case "11" 'Message Table
           Text1.Visible = True
           ret = ShowText(GetMessageTable(ResName), Text1)
      Case "16" 'version info
           Text1.Visible = True
           ret = ShowText(GetVersionInfo(ResName), Text1)
      Case "23", "HTML"
           Text1.Visible = True
           ret = ShowText(GetHTML(ResType, ResName), Text1)
      Case "AVI"
           ret = ShowAVI(ResName, Picture1)
      Case "JPG", "JPEG", "GIF", "PNG", "TIF", "TIFF", "WMF", "EMF"
           ret = ShowPicture(GetPictureExt(ResType, ResName), Picture1)
      Case Else
           Text1.Visible = True
           ret = ShowText(GetHexDump(ResType, ResName), Text1)
  End Select
  If ret = False Then
     If Text1.Visible Then
        Text1.Text = Text1.Text & vbNewLine & "Can not load resourse"
     Else
        Picture1.Print "Can not load resourse"
     End If
  End If
  Picture1.Refresh
  MousePointer = vbDefault
 End Sub

Private Sub VScroll1_Change()
   Picture1.Top = -VScroll1.Value * Screen.TwipsPerPixelY
End Sub

Private Sub VScroll1_GotFocus()
   Picture1.SetFocus
End Sub

Private Sub HScroll1_GotFocus()
   Picture1.SetFocus
End Sub

Private Sub RefreshView()
   If Dir(TEMP_FILE_NAME) <> "" Then
      Call mciSendString("close video", 0&, 0, 0)
      Kill TEMP_FILE_NAME
   End If
   If hDialog Then Call DestroyWindow(hDialog)
   Picture1.Cls
   Picture1.Refresh
   Label1 = ""
   Text1 = ""
End Sub
