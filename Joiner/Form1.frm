VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Dim s As String

'Kill "c:\0.exe"
'LoadMyFile (101)
'Open "c:\0.txt" For Binary As #1
'    Put #1, 1, MyFile
'Close
LoadMyFile (101)
For i = 0 To LenB(LoadResData(101, "txt")) - 1
    s = s & Chr(MyFile(i))
Next
Text1 = s
End Sub
