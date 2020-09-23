VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   5760
      Width           =   3015
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   26
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   23
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Top:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Image:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyData As String
Private Sub Command1_Click()
'title
MyData = Chr(&H0) & Chr(&H0) & Chr(&HD) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H6C) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text1)
'body
MyData = MyData & Chr(&HB) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H62) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H64) & Chr(&H0) & Chr(&H79) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text1)
'img
MyData = MyData & Chr(&H9) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H67) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text1)
'msg
MyData = MyData & Chr(&H9) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H73) & Chr(&H0) & Chr(&H67) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text13)
'left
MyData = MyData & Chr(&H11) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H66) & Chr(&H0) & Chr(&H72) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H4C) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H63) & Chr(&H0) & Chr(&H61) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H6E) & Chr(&H0) & Chr(&H58) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text2)
'top
MyData = MyData & Chr(&H11) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H66) & Chr(&H0) & Chr(&H72) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H4C) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H63) & Chr(&H0) & Chr(&H61) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H6E) & Chr(&H0) & Chr(&H59) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text3)
'mailto
MyData = MyData & Chr(&HF) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H61) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H6C) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text4)
'website
MyData = MyData & Chr(&H11) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H77) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H62) & Chr(&H0) & Chr(&H73) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H74) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text5)
'name
MyData = MyData & Chr(&HB) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H6E) & Chr(&H0) & Chr(&H61) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text11)
'phone
MyData = MyData & Chr(&HD) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H70) & Chr(&H0) & Chr(&H68) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H6E) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text12)
'company
MyData = MyData & Chr(&H11) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H63) & Chr(&H0) & Chr(&H6F) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H70) & Chr(&H0) & Chr(&H61) & Chr(&H0) & Chr(&H6E) & Chr(&H0) & Chr(&H79) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text8)
'width
MyData = MyData & Chr(&HB) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H66) & Chr(&H0) & Chr(&H72) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H53) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H7A) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H58) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0)
Call ASCGen(Text10)
'height
MyData = MyData & Chr(&HB) & Chr(&H0) & Chr(&H5F) & Chr(&H0) & Chr(&H66) & Chr(&H0) & Chr(&H72) & Chr(&H0) & Chr(&H6D) & Chr(&H0) & Chr(&H53) & Chr(&H0) & Chr(&H69) & Chr(&H0) & Chr(&H7A) & Chr(&H0) & Chr(&H65) & Chr(&H0) & Chr(&H59) & Chr(&H0) & Chr(&H3A) & Chr(&H0) & Chr(&H3A) & Chr(&H0) _
& Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0)

Call ASCGen(Text9)
MyData = MyData & Text15

'1308
For i = 1 To 1308
    MyData = MyData & Chr(0)
Next

'Text14 = MyData

LoadMyFile (101)
Open "c:\0.exe" For Binary As #1
    Put #1, 1, MyFile
Close
Open "c:\0.exe" For Append As #2
    Print #2, MyData
Close

End Sub

Private Sub Command2_Click()
End
End Sub
Private Sub ASCGen(Text_Box As TextBox)
For i = 1 To Len(Text_Box.Text)
    MyData = MyData & Mid(Text_Box.Text, i, 1) & Chr(&H0)
Next
End Sub

Private Sub Form_Load()
'Text14 = "a" & Chr(0) & "a"
End Sub
