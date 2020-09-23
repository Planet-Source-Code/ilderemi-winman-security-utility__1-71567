VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Effects 2004"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Tag = Me.cmb.ListIndex
frmMain.doeffect
End Sub
Private Sub Form_Load()
cmb.AddItem "Melt" '0
cmb.AddItem "Powder Blow" '1
cmb.AddItem "Powder" '2
cmb.AddItem "Evaporate" '3
cmb.AddItem "Water Color" '4
cmb.AddItem "Accumulate" '5
cmb.AddItem "Checks" '6
cmb.AddItem "Extreme Checks" '7
cmb.AddItem "Wind Blow" '8
cmb.AddItem "Pour Down" '9
On Error GoTo er2
cmb.ListIndex = GetSetting("MeltSCR", "Effect", "Effect")
Exit Sub
er2:
cmb.ListIndex = 0
End Sub


