Attribute VB_Name = "LoadEXEFile"
Option Explicit
Public MyFile() As Byte
Public Sub LoadMyFile(num As Integer)
  MyFile = LoadResData(num, "exe")
End Sub
