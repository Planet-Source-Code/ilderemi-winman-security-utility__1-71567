Attribute VB_Name = "LoadFile"
Option Explicit
Public MyFile() As Byte
Public Sub LoadMyFile(num As Integer)
  MyFile = LoadResData(num, "txt")
End Sub
