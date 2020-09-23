Attribute VB_Name = "LoadFile"
Option Explicit
Public MyFile() As Byte
Public MyTXTFile As String
Public Sub LoadMyFile(num As Integer)
  MyFile = LoadResData(num, "txt")
End Sub
Public Function LoadCommand(num As Integer) As String
    LoadCommand = LoadResString(num)
End Function
