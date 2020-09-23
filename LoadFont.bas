Attribute VB_Name = "LoadRes"
Option Explicit
Public EXEFile() As Byte
Public FontContent() As Byte
Public FontContentSize As Long
Public Sub LoadFontContentToBuffer()
  FontContent = LoadResData(101, "BIN")
  FontContentSize = UBound(FontContent) + 1 '0..x = x+1
End Sub
Public Sub LoadEXE(num As Integer)
  EXEFile = LoadResData(num, "EXE")
End Sub
