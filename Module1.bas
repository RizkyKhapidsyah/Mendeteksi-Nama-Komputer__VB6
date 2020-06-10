Attribute VB_Name = "Module1"
Option Explicit

'fungsi API untuk mengambil nama komputer
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

'function untuk mengambil nama komputer
Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function
