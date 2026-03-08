Attribute VB_Name = "FormatterApi"
Option Explicit

Public Function FormatMessage(ByVal value As String, ByVal count As Long) As String
    FormatMessage = value & CStr(count)
End Function
