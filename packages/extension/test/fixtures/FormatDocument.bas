Attribute VB_Name = "FormatDocument"
Option Explicit

Public Sub Demo()
If True Then
Debug.Print "ready"
Else
Select Case 1
Case Else
Debug.Print "fallback"
End Select
End If
End Sub
