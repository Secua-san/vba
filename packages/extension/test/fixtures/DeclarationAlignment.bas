Attribute VB_Name = "DeclarationAlignment"
Option Explicit

Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Public Sub Demo()
Dim title As String
Dim count       As Long
Dim enabled As Boolean

Const DefaultTitle As String = "Ready"
Const RetryCount  As Long=3
Const IsEnabled As Boolean   = True

Debug.Print title, count, enabled
End Sub
