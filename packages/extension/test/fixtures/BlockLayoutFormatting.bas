Attribute VB_Name = "BlockLayoutFormatting"
Option Explicit

Public Sub Demo()
Dim value As Long: value = 0
If value = 0 Then: Debug.Print "zero": ElseIf value = 1 Then: Debug.Print "one": Else: Debug.Print "other": End If
Select Case value: Case 0: Debug.Print "case zero": Case Else: With Application: .StatusBar = "fallback": End With: End Select
#If VBA7 Then: value = value + 1: #Else: value = value - 1: #End If
End Sub
