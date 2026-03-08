Attribute VB_Name = "CommentFormatting"
Option Explicit

Public Sub Demo()
'leading comment
Dim value As Long'counter
If True Then'true branch
'inner comment
value = 1'updated
Rem    status
#If VBA7 Then'requires vba7
'conditional comment
#Else'fallback path
Rem    fallback comment
#End If
End If
End Sub
