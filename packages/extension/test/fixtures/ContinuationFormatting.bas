Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
Dim message As String
message =   _
"prefix" &  _
"suffix"

Debug.Print JoinValues( _
message, _
"tail" _
)

message = CreateBuilder() _
.WithName(message) _
.WithSuffix("!")
End Sub
