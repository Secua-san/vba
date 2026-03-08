Attribute VB_Name = "SemanticTokens"
Option Explicit

Private Type CustomerRecord
    Name As String
End Type

Private Const DefaultName As String = "A"

Public Function BuildCustomer(ByVal sourceName As String) As CustomerRecord
    Dim customer As CustomerRecord
    customer.Name = sourceName
    BuildCustomer = customer
End Function

Public Sub Demo()
    Dim current As CustomerRecord
    current = BuildCustomer(DefaultName)
End Sub
