Attribute VB_Name = "BuiltInMemberCompletion"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Application.
    Debug.Print WorksheetFunction.Su
    Debug.Print Application.WorksheetFunction.Su
    Debug.Print ActiveWorkbook.
    Debug.Print ThisWorkbook.
    Debug.Print ActiveWorkbook.Worksheets.
    Debug.Print Worksheets(1).
    Debug.Print Worksheets("A(1)").
    Debug.Print Worksheets(i + 1).
    Debug.Print ActiveWorkbook.Worksheets(1).
    Debug.Print ActiveWorkbook.Worksheets(GetIndex()).
    Debug.Print Application.ActiveCell.
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
