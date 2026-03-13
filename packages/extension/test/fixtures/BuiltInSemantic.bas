Attribute VB_Name = "BuiltInSemantic"
Option Explicit

Public Sub Demo()
    Beep
    MsgBox xlAll
    Debug.Print Application.Name
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print ThisWorkbook.SaveAs
    Debug.Print Application.ActiveCell.Address
End Sub
