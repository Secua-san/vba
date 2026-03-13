Attribute VB_Name = "BuiltInSemantic"
Option Explicit

Public Sub Demo()
    Beep
    MsgBox xlAll
    Debug.Print Application.Name
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Worksheets("A(1)").Evaluate("A1")
    Debug.Print Worksheets(Array("Sheet1", "Sheet2")).Evaluate("A1")
    Debug.Print ThisWorkbook.SaveAs
    Debug.Print Sheet1.Evaluate("A1")
    Debug.Print Application.ActiveCell.Address
End Sub
