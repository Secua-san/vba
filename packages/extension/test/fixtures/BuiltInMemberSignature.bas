Attribute VB_Name = "BuiltInMemberSignature"
Option Explicit

Public Sub Demo()
    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Power(2, 3)
    Debug.Print WorksheetFunction.Average(1, 2, 3)
    Debug.Print WorksheetFunction.EDate(Date, 1)
    Debug.Print WorksheetFunction.EoMonth(Date, 1)
    Debug.Print WorksheetFunction.Find("A", "ABC")
    Debug.Print WorksheetFunction.Search("A", "ABC")
    Debug.Print WorksheetFunction.And(True, False, True)
    Debug.Print WorksheetFunction.Or(True, False, True)
    Debug.Print WorksheetFunction.Xor(True, False, True)
    Debug.Print WorksheetFunction.CountA("A", "")
    Debug.Print WorksheetFunction.CountBlank(Range("A1:A2"))
    Debug.Print WorksheetFunction.Text(Now, "yyyy-mm-dd")
    Debug.Print WorksheetFunction.VLookup("A", Range("A1:B2"), 2, False)
    Call Application.CalculateFull()
    Application.OnTime(Now, "BuiltInMemberSignature.Demo")
    Call Application.WorksheetFunction()
    Call Application.AfterCalculate()
    Call Application.ActiveCell()
    Call Application.NewWorkbook()
    Debug.Print Application.Calculate
End Sub
