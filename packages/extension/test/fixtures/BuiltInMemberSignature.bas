Attribute VB_Name = "BuiltInMemberSignature"
Option Explicit

Public Sub Demo()
    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Power(2, 3)
    Call Application.CalculateFull()
    Application.OnTime(Now, "BuiltInMemberSignature.Demo")
    Debug.Print Application.Calculate
End Sub
