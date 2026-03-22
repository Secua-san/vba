Option Explicit

Private Type Application
    Name As String
End Type

Public Sub ShadowedApplication()
    ' shadowed Application qualifier
    Dim Application As Application
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub
