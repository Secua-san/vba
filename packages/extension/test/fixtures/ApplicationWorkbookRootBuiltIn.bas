Option Explicit

Public Sub Demo()
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function

Private Type Application
    Name As String
End Type

Public Sub ShadowedApplication()
    Dim Application As Application
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub
