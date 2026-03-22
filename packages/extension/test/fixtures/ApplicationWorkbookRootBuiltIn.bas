Option Explicit

Public Sub Demo()
    ' static current-bundle family
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
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(

    ' active-workbook broad-root family
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.
    Call Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(

    ' non-target roots
    Debug.Print Application.Caller.OLEObjects("CheckBox1").Object.
    Debug.Print Application.Caller.OLEObjects("CheckBox1").Object.Value
    Call Application.Caller.OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
