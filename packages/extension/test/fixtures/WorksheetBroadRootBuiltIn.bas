Attribute VB_Name = "WorksheetBroadRootBuiltIn"
Option Explicit

Public Sub Demo()
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Call Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveSheet.OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
