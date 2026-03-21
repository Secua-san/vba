Attribute VB_Name = "OleObjectBuiltIn"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.OLEObjects.
    Debug.Print Sheet1.OLEObjects(1).
    Debug.Print Sheet1.OLEObjects("CheckBox1").
    Debug.Print Sheet1.OLEObjects(i + 1).
    Debug.Print Sheet1.OLEObjects(GetIndex()).
    Debug.Print Sheet1.OLEObjects.Item(1).
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").
    Debug.Print Sheet1.OLEObjects.Item(i + 1).
    Debug.Print Sheet1.OLEObjects.Item(GetIndex()).
    Debug.Print Chart1.OLEObjects(1).
    Debug.Print Chart1.OLEObjects.Item(1).
    Call Sheet1.OLEObjects(1).Activate(
    Call Sheet1.OLEObjects(GetIndex()).Activate(
    Call Sheet1.OLEObjects.Item(1).Activate(
    Call Sheet1.OLEObjects.Item(GetIndex()).Activate(
    Call Sheet1.OLEObjects("CheckBox1").Object.Select(
    Call Sheet1.OLEObjects.Item("CheckBox1").Object.Select(
    Call Chart1.OLEObjects("CheckBox1").Object.Select(
    Debug.Print Sheet1.OLEObjects(1).Name
    Debug.Print Sheet1.OLEObjects.Item(1).Name
    Debug.Print Sheet1.OLEObjects(1).Object.
    Debug.Print Sheet1.OLEObjects("CheckBox1").Object.
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").Object.
    Debug.Print Chart1.OLEObjects("CheckBox1").Object.
    Debug.Print ActiveSheet.OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print Sheet1.OLEObjects("CheckBox1").Object.Value
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
