Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.Shapes.
    Debug.Print Sheet1.Shapes(1).
    Debug.Print Sheet1.Shapes("CheckBox1").
    Debug.Print Sheet1.Shapes(i + 1).
    Debug.Print Sheet1.Shapes(GetIndex()).
    Debug.Print Sheet1.Shapes.Item(1).
    Debug.Print Sheet1.Shapes.Item("CheckBox1").
    Debug.Print Sheet1.Shapes.Item(GetIndex()).
    Debug.Print Chart1.Shapes(1).
    Debug.Print Chart1.Shapes.Item(1).
    Debug.Print Sheet1.Shapes("CheckBox1").Name
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.ProgID
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object.Value
    Call Sheet1.Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheet1.Shapes(1).OLEFormat.Object.Value
    Debug.Print Sheet1.Shapes.Item(1).OLEFormat.Object.Value
    Debug.Print Chart1.Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Chart1.Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print Sheet1.Shapes.Range(Array("CheckBox1")).OLEFormat.Object.Value
    Debug.Print Sheet1.Shapes(GetIndex()).Name
    Debug.Print Sheet1.Shapes.Item(GetIndex()).Name
    Debug.Print Sheet1.Shapes("PlainShape").OLEFormat.Object.Value
    Debug.Print Sheet1.Shapes.Item("PlainShape").OLEFormat.Object.Value
    Debug.Print Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Worksheets("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object(1).Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
