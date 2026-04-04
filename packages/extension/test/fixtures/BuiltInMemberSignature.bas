Attribute VB_Name = "BuiltInMemberSignature"
Option Explicit

Public Sub Demo()
    Dim i As Long
    Dim transposedResult As Variant

    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Power(2, 3)
    Debug.Print WorksheetFunction.Average(1, 2, 3)
    Debug.Print WorksheetFunction.Max(1, 2, 3)
    Debug.Print WorksheetFunction.Min(1, 2, 3)
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
    Debug.Print WorksheetFunction.Match("A", Range("A1:A2"), 0)
    Debug.Print WorksheetFunction.Index(Range("A1:B2"), 1, 2)
    Debug.Print WorksheetFunction.Lookup("A", Range("A1:A2"), Range("B1:B2"))
    Debug.Print WorksheetFunction.HLookup("A", Range("A1:B2"), 2, False)
    Debug.Print WorksheetFunction.Choose(1, "A", "B")
    transposedResult = WorksheetFunction.Transpose(Range("A1:B2"))
    Debug.Print UBound(transposedResult, 1), UBound(transposedResult, 2)
    Debug.Print ActiveCell.Address(False, False, xlA1, False)
    Debug.Print Application.ActiveCell.Address(False, False, xlA1, False)
    Debug.Print Cells.AddressLocal(False, False)
    Debug.Print ActiveCell.HasSpill
    Debug.Print ActiveCell.SavedAsArray
    Debug.Print ActiveCell.SpillParent.Address(False, False)
    Debug.Print ActiveWorkbook.Worksheets.Count
    Debug.Print Worksheets(1).Evaluate("A1")
    Debug.Print Worksheets("A(1)").Evaluate("A1")
    Debug.Print Worksheets(Array("Sheet1", "Sheet2")).Evaluate("A1")
    Call Worksheets(1).SaveAs("Sheet1.csv")
    Call Worksheets(i + 1).SaveAs("Sheet1.csv")
    Call ActiveWorkbook.Worksheets(1).ExportAsFixedFormat(xlTypePDF)
    Call ActiveWorkbook.Worksheets(GetIndex()).ExportAsFixedFormat(xlTypePDF)
    Debug.Print ThisWorkbook.SaveAs
    Call ThisWorkbook.SaveAs("Book1.xlsx")
    Debug.Print Sheet1.Evaluate("A1")
    Call Sheet1.SaveAs("Sheet1.csv")
    Call Chart1.SetSourceData(Range("A1:B2"))
    Debug.Print Chart1.ChartArea
    Call ActiveWorkbook.Close(False)
    Call ActiveWorkbook.ExportAsFixedFormat(xlTypePDF)
    Call Application.CalculateFull()
    Application.OnTime(Now, "BuiltInMemberSignature.Demo")
    Call Application.WorksheetFunction()
    Call Application.AfterCalculate()
    Call Application.ActiveCell()
    Call Application.NewWorkbook()
    Debug.Print Application.Calculate
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
