Attribute VB_Name = "DialogSheetBuiltIn"
Option Explicit

Public Sub Demo()
    Debug.Print DialogSheets.
    Debug.Print DialogSheets(1).
    Debug.Print DialogSheets("Dialog1").
    Debug.Print DialogSheets(Array("Dialog1", "Dialog2")).
    Debug.Print DialogSheets(1).Evaluate("A1")
    Call DialogSheets(1).SaveAs("Dialog1.xlsx")
    Call DialogSheets(1).ExportAsFixedFormat(xlTypePDF)
    Call DialogSheets(Array("Dialog1", "Dialog2")).SaveAs("Dialog1.xlsx")
    Call DialogSheets.Item(1).SaveAs("Dialog1.xlsx")
End Sub
