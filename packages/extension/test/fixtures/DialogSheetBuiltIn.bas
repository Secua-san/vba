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
    Debug.Print Application.DialogSheets.
    Debug.Print Application.DialogSheets(1).Evaluate("A1")
    Call ActiveWorkbook.DialogSheets(1).SaveAs("Dialog1.xlsx")
    Call Application.DialogSheets(Array("Dialog1", "Dialog2")).SaveAs("Dialog1.xlsx")
    Call ThisWorkbook.DialogSheets(1).SaveAs("Dialog1.xlsx")
    Debug.Print DialogSheets(1).DialogFrame.
    Debug.Print DialogSheets(1).DialogFrame.Caption
    Call DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Debug.Print Application.DialogSheets(1).DialogFrame.
    Debug.Print Application.DialogSheets(1).DialogFrame.Text
    Call ActiveWorkbook.DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Call ThisWorkbook.DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Call DialogSheets(Array("Dialog1", "Dialog2")).DialogFrame.Select("DialogFrame1")
    Debug.Print DialogSheets("Dialog1").DialogFrame.
    Call DialogSheets.Item(1).DialogFrame.Select("DialogFrame1")
    Debug.Print DialogSheets(1).DialogFrame.Caption(
    Debug.Print Application.DialogSheets(1).DialogFrame.Text(
End Sub
