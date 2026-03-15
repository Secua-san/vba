Attribute VB_Name = "WorksheetControlCodeName"
Option Explicit

Public Sub Demo()
    Debug.Print Sheet1.chkFinished.
    Debug.Print Sheet1.CheckBox1.
    Debug.Print Chart1.chkFinished.
    Debug.Print ActiveSheet.chkFinished.
    Debug.Print Sheet1.chkFinished.Value
    Debug.Print Sheet1.CheckBox1.Value
    Debug.Print Chart1.chkFinished.Value
    Debug.Print ActiveSheet.chkFinished.Value
    Call Sheet1.chkFinished.Select(
    Call Sheet1.CheckBox1.Select(
    Call Chart1.chkFinished.Select(
    Call ActiveSheet.chkFinished.Select(
End Sub
