Attribute VB_Name = "OleObjectBuiltIn"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.OLEObjects.
    Debug.Print Sheet1.OLEObjects(1).
    Debug.Print Sheet1.OLEObjects("CheckBox1").
    Debug.Print Sheet1.OLEObjects(i + 1).
    Debug.Print Sheet1.OLEObjects(GetIndex()).
    Debug.Print Chart1.OLEObjects(1).
    Call Sheet1.OLEObjects(1).Activate(
    Call Sheet1.OLEObjects(GetIndex()).Activate(
    Debug.Print Sheet1.OLEObjects(1).Name
    Debug.Print Sheet1.OLEObjects(1).Object.
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function
