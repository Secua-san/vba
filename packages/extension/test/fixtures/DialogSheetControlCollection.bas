Attribute VB_Name = "DialogSheetControlCollection"
Option Explicit

Public Sub Demo()
    Dim index As Long
    Debug.Print DialogSheets(1).Buttons.
    Debug.Print DialogSheets(1).Buttons(1).
    Debug.Print DialogSheets(1).Buttons(&H1).
    Debug.Print DialogSheets(1).Buttons(&O7).
    Debug.Print DialogSheets(1).Buttons(1#).
    Debug.Print DialogSheets(1).Buttons(1E+2).
    Debug.Print DialogSheets(1).Buttons("Button 1").
    Debug.Print DialogSheets(1).Buttons(index).
    Debug.Print DialogSheets(1).Buttons(Array(1, 2)).
    Debug.Print DialogSheets(1).Buttons.Item(1).
    Debug.Print DialogSheets(1).Buttons.Item(index).
    Debug.Print DialogSheets(1).CheckBoxes(1).
    Debug.Print DialogSheets(1).OptionButtons("Option 1").
    Call DialogSheets(1).Buttons(1).Select("Button 1")
    Call DialogSheets(1).Buttons.Item(1).Select("Button 1")
    Call DialogSheets(1).CheckBoxes(1).Select("Check 1")
    Call DialogSheets(1).CheckBoxes.Item(1).Select("Check 1")
    Call DialogSheets(1).OptionButtons("Option 1").Select("Option 1")
    Call DialogSheets(1).OptionButtons.Item(1).Select("Option 1")
    Call Application.DialogSheets(1).Buttons(1).Select("Button 1")
    Debug.Print DialogSheets(1).Buttons(1).Caption
    Debug.Print DialogSheets(1).CheckBoxes(1).Value
    Debug.Print DialogSheets(1).OptionButtons("Option 1").Value
    Debug.Print DialogSheets(1).CheckBoxes(1).Value(
    Debug.Print DialogSheets(1).OptionButtons("Option 1").Value(
    Call DialogSheets(1).Buttons(index).Select("Button 1")
    Call DialogSheets(1).Buttons.Item(index).Select("Button 1")
    Call DialogSheets(1).Buttons(Array(1, 2)).Select("Button 1")
    Debug.Print DialogSheets(1).Buttons.Item("Button 1").
    Debug.Print DialogSheets(1).CheckBoxes.Item("Check 1").
    Debug.Print DialogSheets(1).OptionButtons.Item("Option 1").
    Call DialogSheets(1).Buttons(&H1).Select("Button 1")
    Call DialogSheets(1).Buttons(&O7).Select("Button 1")
    Call DialogSheets(1).Buttons(1#).Select("Button 1")
    Call DialogSheets(1).Buttons(1E+2).Select("Button 1")
    Call DialogSheets(1).Buttons.Item("Button 1").Select("Button 1")
    Call DialogSheets(1).CheckBoxes.Item("Check 1").Select("Check 1")
    Call DialogSheets(1).OptionButtons.Item("Option 1").Select("Option 1")
End Sub
