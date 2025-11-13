Option Explicit

' Class module to handle events for dynamically created CommandButton
Public WithEvents Button As MSForms.CommandButton

Private Sub Button_Click()
    ' Debug: Confirm button click
    MsgBox "Copy Sheets button clicked!", vbInformation
    
    ' Collect data from form
    Dim sourceValues(1 To 12) As String
    Dim targetValues(1 To 12) As String
    Dim sourceCheckboxes(1 To 12) As Boolean
    Dim i As Integer
    Dim debugMsg As String
    For i = 1 To 12
        sourceValues(i) = frmCopySheets.Controls("txtSource_" & i).Text
        targetValues(i) = frmCopySheets.Controls("txtTarget_" & i).Text
        sourceCheckboxes(i) = frmCopySheets.Controls("chkSource_" & i).Value
        debugMsg = debugMsg & "Couple " & i & ": Source=" & sourceValues(i) & ", Target=" & targetValues(i) & ", Checkbox=" & sourceCheckboxes(i) & vbCrLf
    Next i
    
    ' Debug: Show form data
    MsgBox "Form data:" & vbCrLf & debugMsg, vbInformation
    
    ' Debug: Confirm macro call attempt
    MsgBox "Attempting to call Module1.MainCopySheets", vbInformation
    
    ' Call the macro in Module1 with error handling
    On Error GoTo ErrorHandler
    Call Module1.MainCopySheets(sourceValues, targetValues, sourceCheckboxes)
    MsgBox "MainCopySheets executed successfully!", vbInformation
    GoTo ExitSub
    
ErrorHandler:
    MsgBox "Error executing macro: " & Err.Description, vbCritical
    Resume ExitSub
    
ExitSub:
    ' Close the form
    Unload frmCopySheets
End Sub
