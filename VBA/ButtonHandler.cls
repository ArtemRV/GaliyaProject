VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class module to handle events for dynamically created CommandButton
Public WithEvents Button As MSForms.CommandButton
Attribute Button.VB_VarHelpID = -1

Private Sub Button_Click()
    ' Debug: Confirm button click
    MsgBox "Copy Sheets button clicked!", vbInformation
    
    ' Find selected Source checkbox
    Dim i As Integer
    Dim selectedSourceIndex As Integer
    selectedSourceIndex = 0
    For i = 1 To frmCopySheets.m_chkSourceClasses.count
        If frmCopySheets.m_chkSourceClasses(i).chk.Value Then
            selectedSourceIndex = i
            Exit For
        End If
    Next i
    
    If selectedSourceIndex = 0 Then
        MsgBox "Please select one Source checkbox!", vbExclamation
        Exit Sub
    End If
    
    ' Collect data for checked column checkboxes
    Dim sourceName As String
    Dim targetName As String
    Dim sourceCol As String
    Dim j As Integer
    j = 1
    Dim checkedCount As Integer
    checkedCount = 0
    Dim debugMsg As String
    
    On Error Resume Next
    Do
        If frmCopySheets.Controls("chkCol_" & j).Value Then
            checkedCount = checkedCount + 1
            sourceName = frmCopySheets.Controls("txtSource_" & selectedSourceIndex).Text
            targetName = frmCopySheets.Controls("lblContent_" & j).Caption
            sourceCol = frmCopySheets.Controls("lblCol_" & j).Caption
            debugMsg = debugMsg & "Processing: Source=" & sourceName & ", Target=" & targetName & ", SourceCol=" & sourceCol & vbCrLf
            ' Call processing for each checked column
            Call Module1.ProcessSingleCopy(sourceName, targetName, sourceCol)
        End If
        j = j + 1
    Loop While Err.Number = 0
    On Error GoTo 0
    
    ' Debug: Show processed pairs
    If checkedCount = 0 Then
        MsgBox "No column checkboxes selected!", vbExclamation
    Else
        MsgBox "Processed " & checkedCount & " pairs:" & vbCrLf & debugMsg, vbInformation
    End If
    
    ' Close the form
    Unload frmCopySheets
End Sub
