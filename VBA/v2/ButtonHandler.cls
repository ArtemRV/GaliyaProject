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
    ' Log the click
    frmCopySheets.LogMessage "Copy Sheets button clicked."

    ' Get Source options
    Dim optSourceClasses As Collection
    Set optSourceClasses = frmCopySheets.GetSourceOptions

    ' Find selected Source option
    Dim i As Long
    Dim selectedSourceIndex As Long
    selectedSourceIndex = 0
    For i = 1 To optSourceClasses.count
        If optSourceClasses(i).opt.Value Then
            selectedSourceIndex = i
            Exit For
        End If
    Next i

    If selectedSourceIndex = 0 Then
        frmCopySheets.LogMessage "Please select one Source option!"
        Exit Sub
    End If

    ' Process checked month checkboxes exactly once each
    Dim ctrl As MSForms.Control
    Dim checkedCount As Long
    checkedCount = 0
    Dim debugMsg As String

    For Each ctrl In frmCopySheets.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If Left(ctrl.Name, 7) = "chkCol_" Then
                If ctrl.Value = True Then
                    checkedCount = checkedCount + 1
                    Dim idx As Long
                    idx = CLng(Mid(ctrl.Name, 8)) ' number part of name
                    Dim sourceName As String
                    sourceName = frmCopySheets.Controls("txtSource_" & selectedSourceIndex).Text
                    Dim targetName As String
                    targetName = frmCopySheets.Controls("lblContent_" & idx).Caption
                    Dim sourceCol As String
                    sourceCol = frmCopySheets.Controls("lblCol_" & idx).Caption

                    debugMsg = debugMsg & "Processing: Source=" & sourceName & ", Target=" & targetName & ", SourceCol=" & sourceCol & vbCrLf
                    frmCopySheets.LogMessage "Processing: " & sourceName & " -> " & targetName & " (" & sourceCol & ")"

                    ' Call processing exactly once for this checked checkbox
                    Module1.ProcessSingleCopy sourceName, targetName, sourceCol
                End If
            End If
        End If
    Next ctrl

    If checkedCount = 0 Then
        frmCopySheets.LogMessage "No column checkboxes selected!"
    Else
        frmCopySheets.LogMessage "Processed " & checkedCount & " pairs."
        frmCopySheets.LogMessage debugMsg
    End If
    
End Sub

