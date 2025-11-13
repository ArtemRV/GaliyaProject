Attribute VB_Name = "Module1"
Option Explicit

Public Sub ProcessSingleCopy(sourceName As String, targetName As String, sourceCol As String)
    ' Debug: Confirm single copy start
    MsgBox "Start of ProcessSingleCopy for Source=" & sourceName & ", Target=" & targetName & ", SourceCol=" & sourceCol, vbInformation
    
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wsSettings As Worksheet
    Dim lastRowSource As Long
    Dim lastColTarget As Long
    Dim startRow As Long
    Dim targetCol As String
    Dim sourceColNum As Long, targetColNum As Long
    Dim r As Long
    Dim highlightCells As Boolean
    Dim debugMode As Boolean
    Dim groupStart As Long, groupEnd As Long
    Dim cellA As String
    Dim delRange As Range
    
    ' Speed optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Check if Settings sheet exists
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    If wsSettings Is Nothing Then
        MsgBox "Settings sheet not found!", vbCritical
        GoTo ExitSub
    End If
    
    ' Read parameters from Settings
    startRow = wsSettings.Range("D6").Value
    targetCol = wsSettings.Range("E6").Value
    highlightCells = (LCase(wsSettings.Range("F6").Value) = "yes")
    debugMode = (LCase(wsSettings.Range("G6").Value) = "yes")
    
    targetColNum = Columns(targetCol & ":" & targetCol).column
    
    ' Get Source worksheet
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceName)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sourceName & "' not found!", vbExclamation
        GoTo ExitSub
    End If
    
    sourceColNum = Columns(sourceCol & ":" & sourceCol).column
    
    ' Copy source sheet
    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.Name = targetName
    Set wsTarget = ActiveSheet
    
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    r = startRow
    Do While r <= lastRowSource
        cellA = Trim(wsTarget.Cells(r, "A").Value)
        
        ' Case 1: copy value if exists
        If wsSource.Cells(r, sourceColNum).Value <> "" Then
            wsTarget.Cells(r, targetColNum).Value = wsSource.Cells(r, sourceColNum).Value
            
            ' Multiply by value from column F and write to G
            If IsNumeric(wsTarget.Cells(r, targetColNum).Value) And IsNumeric(wsTarget.Cells(r, 6).Value) Then
                wsTarget.Cells(r, 7).Value = wsTarget.Cells(r, targetColNum).Value * wsTarget.Cells(r, 6).Value
            Else
                wsTarget.Cells(r, 7).Value = 0
                wsTarget.Cells(r, 7).Interior.Color = vbRed
            End If
            
            ' Highlight if needed
            If highlightCells Then
                wsTarget.Cells(r, targetColNum).Interior.Color = wsSource.Cells(r, sourceColNum).Interior.Color
            End If
            
            r = r + 1
            
        ' Case 2: check group deletion
        ElseIf Application.WorksheetFunction.CountA(wsTarget.Range("A" & r & ":G" & r)) > 0 Then
            If wsTarget.Rows(r).OutlineLevel = 1 And IsNumeric(cellA) Then
                groupStart = r
                groupEnd = r
                
                ' Expand group until outline closes
                Do While groupEnd < wsTarget.Rows.count And _
                        wsTarget.Rows(groupEnd + 1).OutlineLevel > wsTarget.Rows(groupStart).OutlineLevel
                    groupEnd = groupEnd + 1
                Loop
                
                If debugMode Then
                    ' Debug mode: highlight red
                    wsTarget.Rows(groupStart & ":" & groupEnd).Interior.Color = vbRed
                Else
                    ' Collect ranges for bulk deletion
                    If delRange Is Nothing Then
                        Set delRange = wsTarget.Rows(groupStart & ":" & groupEnd)
                    Else
                        Set delRange = Union(delRange, wsTarget.Rows(groupStart & ":" & groupEnd))
                    End If
                End If
                
                r = groupEnd + 1
            Else
                r = r + 1
            End If
        Else
            r = r + 1
        End If
    Loop
    
    ' Bulk delete all collected ranges
    If Not debugMode Then
        If Not delRange Is Nothing Then
            delRange.Delete
            Set delRange = Nothing
        End If
    End If
    
    ' Delete unnecessary columns H and onward
    On Error Resume Next
    lastColTarget = wsTarget.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).column
    On Error GoTo 0
    
    If lastColTarget >= 8 Then
        wsTarget.Range(wsTarget.Columns(8), wsTarget.Columns(lastColTarget)).Delete
    End If
    
    ' Sum sections
    Call SumSections(wsTarget)
    
    ' Scroll to top
    With wsTarget
        .Activate
        Application.Goto .Range("A1"), False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End With
    
ExitSub:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Debug: Confirm single copy completion
    MsgBox "ProcessSingleCopy completed for Target=" & targetName, vbInformation
End Sub

Public Sub SumSections(wsTarget As Worksheet)
    ' Debug: Confirm SumSections start
    MsgBox "Start of SumSections", vbInformation
    
    Dim lastRow As Long
    Dim r As Long
    Dim sectionRow As Long
    Dim sectionSum As Double
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Pattern: any word + space + digits + dot
    regex.Pattern = "^[^\d\s]+\s\d+\.$"
    regex.IgnoreCase = True
    
    lastRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row
    sectionRow = 0
    sectionSum = 0
    
    For r = 1 To lastRow
        Dim cellValue As String
        cellValue = Trim(wsTarget.Cells(r, "B").Value)
        
        ' New section header only if OutlineLevel = 1
        If regex.Test(cellValue) And wsTarget.Rows(r).OutlineLevel = 1 Then
            ' Write sum to previous section if valid
            If sectionRow > 0 Then
                If wsTarget.Rows(sectionRow).OutlineLevel = 1 Then
                    wsTarget.Cells(sectionRow, "G").Value = sectionSum
                End If
            End If
            ' Start new section
            sectionRow = r
            sectionSum = 0
        Else
            ' Accumulate numeric values in G only for OutlineLevel = 1
            If wsTarget.Rows(r).OutlineLevel = 1 And IsNumeric(wsTarget.Cells(r, "G").Value) Then
                sectionSum = sectionSum + wsTarget.Cells(r, "G").Value
            End If
        End If
    Next r
    
    ' Write sum for the last section
    If sectionRow > 0 Then
        If wsTarget.Rows(sectionRow).OutlineLevel = 1 Then
            wsTarget.Cells(sectionRow, "G").Value = sectionSum
        End If
    End If
    
    ' Debug: Confirm SumSections completion
    MsgBox "SumSections completed!", vbInformation
End Sub

