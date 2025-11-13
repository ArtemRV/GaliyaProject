Attribute VB_Name = "Module1"
Option Explicit

' Main procedure to copy a single source column to a target sheet
Public Sub ProcessSingleCopy(sourceName As String, targetName As String, sourceCol As String)
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

    On Error GoTo ErrHandler

    ' Speed optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Get settings
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    startRow = CLng(GetSetting("StartRow"))
    targetCol = Trim(GetSetting("TargetCol"))
    highlightCells = (LCase(Trim(wsSettings.Range("F6").Value)) = "yes")
    debugMode = (LCase(Trim(wsSettings.Range("G6").Value)) = "yes")

    ' Get source sheet
    Set wsSource = ThisWorkbook.Sheets(sourceName)
    sourceColNum = wsSource.Columns(sourceCol).column
    targetColNum = wsSource.Columns(targetCol).column

    ' Delete target sheet if it exists
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = targetName Then
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next sh

    ' Copy source sheet and rename
    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set wsTarget = ActiveSheet
    wsTarget.Name = targetName

    If debugMode Then
        MsgBox "Processing Source=" & sourceName & ", Target=" & targetName & ", Column=" & sourceCol, vbInformation
    End If

    ' Determine last row
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    ' Main processing loop
    r = startRow
    Do While r <= lastRowSource
        cellA = Trim(wsTarget.Cells(r, "A").Value)

        ' Copy value if present
        If wsSource.Cells(r, sourceColNum).Value <> "" Then
            wsTarget.Cells(r, targetColNum).Value = wsSource.Cells(r, sourceColNum).Value

            ' Multiply by F and write to G
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

        ' Group handling
        ElseIf Application.WorksheetFunction.CountA(wsTarget.Range("A" & r & ":G" & r)) > 0 Then
            If wsTarget.Rows(r).OutlineLevel = 1 And IsNumeric(cellA) Then
                groupStart = r
                groupEnd = r

                Do While groupEnd < wsTarget.Rows.count And _
                        wsTarget.Rows(groupEnd + 1).OutlineLevel > wsTarget.Rows(groupStart).OutlineLevel
                    groupEnd = groupEnd + 1
                Loop

                If debugMode Then
                    wsTarget.Rows(groupStart & ":" & groupEnd).Interior.Color = vbRed
                Else
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

    ' Delete collected rows
    If Not debugMode Then
        If Not delRange Is Nothing Then
            delRange.Delete
            Set delRange = Nothing
        End If
    End If

    ' Delete columns H and onward
    On Error Resume Next
    lastColTarget = wsTarget.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).column
    On Error GoTo ErrHandler
    If lastColTarget >= 8 Then
        wsTarget.Range(wsTarget.Columns(8), wsTarget.Columns(lastColTarget)).Delete
    End If

    ' Sum sections
    Call SumSections(wsTarget, debugMode)

    ' Scroll to top
    With wsTarget
        .Activate
        Application.GoTo .Range("A1"), False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End With

    If debugMode Then MsgBox "Completed: Target=" & targetName, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' Get setting from Settings sheet
Private Function GetSetting(ByVal key As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    
    Dim rng As Range
    Set rng = ws.Range("D:E").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        GetSetting = rng.Offset(0, 1).Value ' take value from column E
    Else
        MsgBox "Setting '" & key & "' not found in Settings sheet!", vbCritical
        GetSetting = "" ' return empty string if not found
    End If
End Function

' Sum sections in column G based on headers in column B
Public Sub SumSections(wsTarget As Worksheet, Optional debugMode As Boolean = False)
    Dim lastRow As Long
    Dim r As Long
    Dim sectionRow As Long
    Dim sectionSum As Double
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[^\d\s]+\s\d+\.$"
    regex.IgnoreCase = True

    If debugMode Then MsgBox "Start SumSections", vbInformation

    lastRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row
    sectionRow = 0
    sectionSum = 0

    For r = 1 To lastRow
        Dim cellValue As String
        cellValue = Trim(wsTarget.Cells(r, "B").Value)

        If regex.Test(cellValue) And wsTarget.Rows(r).OutlineLevel = 1 Then
            If sectionRow > 0 Then
                If wsTarget.Rows(sectionRow).OutlineLevel = 1 Then
                    wsTarget.Cells(sectionRow, "G").Value = sectionSum
                End If
            End If
            sectionRow = r
            sectionSum = 0
        Else
            If wsTarget.Rows(r).OutlineLevel = 1 And IsNumeric(wsTarget.Cells(r, "G").Value) Then
                sectionSum = sectionSum + wsTarget.Cells(r, "G").Value
            End If
        End If
    Next r

    If sectionRow > 0 Then
        If wsTarget.Rows(sectionRow).OutlineLevel = 1 Then
            wsTarget.Cells(sectionRow, "G").Value = sectionSum
        End If
    End If

    If debugMode Then MsgBox "SumSections completed!", vbInformation
End Sub


