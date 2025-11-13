Option Explicit

Private m_optSourceClasses As Collection
Private m_txtLog As MSForms.TextBox
Private ButtonHandler As ButtonHandler

' Function to get column letter from number
Private Function ColumnLetter(ByVal colNum As Long) As String
    Dim temp As Long
    Dim letter As String
    While colNum > 0
        temp = (colNum - 1) Mod 26
        letter = Chr(temp + 65) & letter
        colNum = (colNum - temp - 1) / 26
    Wend
    ColumnLetter = letter
End Function

' Get Source option buttons collection
Public Function GetSourceOptions() As Collection
    Set GetSourceOptions = m_optSourceClasses
End Function

' Append message to log textbox
Public Sub LogMessage(ByVal msg As String)
    If m_txtLog Is Nothing Then
        On Error Resume Next
        Set m_txtLog = Me.Controls("txtLog")
        On Error GoTo 0
        If m_txtLog Is Nothing Then Exit Sub
    End If
    With m_txtLog
        .Value = .Value & msg & vbCrLf
        .SelStart = Len(.Value) ' scroll to bottom
    End With
End Sub

' Get message from Settings sheet
Public Function GetMessage(ByVal key As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    
    Dim rng As Range
    Set rng = ws.Range("A:B").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        GetMessage = rng.Offset(0, 1).Value
    Else
        GetMessage = key ' return key if not found
    End If
End Function

' Clear the right part of the form (dynamic controls)
Public Sub ClearRightPart()
    On Error Resume Next
    Me.Controls("lblSelectedSource").Delete
    Dim j As Integer
    j = 1
    Do
        Me.Controls("lblContent_" & j).Delete
        Me.Controls("lblCol_" & j).Delete
        Me.Controls("chkCol_" & j).Delete
        j = j + 1
    Loop While Err.Number = 0
    On Error GoTo 0
End Sub

' Create the right part of the form based on selected Source
Public Sub CreateRightPart(ByVal selected_i As Integer)
    Dim sourceName As String
    sourceName = Me.Controls("txtSource_" & selected_i).Text
    
    ' Base offsets for the right part
    Dim baseLeft As Long
    Dim baseTop As Long
    Dim rowSpacing As Long
    
    baseLeft = 450     ' horizontal offset for the right block
    baseTop = 30       ' vertical start offset
    rowSpacing = 25    ' spacing between rows
    
    ' Create label for selected Source
    Dim lblSelectedSource As MSForms.Label
    Set lblSelectedSource = Me.Controls.Add("Forms.Label.1", "lblSelectedSource")
    With lblSelectedSource
        .Caption = sourceName
        .Left = baseLeft
        .Top = 10
        .Width = 100
        .Height = 15
    End With
    
    ' Get Source worksheet
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceName)
    On Error GoTo 0
    If wsSource Is Nothing Then
        LogMessage "? " & GetMessage("SourceNotFound") & " '" & sourceName & "'"
        Exit Sub
    End If
    
    ' Scan row 17 from column H onwards
    Dim column As Integer
    column = 8 ' H
    Dim j As Integer
    j = 1
    While wsSource.Cells(17, column).Value <> ""
        Dim content As String
        content = Trim(wsSource.Cells(17, column).Value)
        Dim colLetter As String
        colLetter = ColumnLetter(column)
        
        ' Create label for content (month name)
        Dim lblContent As MSForms.Label
        Set lblContent = Me.Controls.Add("Forms.Label.1", "lblContent_" & j)
        With lblContent
            .Caption = content
            .Left = baseLeft + 30
            .Top = baseTop + (j - 1) * rowSpacing
            .Width = 100
            .Height = 15
        End With
        
        ' Create label for column letter
        Dim lblCol As MSForms.Label
        Set lblCol = Me.Controls.Add("Forms.Label.1", "lblCol_" & j)
        With lblCol
            .Caption = colLetter
            .Left = baseLeft + 140
            .Top = baseTop + (j - 1) * rowSpacing
            .Width = 20
            .Height = 15
        End With
        
        ' Create checkbox for column
        Dim chkCol As MSForms.CheckBox
        Set chkCol = Me.Controls.Add("Forms.CheckBox.1", "chkCol_" & j)
        With chkCol
            .Caption = ""
            .Left = baseLeft + 170
            .Top = baseTop + (j - 1) * rowSpacing
            .Width = 20
            .Height = 20
        End With
        
        j = j + 1
        column = column + 1
    Wend
End Sub

' Create footer controls dynamically
Private Sub CreateFooterControls()
    ' Constants for layout
    Const BASE_LEFT As Long = 450
    Const BASE_TOP As Long = 100
    Const ROW_SPACING As Long = 25
   
    ' Create footer radio buttons dynamically
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Find the last occupied row in the Settings sheet
    Dim lastRow As Long
    lastRow = wsSettings.Cells(wsSettings.Rows.count, "G").End(xlUp).Row
    
    Dim footerRow As Long
    Dim footerCount As Integer
    footerCount = 0
    Dim footerNames() As String
    ReDim footerNames(1 To 1) ' Initial size, will preserve and resize as needed
    
    ' Scan from row 1 to the last occupied row in column G
    For footerRow = 1 To lastRow
        Dim footerValue As String
        footerValue = Trim(wsSettings.Cells(footerRow, "G").Value)
        If footerValue <> "" And Left(footerValue, 7) = "Footer_" Then
            footerCount = footerCount + 1
            ReDim Preserve footerNames(1 To footerCount)
            ' Extract name after "Footer_"
            footerNames(footerCount) = Mid(footerValue, 8)
        End If
    Next footerRow
   
    ' Create footer labels and radio buttons
    Dim footerTop As Long
    footerTop = BASE_TOP + 2 * ROW_SPACING ' Start below highlight
    Dim i As Integer
    Dim lblFooter As MSForms.Label
    Dim optFooter As MSForms.OptionButton
    For i = 1 To footerCount
        ' Create label for footer name
        Set lblFooter = Me.Controls.Add("Forms.Label.1", "lblFooter_" & i)
        With lblFooter
            .Caption = footerNames(i)
            .Left = BASE_LEFT + 200
            .Top = footerTop + (i - 1) * ROW_SPACING
            .Width = 80
            .Height = 15
        End With
       
        ' Create radio button for footer
        Set optFooter = Me.Controls.Add("Forms.OptionButton.1", "optFooter_" & i)
        With optFooter
            .Caption = ""
            .Left = BASE_LEFT + 290
            .Top = footerTop + (i - 1) * ROW_SPACING
            .Width = 20
            .Height = 20
            .Value = False ' Default none selected
        End With
    Next i
End Sub

' UserForm initialization
Private Sub UserForm_Initialize()
    ' Set initial form properties
    Me.Caption = "Menu"
    Me.Width = 900
    Me.Height = 450
    Me.BackColor = RGB(255, 192, 203) ' pink background
    
    ' === Log window on the left ===
    Set m_txtLog = Me.Controls.Add("Forms.TextBox.1", "txtLog")
    With m_txtLog
        .Left = 10
        .Top = 10
        .Width = 280
        .Height = 400
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = &HFFFFFF
    End With
    
    ' Constants for right part layout
    Const LABEL_WIDTH As Integer = 100
    Const LABEL_HEIGHT As Integer = 15
    Const TEXTBOX_WIDTH As Integer = 100
    Const TEXTBOX_HEIGHT As Integer = 20
    Const OPTION_WIDTH As Integer = 20
    Const OPTION_HEIGHT As Integer = 20
    Const MARGIN_LEFT_COL1 As Integer = 320
    Const OPTION_OFFSET As Integer = 10
    Const VERTICAL_SPACING As Integer = 25
    Const TOP_OFFSET As Integer = 30
    Const LABEL_TOP As Integer = 10
    Const BASE_LEFT As Long = 450 ' horizontal offset for the right block
    Const BASE_TOP As Long = 30 ' vertical start offset
    Const ROW_SPACING As Long = 25 ' spacing between rows
    
    ' Create label for Source column
    Dim lblSource As MSForms.Label
    Set lblSource = Me.Controls.Add("Forms.Label.1", "lblSource")
    With lblSource
        .Caption = GetMessage("SourceLabel") ' e.g. "Source"
        .Left = MARGIN_LEFT_COL1
        .Top = LABEL_TOP
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    
    ' Create debug mode label and checkbox
    Dim lblDebug As MSForms.Label
    Set lblDebug = Me.Controls.Add("Forms.Label.1", "lblDebug")
    With lblDebug
        .Caption = GetMessage("DebugLabel") ' e.g., "Debug Mode"
        .Left = BASE_LEFT + 200
        .Top = BASE_TOP
        .Width = 80
        .Height = 15
    End With
    
    Dim chkDebug As MSForms.CheckBox
    Set chkDebug = Me.Controls.Add("Forms.CheckBox.1", "chkDebug")
    With chkDebug
        .Caption = ""
        .Left = BASE_LEFT + 290
        .Top = BASE_TOP
        .Width = 20
        .Height = 20
        .Value = False ' Default: debug mode off
        ' Log initial state
        LogMessage GetMessage("DebugModeSetTo") & .Value
    End With
    
    ' Create cell highlighting label and checkbox
    Dim lblHighlight As MSForms.Label
    Set lblHighlight = Me.Controls.Add("Forms.Label.1", "lblHighlight")
    With lblHighlight
        .Caption = GetMessage("HighlightLabel") ' e.g., "Highlight Cells"
        .Left = BASE_LEFT + 200
        .Top = BASE_TOP + ROW_SPACING
        .Width = 80
        .Height = 15
    End With
    
    Dim chkHighlight As MSForms.CheckBox
    Set chkHighlight = Me.Controls.Add("Forms.CheckBox.1", "chkHighlight")
    With chkHighlight
        .Caption = ""
        .Left = BASE_LEFT + 290
        .Top = BASE_TOP + ROW_SPACING
        .Width = 20
        .Height = 20
        .Value = True ' Default: highlight cells on
        ' Log initial state
        LogMessage GetMessage("HighlightCellsSetTo") & .Value
    End With

    ' Create footer controls
    Call CreateFooterControls
    
    ' Count sheets starting with "N!"
    Dim ws As Worksheet
    Dim count As Integer
    count = 0
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 2) = "N!" Then
            count = count + 1
        End If
    Next ws
    
    ' Collect sheet names
    Dim sourceNames() As String
    ReDim sourceNames(1 To count)
    Dim k As Integer
    k = 1
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 2) = "N!" Then
            sourceNames(k) = ws.Name
            k = k + 1
        End If
    Next ws
    
    ' Initialize Source option buttons classes
    Set m_optSourceClasses = New Collection
    
    ' Create Source textboxes and option buttons
    Dim i As Integer
    Dim txtBox As MSForms.TextBox
    Dim optBtn As MSForms.OptionButton
    Dim classInst As OptionButtonClass
    For i = 1 To count
        ' Create textbox
        Set txtBox = Me.Controls.Add("Forms.TextBox.1", "txtSource_" & i)
        With txtBox
            .Left = MARGIN_LEFT_COL1
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = TEXTBOX_WIDTH
            .Height = TEXTBOX_HEIGHT
            .Text = sourceNames(i)
        End With
        
        ' Create option button
        Set optBtn = Me.Controls.Add("Forms.OptionButton.1", "optSource_" & i)
        With optBtn
            .Caption = ""
            .Left = MARGIN_LEFT_COL1 + TEXTBOX_WIDTH + OPTION_OFFSET
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = OPTION_WIDTH
            .Height = OPTION_HEIGHT
            .Tag = i ' Store index
        End With
        
        ' Add class instance for event handling
        Set classInst = New OptionButtonClass
        Set classInst.opt = optBtn
        m_optSourceClasses.Add classInst
    Next i
    
    ' Create Copy Sheets button
    Dim btnCopySheets As MSForms.CommandButton
    Set btnCopySheets = Me.Controls.Add("Forms.CommandButton.1", "btnCopySheets")
    With btnCopySheets
        .Caption = GetMessage("CopySheetsButton") ' e.g. "Copy Sheets"
        .Left = 330
        .Top = 380
        .Width = 100
        .Height = 30
    End With
    
    ' Initialize button handler
    Set ButtonHandler = New ButtonHandler
    Set ButtonHandler.Button = btnCopySheets
    
    ' Initial log message
    LogMessage GetMessage("FormInitialized")
End Sub

' Handle checkbox change events to log state
Private Sub chkDebug_Change()
    If Not Me.Controls("chkDebug") Is Nothing Then
        LogMessage GetMessage("DebugModeSetTo") & Me.Controls("chkDebug").Value
    End If
End Sub

Private Sub chkHighlight_Change()
    If Not Me.Controls("chkHighlight") Is Nothing Then
        LogMessage GetMessage("HighlightCellsSetTo") & Me.Controls("chkHighlight").Value
    End If
End Sub

' Get selected footer name
Public Function GetSelectedFooter() As String
    Dim i As Integer
    For i = 1 To 3
        On Error Resume Next
        If Me.Controls("optFooter_" & i).Value Then
            GetSelectedFooter = Me.Controls("lblFooter_" & i).Caption
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next i
    GetSelectedFooter = "" ' None selected
End Function
