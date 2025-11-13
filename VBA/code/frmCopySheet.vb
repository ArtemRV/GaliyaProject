' Variable to hold the button handler
Private ButtonHandler As CommandButtonHandler

Private Sub UserForm_Initialize()
    ' Set initial form properties
    Me.Caption = "Menu"
    Me.Width = 800
    Me.Height = 400
    
    ' Define common settings for controls
    Const LABEL_WIDTH As Integer = 100
    Const LABEL_HEIGHT As Integer = 15
    Const TEXTBOX_WIDTH As Integer = 100
    Const TEXTBOX_HEIGHT As Integer = 20
    Const CHECKBOX_WIDTH As Integer = 20
    Const CHECKBOX_HEIGHT As Integer = 20
    Const MARGIN_LEFT_COL1 As Integer = 20
    Const MARGIN_LEFT_COL2 As Integer = 160
    Const CHECKBOX_OFFSET As Integer = 10
    Const VERTICAL_SPACING As Integer = 25
    Const TOP_OFFSET As Integer = 30
    Const LABEL_TOP As Integer = 10
    
    ' Create label for Source column
    Dim lblSource As MSForms.Label
    Set lblSource = Me.Controls.Add("Forms.Label.1", "lblSource")
    With lblSource
        .Caption = "Source"
        .Left = MARGIN_LEFT_COL1
        .Top = LABEL_TOP
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    
    ' Create label for Target column
    Dim lblTarget As MSForms.Label
    Set lblTarget = Me.Controls.Add("Forms.Label.1", "lblTarget")
    With lblTarget
        .Caption = "Target"
        .Left = MARGIN_LEFT_COL2
        .Top = LABEL_TOP
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    
    ' Create textboxes and checkboxes for both columns
    Dim i As Integer
    Dim txtBox As MSForms.TextBox
    Dim chkBox As MSForms.CheckBox
    
    ' Source column: 12 textboxes and checkboxes
    For i = 1 To 12
        ' Create textbox
        Set txtBox = Me.Controls.Add("Forms.TextBox.1", "txtSource_" & i)
        With txtBox
            .Left = MARGIN_LEFT_COL1
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = TEXTBOX_WIDTH
            .Height = TEXTBOX_HEIGHT
        End With
        
        ' Create checkbox
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkSource_" & i)
        With chkBox
            .Caption = ""
            .Left = MARGIN_LEFT_COL1 + TEXTBOX_WIDTH + CHECKBOX_OFFSET
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = CHECKBOX_WIDTH
            .Height = CHECKBOX_HEIGHT
        End With
    Next i
    
    ' Target column: 12 textboxes and checkboxes
    For i = 1 To 12
        ' Create textbox
        Set txtBox = Me.Controls.Add("Forms.TextBox.1", "txtTarget_" & i)
        With txtBox
            .Left = MARGIN_LEFT_COL2
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = TEXTBOX_WIDTH
            .Height = TEXTBOX_HEIGHT
        End With
        
        ' Create checkbox
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkTarget_" & i)
        With chkBox
            .Caption = ""
            .Left = MARGIN_LEFT_COL2 + TEXTBOX_WIDTH + CHECKBOX_OFFSET
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = CHECKBOX_WIDTH
            .Height = CHECKBOX_HEIGHT
        End With
    Next i
    
    ' Source column: 12 textboxes and checkboxes
    For i = 1 To 12
        ' Create textbox
        Set txtBox = Me.Controls.Add("Forms.TextBox.1", "txtTarget_" & i)
        With txtBox
            .Left = MARGIN_LEFT_COL2
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = TEXTBOX_WIDTH
            .Height = TEXTBOX_HEIGHT
        End With
        
        ' Create checkbox
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkTarget_" & i)
        With chkBox
            .Caption = ""
            .Left = MARGIN_LEFT_COL2 + TEXTBOX_WIDTH + CHECKBOX_OFFSET
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = CHECKBOX_WIDTH
            .Height = CHECKBOX_HEIGHT
        End With
    Next i
    
    ' Auto-fill Source textboxes with worksheet names starting with "N!"
    Dim ws As Worksheet
    Dim txtIndex As Integer
    txtIndex = 1
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 2) = "N!" And txtIndex <= 12 Then
            Me.Controls("txtSource_" & txtIndex).Text = ws.Name
            txtIndex = txtIndex + 1
        End If
    Next ws
    
    ' Create Copy Sheets button to trigger macro
    Dim btnCopySheets As MSForms.CommandButton
    Set btnCopySheets = Me.Controls.Add("Forms.CommandButton.1", "btnCopySheets")
    With btnCopySheets
        .Caption = "Copy Sheets"
        .Left = 170
        .Top = 340
        .Width = 100
        .Height = 30
    End With
    
    ' Initialize button handler to capture Click event
    Set ButtonHandler = New CommandButtonHandler
    Set ButtonHandler.Button = btnCopySheets
End Sub
