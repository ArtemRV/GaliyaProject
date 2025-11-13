VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopySheets 
   Caption         =   "Copy Sheets"
   ClientHeight    =   7440
   ClientLeft      =   -84
   ClientTop       =   -312
   ClientWidth     =   15780
   OleObjectBlob   =   "frmCopySheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCopySheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Collection for Source checkboxes classes
Private m_chkSourceClasses As Collection

' Variable to hold the button handler
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

' Get Source checkboxes collection
Public Function GetSourceCheckboxes() As Collection
    Set GetSourceCheckboxes = m_chkSourceClasses
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
    
    ' Create label for selected Source
    Dim lblSelectedSource As MSForms.Label
    Set lblSelectedSource = Me.Controls.Add("Forms.Label.1", "lblSelectedSource")
    With lblSelectedSource
        .Caption = sourceName
        .Left = 230
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
        MsgBox "Source sheet '" & sourceName & "' not found!", vbExclamation
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
        
        ' Create label for content (e.g., "September")
        Dim lblContent As MSForms.Label
        Set lblContent = Me.Controls.Add("Forms.Label.1", "lblContent_" & j)
        With lblContent
            .Caption = content
            .Left = 260
            .Top = 30 + (j - 1) * 25
            .Width = 100
            .Height = 15
        End With
        
        ' Create label for column name (e.g., "H")
        Dim lblCol As MSForms.Label
        Set lblCol = Me.Controls.Add("Forms.Label.1", "lblCol_" & j)
        With lblCol
            .Caption = colLetter
            .Left = 370
            .Top = 30 + (j - 1) * 25
            .Width = 20
            .Height = 15
        End With
        
        ' Create checkbox for column
        Dim chkCol As MSForms.CheckBox
        Set chkCol = Me.Controls.Add("Forms.CheckBox.1", "chkCol_" & j)
        With chkCol
            .Caption = ""
            .Left = 400
            .Top = 30 + (j - 1) * 25
            .Width = 20
            .Height = 20
        End With
        
        j = j + 1
        column = column + 1
    Wend
End Sub

Private Sub UserForm_Initialize()
    ' Debug: Confirm form initialization
    MsgBox "Form initializing!", vbInformation
    
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
    
    ' Initialize Source checkboxes classes
    Set m_chkSourceClasses = New Collection
    
    ' Create Source textboxes and checkboxes
    Dim i As Integer
    Dim txtBox As MSForms.TextBox
    Dim chkBox As MSForms.CheckBox
    Dim classInst As CheckBoxClass
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
        
        ' Create checkbox
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkSource_" & i)
        With chkBox
            .Caption = ""
            .Left = MARGIN_LEFT_COL1 + TEXTBOX_WIDTH + CHECKBOX_OFFSET
            .Top = TOP_OFFSET + (i - 1) * VERTICAL_SPACING
            .Width = CHECKBOX_WIDTH
            .Height = CHECKBOX_HEIGHT
            .Tag = i ' Store index for reference
        End With
        
        ' Add class instance for event handling
        Set classInst = New CheckBoxClass
        Set classInst.chk = chkBox
        m_chkSourceClasses.Add classInst
    Next i
    
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
    Set ButtonHandler = New ButtonHandler
    Set ButtonHandler.Button = btnCopySheets
End Sub
