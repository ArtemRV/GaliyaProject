VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Class module: (Name) = CheckBoxClass ---
Option Explicit

' Ќужен WithEvents, чтобы ловить событи€ клика по динамич. чекбоксам
Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1

Private Sub chk_Click()
    Dim selIndex As Long
    ' Tag у чекбокса вы уже заполн€ете числом Ч используем его как индекс
    selIndex = CLng(chk.Tag)
    
    ' —делаем поведение "только один выбран" Ч снимаем флажки с других chkSource_*
    Dim ctrl As MSForms.Control
    For Each ctrl In frmCopySheets.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If Left(ctrl.Name, 10) = "chkSource_" Then
                If ctrl.Name <> chk.Name Then ctrl.Value = False
            End If
        End If
    Next ctrl
    
    ' ќбновл€ем правую часть формы
    frmCopySheets.ClearRightPart
    frmCopySheets.CreateRightPart selIndex
End Sub

