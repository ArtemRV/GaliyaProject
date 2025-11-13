VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Class module: OptionButtonClass ---
Option Explicit

Public WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1

Private Sub opt_Click()
    Dim selIndex As Long
    selIndex = CLng(opt.Tag)
    
    ' ѕерерисовать правую часть при выборе источника
    frmCopySheets.ClearRightPart
    frmCopySheets.CreateRightPart selIndex
End Sub

