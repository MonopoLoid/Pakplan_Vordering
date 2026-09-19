VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' WORKBOOK OPEN
' The only place this event genuinely fires automatically - VBA only wires
' Workbook_Open to the real Open event when it lives in this class module.
' =============================================================================
Private Sub Workbook_Open()
    InitiateConstants
    EnsureDataSheetDefaults
End Sub
