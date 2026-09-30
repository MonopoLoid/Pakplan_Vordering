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
'
' Calls are unqualified deliberately (2026-09-30) - qualifying just
' EnsureDataSheetDefaults masked the symptom without fixing the underlying
' problem, since other ambiguity-prone spots caused errors too and this one
' simply ran first. Re-qualifying (or reintroducing Module1_Backup) needs a
' thorough check for every other name collision first, not just this one.
' =============================================================================
Private Sub Workbook_Open()
    LocalConfig.InitiateConstants
    EnsureDataSheetDefaults
End Sub
