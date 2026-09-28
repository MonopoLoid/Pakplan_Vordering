' =============================================================================
' WORKBOOK OPEN
' The only place this event genuinely fires automatically - VBA only wires
' Workbook_Open to the real Open event when it lives in this class module.
'
' Both calls are qualified with their module name (Module1.EnsureDataSheetDefaults
' instead of a bare call) so this keeps compiling even if another standard
' module ever ends up with an identically-named Public procedure - e.g. the
' rollback feature's Module1_Backup, which is a full copy of Module1's code
' and would otherwise make an unqualified call here ambiguous.
' =============================================================================
Private Sub Workbook_Open()
    LocalConfig.InitiateConstants
    Module1.EnsureDataSheetDefaults
End Sub
