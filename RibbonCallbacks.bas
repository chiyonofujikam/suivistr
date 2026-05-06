Option Explicit

' Save-copy ribbon action.
Public Sub Ribbon_SaveWorkbookCopy(ByVal control As Object)
    RunMacroSafe "Sauvegarde copie classeur", "SaveWorkbookCopyToSelectedDestination", "SaveModule.SaveWorkbookCopyToSelectedDestination"
End Sub

' Main update ribbon action.
Public Sub Ribbon_UpdateSuivi(ByVal control As Object)
    RunMacroSafe "Mise a jour", "UpdateSuiviLivrable", "UpdateSuiviLivrable.UpdateSuiviLivrable"
End Sub

' K-only update ribbon action.
Public Sub Ribbon_UpdateSLAvancement(ByVal control As Object)
    RunMacroSafe "MAJ Avancement", "UpdateSLAvancement", "UpdateSLAvancement.UpdateSLAvancement"
End Sub

' Archive Suivi ribbon action.
Public Sub Ribbon_ArchiveSuivi(ByVal control As Object)
    RunMacroSafe "Archivage", "ArchiveSuiviLivrable", "ArchiveSuiviLivrable.ArchiveSuiviLivrable"
End Sub

' PowQ Extract ribbon action.
Public Sub Ribbon_PowQExtract(ByVal control As Object)
    RunMacroSafe "PowQ Extract", "Update_PowQ_Exract", "PowQUpdate.Update_PowQ_Exract"
End Sub

' PowQ UVR ribbon action.
Public Sub Ribbon_PowQUVR(ByVal control As Object)
    RunMacroSafe "PowQ UVR", "Update_PowQ_Suivi_UVR", "PowQUpdate.Update_PowQ_Suivi_UVR"
End Sub

' PowQ EDU_CE_VHST ribbon action.
Public Sub Ribbon_PowQEDUCEVHST(ByVal control As Object)
    RunMacroSafe "PowQ EDU_CE_VHST", "Update_PowQ_EDU_CE_VHST", "PowQUpdate.Update_PowQ_EDU_CE_VHST"
End Sub

' PowQ all-in-one ribbon action.
Public Sub Ribbon_PowQAll(ByVal control As Object)
    RunMacroSafe "PowQ Tout", "Update_PowQ_All", "PowQUpdate.Update_PowQ_All"
End Sub

' BN_Suivi fill ribbon action.
Public Sub Ribbon_AddBNSuivi(ByVal control As Object)
    RunMacroSafe SH_BN, "AddBNSuivi", "AddBNSuivi.AddBNSuivi"
End Sub

' BN_Suivi archive ribbon action.
Public Sub Ribbon_ArchiveBNSuivi(ByVal control As Object)
    RunMacroSafe "Archivage " & SH_BN, "ArchiveBNSuivi", "ArchiveBNSuivi.ArchiveBNSuivi"
End Sub

' Collect Suivi_Livrable ribbon action.
Public Sub Ribbon_CollectSuiviLivrable(ByVal control As Object)
    RunMacroSafe "Collect Suivi_Livrable", "CollectSuiviLivrableFromSelectedFiles", "CollectSuiviLiv.CollectSuiviLivrableFromSelectedFiles"
End Sub

' Run macro by workbook name, then global name.
Private Sub RunMacroSafe(ByVal actionLabel As String, ParamArray macroNames() As Variant)
    Dim i As Long
    Dim wbQualified As String
    Dim lastErr As String
    Dim tested As String

    For i = LBound(macroNames) To UBound(macroNames)
        If i > LBound(macroNames) Then tested = tested & ", "
        tested = tested & CStr(macroNames(i))

        On Error Resume Next
        wbQualified = "'" & ThisWorkbook.Name & "'!" & CStr(macroNames(i))
        Application.Run wbQualified
        If Err.Number = 0 Then Exit Sub
        ReleaseSuiviCRLockIfOwned
        lastErr = Err.Description
        Err.Clear

        Application.Run CStr(macroNames(i))
        If Err.Number = 0 Then Exit Sub
        ReleaseSuiviCRLockIfOwned
        lastErr = Err.Description
        Err.Clear
        On Error GoTo 0
    Next i

    ReleaseSuiviCRLockIfOwned
    MsgBox "Impossible d'executer l'action '" & actionLabel & "'." & vbCrLf & _
           "Macros testees: " & tested & vbCrLf & vbCrLf & _
           "Detail: " & lastErr, vbExclamation, "Ribbon - Macro introuvable"
End Sub

