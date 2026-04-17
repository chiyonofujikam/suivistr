Option Explicit

' Ribbon callback for save-copy button.
Public Sub Ribbon_SaveWorkbookCopy(ByVal control As Object)
    RunMacroSafe "Sauvegarde copie classeur", "SaveWorkbookCopyToSelectedDestination", "SaveModule.SaveWorkbookCopyToSelectedDestination"
End Sub

' Ribbon callback for main update button.
Public Sub Ribbon_UpdateSuivi(ByVal control As Object)
    RunMacroSafe "Mise a jour", "UpdateSuiviLivrable", "UpdateSuiviLivrable.UpdateSuiviLivrable"
End Sub

' Ribbon callback for archive button.
Public Sub Ribbon_ArchiveSuivi(ByVal control As Object)
    RunMacroSafe "Archivage", "ArchiveSuiviLivrable", "ArchiveSuiviLivrable.ArchiveSuiviLivrable"
End Sub

' Ribbon callback for PowQ extract button.
Public Sub Ribbon_PowQExtract(ByVal control As Object)
    RunMacroSafe "PowQ Extract", "Update_PowQ_Exract", "PowQUpdate.Update_PowQ_Exract"
End Sub

' Ribbon callback for PowQ UVR button.
Public Sub Ribbon_PowQUVR(ByVal control As Object)
    RunMacroSafe "PowQ UVR", "Update_PowQ_Suivi_UVR", "PowQUpdate.Update_PowQ_Suivi_UVR"
End Sub

' Ribbon callback for PowQ EDU_CE_VHST button.
Public Sub Ribbon_PowQEDUCEVHST(ByVal control As Object)
    MsgBox "Cette fonctionnalite est en cours de developpement.", vbExclamation, "En cours de developpement"
End Sub

' Ribbon callback for PowQ 'all' button (runs all PowQ actions - en cours de dev).
Public Sub Ribbon_PowQAll(ByVal control As Object)
    MsgBox "Cette fonctionnalite est en cours de developpement.", vbExclamation, "En cours de developpement"
End Sub

' Ribbon callback for BN_Suivi button.
Public Sub Ribbon_AddBNSuivi(ByVal control As Object)
    RunMacroSafe "BN_Suivi dossier Safety", "AddBNSuivi", "AddBNSuivi.AddBNSuivi"
End Sub

' Ribbon callback for BN_Suivi archive button.
Public Sub Ribbon_ArchiveBNSuivi(ByVal control As Object)
    RunMacroSafe "Archivage BN_Suivi dossier Safety", "ArchiveBNSuivi", "ArchiveBNSuivi.ArchiveBNSuivi"
End Sub

' Runs macro names safely using workbook-qualified then global lookup.
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
        lastErr = Err.Description
        Err.Clear

        Application.Run CStr(macroNames(i))
        If Err.Number = 0 Then Exit Sub
        lastErr = Err.Description
        Err.Clear
        On Error GoTo 0
    Next i

    MsgBox "Impossible d'executer l'action '" & actionLabel & "'." & vbCrLf & _
           "Macros testees: " & tested & vbCrLf & vbCrLf & _
           "Detail: " & lastErr, vbExclamation, "Ribbon - Macro introuvable"
End Sub

