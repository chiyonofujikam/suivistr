Option Explicit

' Saves current workbook, then saves a copy to a selected folder.
Public Sub SaveWorkbookCopyToSelectedDestination()
    Dim wsVHST As Worksheet
    Dim folderPath As String
    Dim poleValue As String
    Dim poleLabel As String
    Dim poleCol As Long
    Dim fileName As String
    Dim targetPath As String
    Dim dlg As Object
    Dim confirmMsg As String
    Dim resp As VbMsgBoxResult

    On Error GoTo ErrHandler

    Set dlg = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With dlg
        .Title = "Selectionner le dossier de destination pour la copie"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = CStr(.SelectedItems(1))
    End With

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    poleValue = ""
    On Error Resume Next
    Set wsVHST = ThisWorkbook.Worksheets(SH_VHST)
    On Error GoTo ErrHandler
    If Not wsVHST Is Nothing Then
        poleCol = FindColumnByHeader(wsVHST, "Pole")
        If poleCol > 0 Then
            poleValue = Trim$(CStr(wsVHST.Cells(2, poleCol).Value & ""))
        End If
    End If
    poleLabel = poleValue
    poleValue = SanitizeFileNamePart(poleValue)

    If Trim$(poleLabel) <> "" Then
        confirmMsg = "Cette action va sauvegarder un fichier Suivi STR pour le Pole : " & poleLabel & "." & vbCrLf & vbCrLf & _
                     "Voulez-vous continuer ?"
    Else
        confirmMsg = "Le nom du Pole n'est pas renseigne (colonne 'Pole', ligne 2 de " & SH_VHST & ")." & vbCrLf & vbCrLf & _
                     "Cette action va sauvegarder un fichier Suivi STR sans nom de Pole." & vbCrLf & vbCrLf & _
                     "Voulez-vous continuer ?"
    End If
    resp = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirmation sauvegarde")
    If resp <> vbYes Then Exit Sub

    fileName = "Suivi_Pole_" & poleValue & ".xlsm"
    targetPath = folderPath & fileName

    ThisWorkbook.Save
    ThisWorkbook.SaveCopyAs targetPath
    ClearFunctionsInCopiedWorkbook targetPath

    MsgBox "Copie enregistree :" & vbCrLf & targetPath & vbCrLf & vbCrLf & _
           "Veuillez renseigner manuellement la liste des fonctions dans la colonne 'Fonctions' de " & SH_VHST & " du fichier copie.", _
           vbInformation, "Sauvegarde terminee"
    Exit Sub

ErrHandler:
    MsgBox "Echec de la sauvegarde de la copie : " & Err.Description, vbCritical, "Sauvegarde"
End Sub

' Finds a column index by header value in row 1.
Private Function FindColumnByHeader(ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim currentHeader As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        currentHeader = Trim$(CStr(ws.Cells(1, c).Value & ""))
        If StrComp(currentHeader, headerName, vbTextCompare) = 0 Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
    FindColumnByHeader = 0
End Function

' Clears functions in "Fonctions" column in copied workbook.
Private Sub ClearFunctionsInCopiedWorkbook(ByVal copyPath As String)
    Dim wbCopy As Workbook
    Dim wsCopy As Worksheet
    Dim fonctionsCol As Long
    Dim lastRow As Long

    Set wbCopy = Workbooks.Open(Filename:=copyPath, UpdateLinks:=0, ReadOnly:=False)
    Set wsCopy = wbCopy.Worksheets(SH_VHST)
    fonctionsCol = FindColumnByHeader(wsCopy, "Fonctions")

    If fonctionsCol > 0 Then
        lastRow = wsCopy.Cells(wsCopy.Rows.Count, fonctionsCol).End(xlUp).Row
        If lastRow < 2 Then lastRow = 2
        wsCopy.Range(wsCopy.Cells(2, fonctionsCol), wsCopy.Cells(lastRow, fonctionsCol)).ClearContents
    End If

    wbCopy.Save
    wbCopy.Close SaveChanges:=False
End Sub

' Removes filesystem-invalid characters from filename fragment.
Private Function SanitizeFileNamePart(ByVal s As String) As String
    Dim badChars As Variant
    Dim ch As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, CStr(ch), "_")
    Next ch

    SanitizeFileNamePart = s
End Function
