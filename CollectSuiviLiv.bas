Option Explicit

Public Sub CollectSuiviLivrableFromSelectedFiles()
    Dim dlg As Object
    Dim selectedPath As Variant
    Dim wbOutput As Workbook
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsCopied As Worksheet
    Dim savePath As String
    Dim saveFolder As String
    Dim folderPicker As Object
    Dim createdCount As Long
    Dim statusReport As String
    Dim totalSelected As Long
    Dim missingCount As Long
    Dim failedCount As Long
    Dim copiedCount As Long
    Dim shortName As String
    Dim wasAlreadyOpen As Boolean
    Dim defaultSheetName As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    Dim initiallyOpenWorkbooks As Collection
    Dim wb As Workbook
    Dim openSavedFileResp As VbMsgBoxResult
    Dim outputWindow As Window
    Dim wsGlobal As Worksheet
    Dim copiedRows As Long
    Dim globalRows As Long
    Dim sourceHeaders As Variant
    Dim baselineHeaderSignature As String
    Dim headerMismatchCount As Long
    Dim currentHeaderSignature As String
    Dim baseSheetName As String

    On Error GoTo ErrHandler

    ' Snapshot open workbooks to avoid closing user-open files.
    Set initiallyOpenWorkbooks = New Collection
    For Each wb In Application.Workbooks
        initiallyOpenWorkbooks.Add wb
    Next wb

    Set dlg = Application.FileDialog(3)
    With dlg
        .Title = "Selectionner les fichiers Excel a collecter"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xls;*.xlsx;*.xlsm", 1
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count = 0 Then Exit Sub
    End With
    totalSelected = dlg.SelectedItems.Count

    Set wbOutput = Workbooks.Add(xlWBATWorksheet)
    defaultSheetName = wbOutput.Worksheets(1).Name
    Set wsGlobal = wbOutput.Worksheets(1)
    wsGlobal.Name = "Global"
    wsGlobal.Cells(1, 1).Value = "Pole"
    On Error Resume Next
    Set outputWindow = wbOutput.Windows(1)
    If Not outputWindow Is Nothing Then outputWindow.Visible = False
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each selectedPath In dlg.SelectedItems
        shortName = BaseFileNameWithoutExtension(CStr(selectedPath))
        sourceSheetName = SH_LIV
        Set wsSource = Nothing
        Set wbSource = Nothing
        wasAlreadyOpen = False

        If StrComp(CStr(selectedPath), ThisWorkbook.FullName, vbTextCompare) = 0 Then
            Set wbSource = ThisWorkbook
            wasAlreadyOpen = True
        Else
            Set wbSource = GetOpenWorkbookByPath(CStr(selectedPath))
            If Not wbSource Is Nothing Then
                wasAlreadyOpen = True
            Else
                On Error Resume Next
                Set wbSource = Workbooks.Open(CStr(selectedPath), ReadOnly:=True, UpdateLinks:=0)
                If Err.Number <> 0 Or wbSource Is Nothing Then
                    failedCount = failedCount + 1
                    If Len(statusReport) > 0 Then statusReport = statusReport & vbCrLf
                    statusReport = statusReport & "- " & shortName & " : ECHEC (ouverture impossible)"
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo NextFile
                End If
                On Error GoTo ErrHandler
            End If
        End If
        wasAlreadyOpen = WorkbookWasInitiallyOpen(wbSource, initiallyOpenWorkbooks)


        On Error Resume Next
        Set wsSource = wbSource.Worksheets(sourceSheetName)
        On Error GoTo ErrHandler

        If wsSource Is Nothing Then
            missingCount = missingCount + 1
            If Len(statusReport) > 0 Then statusReport = statusReport & vbCrLf
            statusReport = statusReport & "- " & shortName & " : IGNORE (feuille '" & sourceSheetName & "' introuvable)"
            If Not wasAlreadyOpen Then wbSource.Close SaveChanges:=False
            Set wbSource = Nothing
            GoTo NextFile
        End If

        sourceHeaders = ReadSuiviLivHeaders(wsSource)
        If wsGlobal.Cells(1, 2).Value2 = "" Then
            wsGlobal.Range("B1:Y1").Value2 = sourceHeaders
        End If
        currentHeaderSignature = BuildHeaderSignature(sourceHeaders)
        If Len(baselineHeaderSignature) = 0 Then
            baselineHeaderSignature = currentHeaderSignature
        ElseIf StrComp(baselineHeaderSignature, currentHeaderSignature, vbTextCompare) <> 0 Then
            headerMismatchCount = headerMismatchCount + 1
            If Len(statusReport) > 0 Then statusReport = statusReport & vbCrLf
            statusReport = statusReport & "- " & shortName & " : ATTENTION (en-tetes differents)"
        End If

        baseSheetName = ResolveCollectedSheetBaseName(wbSource, CStr(selectedPath))
        targetSheetName = BuildUniqueSheetName(wbOutput, baseSheetName)
        Set wsCopied = CopySuiviLivDataToWorkbook(wsSource, wbOutput, targetSheetName, sourceHeaders, copiedRows)
        If wsCopied Is Nothing Then
            failedCount = failedCount + 1
            If Len(statusReport) > 0 Then statusReport = statusReport & vbCrLf
            statusReport = statusReport & "- " & shortName & " : ECHEC (copie de la feuille impossible)"
            If Not wasAlreadyOpen Then wbSource.Close SaveChanges:=False
            Set wbSource = Nothing
            GoTo NextFile
        End If
        EnsureWorksheetTable wsCopied
        globalRows = globalRows + AppendSuiviLivDataToGlobal(wsSource, wsGlobal, targetSheetName)
        createdCount = createdCount + 1
        copiedCount = copiedCount + 1
        If Len(statusReport) > 0 Then statusReport = statusReport & vbCrLf
        statusReport = statusReport & "- " & shortName & " : OK (" & CStr(copiedRows) & " lignes)"

        If Not wasAlreadyOpen Then wbSource.Close SaveChanges:=False
        Set wbSource = Nothing

NextFile:
        Set wsSource = Nothing
        Set wsCopied = Nothing
    Next selectedPath

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If createdCount = 0 Then
        wbOutput.Close SaveChanges:=False
        MsgBox "Aucune feuille '" & SH_LIV & "' n'a ete trouvee dans les fichiers selectionnes." & vbCrLf & vbCrLf & _
               "Statut :" & vbCrLf & statusReport, vbExclamation, "Collect Suivi_Livrable"
        Exit Sub
    End If

    EnsureWorksheetTable wsGlobal

    Set folderPicker = Application.FileDialog(4)
    With folderPicker
        .Title = "Selectionner le dossier de sauvegarde"
        If .Show <> -1 Then
            wbOutput.Close SaveChanges:=False
            Set wbOutput = Nothing
            MsgBox "Collecte annulee : aucun fichier n'a ete enregistre.", vbInformation, "Collect Suivi_Livrable"
            MsgBox "Statut :" & vbCrLf & statusReport, vbInformation, "Collect Suivi_Livrable"
            Exit Sub
        End If
        saveFolder = CStr(.SelectedItems(1))
    End With

    If Right$(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\"
    savePath = saveFolder & "Collect_" & Format$(Now, "hhnnss_ddmmyyyy") & ".xlsx"

    ' Save first, then ask whether to show the generated workbook.
    wbOutput.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    openSavedFileResp = MsgBox("Collecte terminee :" & vbCrLf & _
                               "- Fichiers selectionnes : " & CStr(totalSelected) & vbCrLf & _
                               "- Copies OK : " & CStr(copiedCount) & vbCrLf & _
                               "- Lignes dans 'Global' : " & CStr(globalRows) & vbCrLf & _
                               "- Ecarts d'en-tetes : " & CStr(headerMismatchCount) & vbCrLf & _
                               "- Ignores (sans '" & SH_LIV & "') : " & CStr(missingCount) & vbCrLf & _
                               "- Echecs : " & CStr(failedCount) & vbCrLf & vbCrLf & _
                               "Fichier genere :" & vbCrLf & savePath & vbCrLf & vbCrLf & _
                               "Statut detaille :" & vbCrLf & statusReport & vbCrLf & vbCrLf & _
                               "Voulez-vous l'ouvrir maintenant ?", _
                               vbYesNo + vbQuestion, "Collect Suivi_Livrable")
    If openSavedFileResp = vbYes Then
        On Error Resume Next
        Set outputWindow = wbOutput.Windows(1)
        If Not outputWindow Is Nothing Then outputWindow.Visible = True
        On Error GoTo ErrHandler
        wbOutput.Activate
    Else
        wbOutput.Close SaveChanges:=False
    End If
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not wbSource Is Nothing Then
        If Not (wbSource Is ThisWorkbook) Then
            On Error Resume Next
            wbSource.Close SaveChanges:=False
            On Error GoTo 0
        End If
    End If
    If Not wbOutput Is Nothing Then
        On Error Resume Next
        wbOutput.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Erreur pendant la collecte : " & Err.Description, vbCritical, "Collect Suivi_Livrable"
End Sub

Private Function GetOpenWorkbookByPath(ByVal targetPath As String) As Workbook
    Dim wb As Workbook
    Dim normalizedTarget As String
    Dim normalizedWbPath As String

    normalizedTarget = LCase$(Replace(targetPath, "/", "\"))
    For Each wb In Application.Workbooks
        normalizedWbPath = LCase$(Replace(CStr(wb.FullName), "/", "\"))
        If StrComp(normalizedWbPath, normalizedTarget, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
End Function

Private Function CopySuiviLivDataToWorkbook(ByVal wsSource As Worksheet, ByVal wbTarget As Workbook, _
                                            ByVal newSheetName As String, ByVal headers As Variant, _
                                            ByRef copiedRowCount As Long) As Worksheet
    Dim wsNew As Worksheet
    Dim srcRange As Range
    Dim lastRow As Long
    Dim headerRange As Range

    On Error GoTo CopyErr

    Set wsNew = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
    wsNew.Name = newSheetName

    Set headerRange = wsSource.Range("B3:Y3")
    wsNew.Range("A1").Resize(1, headerRange.Columns.Count).Value2 = headerRange.Value2

    lastRow = GetLastSuiviLivRow(wsSource)
    copiedRowCount = 0
    If lastRow >= 4 Then
        Set srcRange = wsSource.Range("B4:Y" & CStr(lastRow))
        wsNew.Range("A2").Resize(srcRange.Rows.Count, srcRange.Columns.Count).Value2 = srcRange.Value2
        copiedRowCount = srcRange.Rows.Count
    End If

    Set CopySuiviLivDataToWorkbook = wsNew
    Exit Function

CopyErr:
    Application.CutCopyMode = False
    Set CopySuiviLivDataToWorkbook = Nothing
End Function

Private Function AppendSuiviLivDataToGlobal(ByVal wsSource As Worksheet, ByVal wsGlobal As Worksheet, _
                                            ByVal sourceSheetName As String) As Long
    Dim lastRow As Long
    Dim srcRange As Range
    Dim srcData As Variant
    Dim startRow As Long
    Dim r As Long
    Dim rowCount As Long

    lastRow = GetLastSuiviLivRow(wsSource)
    If lastRow < 4 Then
        AppendSuiviLivDataToGlobal = 0
        Exit Function
    End If

    Set srcRange = wsSource.Range("B4:Y" & CStr(lastRow))
    srcData = srcRange.Value2
    rowCount = srcRange.Rows.Count

    startRow = wsGlobal.Cells(wsGlobal.Rows.Count, 1).End(xlUp).Row

    For r = 1 To rowCount
        wsGlobal.Cells(startRow + r, 1).Value2 = sourceSheetName
        wsGlobal.Range(wsGlobal.Cells(startRow + r, 2), wsGlobal.Cells(startRow + r, 25)).Value2 = _
            Application.Index(srcData, r, 0)
    Next r

    AppendSuiviLivDataToGlobal = rowCount
End Function

Private Function GetLastSuiviLivRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row ' Column B
    If lastRow < 3 Then
        GetLastSuiviLivRow = 0
    Else
        GetLastSuiviLivRow = lastRow
    End If
End Function

Private Function ReadSuiviLivHeaders(ByVal ws As Worksheet) As Variant
    ReadSuiviLivHeaders = ws.Range("B3:Y3").Value2
End Function

Private Function BuildHeaderSignature(ByVal headers As Variant) As String
    Dim c As Long
    Dim part As String
    Dim result As String

    For c = 1 To UBound(headers, 2)
        part = Trim$(CStr(headers(1, c)))
        If Len(part) > 0 Then
            If Len(result) > 0 Then result = result & "|"
            result = result & LCase$(part)
        End If
    Next c
    BuildHeaderSignature = result
End Function


Private Function WorkbookWasInitiallyOpen(ByVal targetWb As Workbook, ByVal initiallyOpen As Collection) As Boolean
    Dim wb As Workbook
    If targetWb Is Nothing Then Exit Function
    For Each wb In initiallyOpen
        If wb Is targetWb Then
            WorkbookWasInitiallyOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function BaseFileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    fileName = Dir$(filePath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        BaseFileNameWithoutExtension = Left$(fileName, dotPos - 1)
    Else
        BaseFileNameWithoutExtension = fileName
    End If
End Function

Private Function ResolveCollectedSheetBaseName(ByVal wbSource As Workbook, ByVal sourcePath As String) As String
    Dim wsConfig As Worksheet
    Dim candidate As String

    On Error Resume Next
    Set wsConfig = wbSource.Worksheets(SH_CONFIG)
    On Error GoTo 0

    If Not wsConfig Is Nothing Then
        candidate = Trim$(CStr(wsConfig.Range("A2").Value2 & ""))
        If Len(candidate) > 0 Then
            ResolveCollectedSheetBaseName = candidate
            Exit Function
        End If
    End If

    ResolveCollectedSheetBaseName = BaseFileNameWithoutExtension(sourcePath)
End Function

Private Function BuildUniqueSheetName(ByVal wb As Workbook, ByVal rawName As String) As String
    Dim cleaned As String
    Dim candidate As String
    Dim idx As Long
    Dim suffix As String

    cleaned = CleanSheetName(rawName)
    If Len(cleaned) = 0 Then cleaned = "Suivi_Livrable"

    candidate = Left$(cleaned, 31)
    If Not SheetExistsInWorkbook(wb, candidate) Then
        BuildUniqueSheetName = candidate
        Exit Function
    End If

    idx = 1
    Do
        suffix = "_" & CStr(idx)
        candidate = Left$(cleaned, 31 - Len(suffix)) & suffix
        idx = idx + 1
    Loop While SheetExistsInWorkbook(wb, candidate)

    BuildUniqueSheetName = candidate
End Function

Private Function CleanSheetName(ByVal sheetName As String) As String
    Dim result As String
    result = sheetName
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Trim$(result)
    CleanSheetName = result
End Function

Private Function SheetExistsInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExistsInWorkbook = Not ws Is Nothing
End Function

Private Sub EnsureWorksheetTable(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim lo As ListObject
    Dim tableName As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop

    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tableName = BuildUniqueTableName(ws.Parent, ws.Name)
    lo.Name = tableName
End Sub

Private Function BuildUniqueTableName(ByVal wb As Workbook, ByVal rawName As String) As String
    Dim baseName As String
    Dim candidate As String
    Dim idx As Long

    baseName = CleanTableName(rawName)
    If Len(baseName) = 0 Then baseName = "Table1"

    candidate = baseName
    idx = 1
    Do While TableNameExistsInWorkbook(wb, candidate)
        candidate = Left$(baseName, 255 - Len("_" & CStr(idx))) & "_" & CStr(idx)
        idx = idx + 1
    Loop

    BuildUniqueTableName = candidate
End Function

Private Function CleanTableName(ByVal rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    For i = 1 To Len(rawName)
        ch = Mid$(rawName, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i

    result = Replace(result, " ", "_")
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop

    If Len(result) = 0 Then
        result = "Table1"
    ElseIf Not ((Left$(result, 1) >= "A" And Left$(result, 1) <= "Z") Or _
                (Left$(result, 1) >= "a" And Left$(result, 1) <= "z") Or _
                Left$(result, 1) = "_") Then
        result = "_" & result
    End If

    If Len(result) > 255 Then result = Left$(result, 255)
    CleanTableName = result
End Function

Private Function TableNameExistsInWorkbook(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                TableNameExistsInWorkbook = True
                Exit Function
            End If
        Next lo
    Next ws
End Function
