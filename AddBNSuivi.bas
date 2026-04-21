Option Explicit

' Syncs the BN_Suivi sheet from VHST (STR/max sprints/fonctions) and Suivi_CR (joined AA text).
Public Sub AddBNSuivi()
    Dim wsVHST As Worksheet
    Dim wsConfig As Worksheet
    Dim wsCR As Worksheet
    Dim wsBN As Worksheet
    Dim lastVHSTRow As Long
    Dim lastCRRow As Long
    Dim lastBNRow As Long
    Dim r As Long
    Dim sprintIndex As Long
    Dim maxSprint As Long
    Dim key As String
    Dim dictKey As Variant
    Dim dictCombos As Object
    Dim dictBNRows As Object
    Dim dataRow As Long
    Dim strVal As String
    Dim fonctionVal As String
    Dim sprintVal As String
    Dim joinedText As String
    Dim firstDataRowBN As Long
    Dim fonctions As Collection
    Dim fonctionItem As Variant
    Dim addedCount As Long
    Dim nomStrCol As Long
    Dim sprintsCol As Long
    Dim rowsToDelete As Collection
    Dim deleteSummary As String
    Dim deleteResp As VbMsgBoxResult

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsVHST = ThisWorkbook.Worksheets(SH_VHST)
    Set wsConfig = ThisWorkbook.Worksheets(SH_CONFIG)
    Set wsCR = ThisWorkbook.Worksheets(SH_CR)
    Set wsBN = ThisWorkbook.Worksheets(SH_BN)

    nomStrCol = FindHeaderColumn(wsVHST, 1, HDR_NOM_STR)
    If nomStrCol = 0 Then
        Err.Raise vbObjectError + 3011, "AddBNSuivi", _
                  "Colonne '" & HDR_NOM_STR & "' introuvable sur la ligne 1 de '" & wsVHST.Name & "'."
    End If

    sprintsCol = FindHeaderColumn(wsVHST, 1, HDR_SPRINTS)
    If sprintsCol = 0 Then
        Err.Raise vbObjectError + 3012, "AddBNSuivi", _
                  "Colonne '" & HDR_SPRINTS & "' introuvable sur la ligne 1 de '" & wsVHST.Name & "'."
    End If

    lastVHSTRow = wsVHST.Cells(wsVHST.Rows.Count, nomStrCol).End(xlUp).Row
    If lastVHSTRow < 2 Then lastVHSTRow = 1

    lastCRRow = wsCR.Cells(wsCR.Rows.Count, COL_B).End(xlUp).Row
    If lastCRRow < CR_FIRST_ROW Then lastCRRow = CR_FIRST_ROW - 1
    lastBNRow = wsBN.Cells(wsBN.Rows.Count, COL_B).End(xlUp).Row
    If lastBNRow < DATA_ROW_3 Then lastBNRow = DATA_ROW_2

    firstDataRowBN = BN_FIRST_DATA_ROW

    Set dictCombos = CreateObject("Scripting.Dictionary")
    dictCombos.CompareMode = vbTextCompare

    Set fonctions = New Collection
    LoadAllFonctionsFromConfig wsConfig, fonctions

    For r = 2 To lastVHSTRow
        strVal = Trim$(CStr(wsVHST.Cells(r, nomStrCol).Value))
        If strVal <> "" Then
            If IsNumeric(wsVHST.Cells(r, sprintsCol).Value) Then
                maxSprint = CLng(wsVHST.Cells(r, sprintsCol).Value)
                If maxSprint > 0 Then
                    For Each fonctionItem In fonctions
                        For sprintIndex = 1 To maxSprint
                            sprintVal = CStr(sprintIndex)
                            key = LCase$(strVal & "|" & CStr(fonctionItem) & "|" & sprintVal)
                            If Not dictCombos.Exists(key) Then
                                dictCombos.Add key, Array(strVal, CStr(fonctionItem), sprintVal)
                            End If
                        Next sprintIndex
                    Next fonctionItem
                End If
            End If
        End If
    Next r

    Set dictBNRows = CreateObject("Scripting.Dictionary")
    dictBNRows.CompareMode = vbTextCompare

    For r = firstDataRowBN To lastBNRow
        strVal = Trim$(CStr(wsBN.Cells(r, COL_B).Value))
        If strVal <> "" Then
            fonctionVal = Trim$(CStr(wsBN.Cells(r, COL_C).Value))
            sprintVal = Trim$(CStr(wsBN.Cells(r, COL_D).Value))
            key = LCase$(strVal & "|" & fonctionVal & "|" & sprintVal)
            If Not dictBNRows.Exists(key) Then
                dictBNRows.Add key, r
            End If
        End If
    Next r

    ' Detect obsolete BN rows: combos that exist in BN but no longer exist in VHST/config generated combos.
    Set rowsToDelete = BuildObsoleteBNRows(wsBN, dictCombos, dictBNRows)
    If rowsToDelete.Count > 0 Then
        deleteSummary = BuildObsoleteRowsSummary(wsBN, rowsToDelete)
        deleteResp = MsgBox( _
            "Des lignes existantes de '" & SH_BN & "' ne correspondent plus aux sprints/fonctions actuels." & vbCrLf & vbCrLf & _
            "Ces lignes seront supprimees si vous cliquez Oui." & vbCrLf & vbCrLf & _
            deleteSummary & vbCrLf & _
            "Voulez-vous supprimer ces ligne(s) maintenant ?", _
            vbYesNo + vbExclamation, "BN_Suivi - Lignes obsoletes")

        If deleteResp = vbYes Then
            DeleteRowsDescending wsBN, rowsToDelete
            lastBNRow = wsBN.Cells(wsBN.Rows.Count, COL_B).End(xlUp).Row
            If lastBNRow < DATA_ROW_3 Then lastBNRow = DATA_ROW_2
        End If
    End If

    ' Rebuild BN index after optional deletions to keep row references valid.
    Set dictBNRows = CreateObject("Scripting.Dictionary")
    dictBNRows.CompareMode = vbTextCompare
    For r = firstDataRowBN To lastBNRow
        strVal = Trim$(CStr(wsBN.Cells(r, COL_B).Value))
        If strVal <> "" Then
            fonctionVal = Trim$(CStr(wsBN.Cells(r, COL_C).Value))
            sprintVal = Trim$(CStr(wsBN.Cells(r, COL_D).Value))
            key = LCase$(strVal & "|" & fonctionVal & "|" & sprintVal)
            If Not dictBNRows.Exists(key) Then
                dictBNRows.Add key, r
            End If
        End If
    Next r

    For Each dictKey In dictCombos.Keys
        key = CStr(dictKey)
        strVal = dictCombos(key)(0)
        fonctionVal = dictCombos(key)(1)
        sprintVal = dictCombos(key)(2)

        joinedText = ComputeAAJoinedText(wsCR, lastCRRow, strVal, fonctionVal, sprintVal)

        If dictBNRows.Exists(key) Then
            dataRow = CLng(dictBNRows(key))
        Else
            lastBNRow = lastBNRow + 1
            dataRow = lastBNRow
            wsBN.Cells(dataRow, COL_B).Value = strVal
            wsBN.Cells(dataRow, COL_C).Value = fonctionVal
            wsBN.Cells(dataRow, COL_D).Value = sprintVal

            dictBNRows.Add key, dataRow
            addedCount = addedCount + 1
        End If

        wsBN.Cells(dataRow, COL_E).Value = joinedText
        ApplyBNSuiviRowBorders wsBN, dataRow
    Next dictKey

    lastBNRow = wsBN.Cells(wsBN.Rows.Count, COL_B).End(xlUp).Row
    SortBNSuiviRows wsBN, firstDataRowBN, lastBNRow

    If addedCount > 0 Then
        MsgBox "Traitement BN_Suivi termine." & vbCrLf & vbCrLf & _
               addedCount & " ligne(s) ajoutee(s) dans """ & SH_BN & """.", _
               vbInformation, "BN_Suivi"
    Else
        MsgBox "Traitement BN_Suivi termine." & vbCrLf & vbCrLf & _
               "Aucun changement : aucune nouvelle ligne n'a ete ajoutee.", _
               vbInformation, "BN_Suivi"
    End If

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Erreur lors de la mise a jour de '" & SH_BN & "' : " & Err.Description, _
           vbExclamation, "BN_Suivi"
    Resume CleanExit
End Sub

Private Function BuildObsoleteBNRows(ByVal wsBN As Worksheet, ByVal dictCombos As Object, ByVal dictBNRows As Object) As Collection
    Dim result As New Collection
    Dim k As Variant
    Dim rowNum As Long

    For Each k In dictBNRows.Keys
        If Not dictCombos.Exists(CStr(k)) Then
            rowNum = CLng(dictBNRows(CStr(k)))
            If rowNum >= BN_FIRST_DATA_ROW Then result.Add rowNum
        End If
    Next k

    Set BuildObsoleteBNRows = result
End Function

Private Function BuildObsoleteRowsSummary(ByVal wsBN As Worksheet, ByVal rowsToDelete As Collection) As String
    Dim maxPreview As Long
    Dim i As Long
    Dim rowNum As Long
    Dim preview As String
    Dim strVal As String
    Dim fonctionVal As String
    Dim sprintVal As String

    maxPreview = BN_OBSOLETE_PREVIEW_MAX
    preview = "Nombre de lignes obsoletes : " & CStr(rowsToDelete.Count) & vbCrLf

    For i = 1 To rowsToDelete.Count
        If i > maxPreview Then Exit For
        rowNum = CLng(rowsToDelete(i))
        strVal = Trim$(CStr(wsBN.Cells(rowNum, COL_B).Value & ""))
        fonctionVal = Trim$(CStr(wsBN.Cells(rowNum, COL_C).Value & ""))
        sprintVal = Trim$(CStr(wsBN.Cells(rowNum, COL_D).Value & ""))
        preview = preview & " - " & strVal & " | " & fonctionVal & " | Sprint: " & sprintVal & " (ligne " & CStr(rowNum) & ")" & vbCrLf
    Next i

    If rowsToDelete.Count > maxPreview Then
        preview = preview & " - ... et " & CStr(rowsToDelete.Count - maxPreview) & " autre(s) ligne(s)." & vbCrLf
    End If

    BuildObsoleteRowsSummary = preview
End Function

Private Sub DeleteRowsDescending(ByVal ws As Worksheet, ByVal rowsToDelete As Collection)
    Dim arr() As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    If rowsToDelete.Count = 0 Then Exit Sub

    ReDim arr(1 To rowsToDelete.Count)
    For i = 1 To rowsToDelete.Count
        arr(i) = CLng(rowsToDelete(i))
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    For i = LBound(arr) To UBound(arr)
        ws.Rows(arr(i)).Delete Shift:=xlUp
    Next i
End Sub

Private Function ComputeAAJoinedText(wsCR As Worksheet, ByVal lastCRRow As Long, _
                                     ByVal strVal As String, ByVal fonctionVal As String, _
                                     ByVal sprintVal As String) As String
    Dim r As Long
    Dim parts As Collection
    Dim eVal As String
    Dim dVal As String
    Dim aaVal As String
    Dim arr() As String
    Dim i As Long
    Dim result As String

    Set parts = New Collection

    For r = CR_FIRST_ROW To lastCRRow
        If StrComp(Trim$(CStr(wsCR.Cells(r, COL_B).Value)), strVal, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(wsCR.Cells(r, COL_D).Value)), fonctionVal, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(wsCR.Cells(r, COL_C).Value)), sprintVal, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(wsCR.Cells(r, COL_O).Value)), YES_FR, vbTextCompare) = 0 Then

            eVal = CStr(wsCR.Cells(r, COL_E).Value & "")
            If Trim$(eVal) <> "" Then
                dVal = CStr(wsCR.Cells(r, COL_D).Value & "")
                aaVal = ComputeAAValue(eVal, dVal)
                If aaVal <> "" Then
                    parts.Add aaVal
                End If
            End If
        End If
    Next r

    If parts.Count = 0 Then
        ComputeAAJoinedText = ""
        Exit Function
    End If

    ReDim arr(1 To parts.Count)
    For i = 1 To parts.Count
        arr(i) = CStr(parts(i))
    Next i

    result = Join(arr, ";" & vbLf)
    ComputeAAJoinedText = result
End Function

Private Function ComputeAAValue(ByVal eVal As String, ByVal dVal As String) As String
    Dim pos As Long
    Dim leftLen As Long

    If Trim$(eVal) = "" Then
        ComputeAAValue = ""
        Exit Function
    End If

    On Error GoTo FallbackToE

    pos = InStr(1, eVal, dVal, vbTextCompare)
    If pos <= 0 Then GoTo FallbackToE

    leftLen = pos - Len(dVal)
    If leftLen < 0 Then GoTo FallbackToE

    ComputeAAValue = Left$(eVal, leftLen)
    Exit Function

FallbackToE:
    ComputeAAValue = eVal
End Function

Private Function SplitFonctions(ByVal rawValue As String) As Collection
    Dim result As New Collection
    Dim seen As Object
    Dim normalized As String
    Dim parts() As String
    Dim item As Variant
    Dim oneFonction As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    normalized = rawValue
    normalized = Replace(normalized, vbCrLf, ";")
    normalized = Replace(normalized, vbCr, ";")
    normalized = Replace(normalized, vbLf, ";")
    normalized = Replace(normalized, ",", ";")

    parts = Split(normalized, ";")
    For Each item In parts
        oneFonction = Trim$(CStr(item))
        If oneFonction <> "" Then
            If Not seen.Exists(oneFonction) Then
                seen.Add oneFonction, True
                result.Add oneFonction
            End If
        End If
    Next item

    If result.Count = 0 Then
        oneFonction = Trim$(rawValue)
        If oneFonction <> "" Then result.Add oneFonction
    End If

    Set SplitFonctions = result
End Function

Private Sub LoadAllFonctionsFromConfig(wsConfig As Worksheet, ByRef fonctions As Collection)
    Dim seen As Object
    Dim r As Long
    Dim rawFonctions As String
    Dim oneFonctions As Collection
    Dim fonctionItem As Variant
    Dim fonctionsCol As Long
    Dim lastConfigRow As Long

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    fonctionsCol = FindHeaderColumn(wsConfig, 1, HDR_FONCTIONS)
    If fonctionsCol = 0 Then
        Err.Raise vbObjectError + 3010, "LoadAllFonctionsFromConfig", _
                  "Colonne '" & HDR_FONCTIONS & "' introuvable sur la ligne 1 de '" & wsConfig.Name & "'."
    End If

    lastConfigRow = wsConfig.Cells(wsConfig.Rows.Count, fonctionsCol).End(xlUp).Row
    For r = 2 To lastConfigRow
        rawFonctions = Trim$(CStr(wsConfig.Cells(r, fonctionsCol).Value))
        If rawFonctions <> "" Then
            Set oneFonctions = SplitFonctions(rawFonctions)
            For Each fonctionItem In oneFonctions
                If Not seen.Exists(CStr(fonctionItem)) Then
                    seen.Add CStr(fonctionItem), True
                    fonctions.Add CStr(fonctionItem)
                End If
            Next fonctionItem
        End If
    Next r
End Sub

Private Function FindHeaderColumn(ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim cellValue As String

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    For c = 1 To lastCol
        cellValue = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If StrComp(cellValue, headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Sub ApplyBNSuiviRowBorders(wsBN As Worksheet, ByVal rowNum As Long)
    Dim rng As Range
    Dim borderItem As Variant

    Set rng = wsBN.Range(wsBN.Cells(rowNum, COL_B), wsBN.Cells(rowNum, COL_G))

    For Each borderItem In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical)
        With rng.Borders(CLng(borderItem))
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next borderItem
End Sub

Private Sub SortBNSuiviRows(wsBN As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    Dim lastCol As Long
    Dim sortRange As Range

    If lastDataRow < firstDataRow Then Exit Sub

    lastCol = wsBN.Cells(2, wsBN.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_G Then lastCol = COL_G

    Set sortRange = wsBN.Range(wsBN.Cells(firstDataRow, COL_B), wsBN.Cells(lastDataRow, lastCol))

    With wsBN.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsBN.Range(wsBN.Cells(firstDataRow, COL_B), wsBN.Cells(lastDataRow, COL_B)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsBN.Range(wsBN.Cells(firstDataRow, COL_C), wsBN.Cells(lastDataRow, COL_C)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsBN.Range(wsBN.Cells(firstDataRow, COL_D), wsBN.Cells(lastDataRow, COL_D)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

