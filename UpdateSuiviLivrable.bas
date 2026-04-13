Option Explicit

' Module: UpdateSuiviLivrable
' Purpose: Synchronize `Suivi_Livrables` from `Suivi_CR` using template `Suivi_Livrables_Tmp`,
'          and compute derived columns from PowQ extracts (PowQ_Extract, PowQ_Suivi_UVR, PowQ_EDU_CE_VHST).
' Inputs:
' - Sheets referenced by constants in Globals.bas
' - Shared folder selected via SHARED_FOLDER_PATH() for LOCK.txt and status.json
Public Sub UpdateSuiviLivrable()
    Dim statusPath As String
    Dim lockCreated As Boolean
    Dim configDir As String
    Dim wsCR As Worksheet
    Dim lockValue As String
    Dim crArr As Variant
    Dim powqArr As Variant
    Dim uvrArr As Variant
    Dim vhstArr As Variant
    Dim livArr As Variant
    Dim finRefCol As Long
    Dim uvrColMap As Object
    Dim maxSprintMap As Object
    Dim isFirstRun As Boolean
    Dim oldJson As String
    Dim jsonSnapshot As String
    Dim oldSnapshot As Object
    Dim newSTRs As Collection
    Dim modifiedRows As Collection
    Dim r As Long
    Dim c As Long
    Dim strVal As String
    Dim colLetter As String
    Dim currentVal As String
    Dim oldVal As String
    Dim changed As Boolean
    Dim oldCells As Object
    Dim entry As Object
    Dim wsLiv As Worksheet
    Dim wsTmp As Worksheet
    Dim insertRow As Long
    Dim firstNewRow As Long
    Dim iterItem As Variant
    Dim bv As String, cv As String, dv As String, ev As String
    Dim rr As Long
    Dim blockEnd As Long
    Dim blockStartRow As Long
    Dim insertedCount As Long
    Dim updatedCount As Long
    Dim totalInsertedRows As Long
    Dim uniqueSTRs As Object
    Dim strsToInsert As Object
    Dim strsToUpdate As Object
    Dim strKey As Variant
    Dim matchRows As Collection
    Dim matchItem As Variant
    Dim lastTmpCol As Long
    Dim sprintMap As Object
    Dim blockInfo As Object
    Dim strSprints As Collection
    Dim maxSprintKey As String
    Dim rangesCol As Collection
    Dim tmplPair As Variant
    Dim tmplStart As Long
    Dim tmplEnd As Long
    Dim nrows As Long
    Dim sp As Variant
    Dim segIdx As Long
    Dim yp As Variant
    Dim ycol As Collection
    Dim br As Variant
    Dim i As Long
    Dim adl1Start As Long
    Dim swdsStart As Long
    Dim yellowRanges As Collection
    Dim yr As Variant
    Dim colI2 As Long
    Dim desiredSprints As Collection
    Dim existingSprints As Object
    Dim missingSprints As Collection
    Dim desiredSprintSet As Object
    Dim extraSprints As Object
    Dim delCount As Long
    Dim minRow As Long, maxRow As Long
    Dim swdsMarker As String
    Dim swdsStartRow As Long
    Dim spKey As Variant
    Dim insRow As Long
    Dim segRows As Long
    Dim pair As Variant
    Dim lastBorderCol As Long
    Dim msg As String
    Dim errNumber As Long
    Dim errMessage As String
    Dim errSource As String

    On Error GoTo ErrHandler
    lockCreated = False
    Set wsCR = ThisWorkbook.Sheets(SH_CR)

    ' ---------------------------------------------------------
    '  LOCK CHECK -- set immediately on click
    ' ---------------------------------------------------------
    If Trim$(CStr(wsCR.Range("I1").Value & "")) <> "" Then
        WaitWhileLocked wsCR, "I1"
    End If
    lockValue = "LOCKED by: " & Environ$("USERNAME") & " at " & Format$(Now, "YYYY-MM-DD HH:NN:SS")
    wsCR.Range("I1").Value = lockValue
    lockCreated = True

    wsCR.Range("I1").Locked = False
    wsCR.Protect Password:="suivi_update", UserInterfaceOnly:=False

    configDir = SHARED_FOLDER_PATH & "config\"
    If Dir$(configDir, vbDirectory) = "" Then MkDir configDir

    statusPath = configDir & "status.json"

    ' ---------------------------------------------------------
    '  VALIDATE SHEETS
    ' ---------------------------------------------------------
    ValidateRequiredSheets

    ' ---------------------------------------------------------
    '  PERFORMANCE SETTINGS
    ' ---------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Suivi Update: Loading data..."

    ' ---------------------------------------------------------
    '  LOAD SHEET DATA INTO ARRAYS
    ' ---------------------------------------------------------
    crArr = LoadSheetData(ThisWorkbook.Sheets(SH_CR))
    powqArr = LoadSheetData(ThisWorkbook.Sheets(SH_EXTRACT))
    uvrArr = LoadSheetData(ThisWorkbook.Sheets(SH_UVR))
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))
    finRefCol = FindFinRefColumn(powqArr)

    Call CheckAndOfferUpdateVHSTMaxSprints(ThisWorkbook.Sheets(SH_VHST), crArr, vhstArr)
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))

    ' ---------------------------------------------------------
    '  FIRST-RUN CHECK
    ' ---------------------------------------------------------
    isFirstRun = True
    If FileExists(statusPath) Then
        If FileLen(statusPath) > 0 Then isFirstRun = False
    End If

    If isFirstRun Then
        Application.StatusBar = "Suivi Update: Creating initial snapshot..."
        jsonSnapshot = SerializeSnapshotToJson(crArr)
        WriteTextFile statusPath, jsonSnapshot
        ' Continue in the same run: treat everything as new (empty old snapshot).
    End If

    ' ---------------------------------------------------------
    '  COMPUTE DIFF (new STR values + modified rows)
    ' ---------------------------------------------------------
    Application.StatusBar = "Suivi Update: Computing changes..."
    If isFirstRun Then
        Set oldSnapshot = CreateObject("Scripting.Dictionary")
    Else
        oldJson = ReadTextFile(statusPath)
        Set oldSnapshot = ParseSnapshotFromJson(oldJson)
    End If

    Set newSTRs = New Collection
    Set modifiedRows = New Collection

    Dim rowKey As String
    For r = CR_FIRST_ROW To UBound(crArr, 1)
        strVal = CStr(crArr(r, COL_B) & "")
        If strVal = "" Then GoTo NextCrRow

        rowKey = SnapshotRowKey(crArr, r)

        If Not oldSnapshot.Exists(rowKey) Then
            Set entry = CreateObject("Scripting.Dictionary")
            entry("STR") = strVal
            entry("row") = r
            newSTRs.Add entry
        Else
            Set oldCells = oldSnapshot(rowKey)
            changed = False
            For c = COL_B To 17
                If c > UBound(crArr, 2) Then Exit For
                colLetter = ColNumToLetter(c)
                currentVal = NormalizeValue(crArr(r, c))
                If oldCells.Exists(colLetter) Then
                    oldVal = NormalizeValue(oldCells(colLetter))
                Else
                    oldVal = ""
                End If
                If currentVal <> oldVal Then
                    changed = True
                    Exit For
                End If
            Next c
            If changed Then
                Set entry = CreateObject("Scripting.Dictionary")
                entry("STR") = strVal
                entry("row") = r
                modifiedRows.Add entry
            End If
        End If
NextCrRow:
    Next r

    ' ---------------------------------------------------------
    '  EARLY EXIT IF NO CHANGES
    ' ---------------------------------------------------------
    If newSTRs.Count = 0 And modifiedRows.Count = 0 Then
        jsonSnapshot = SerializeSnapshotToJson(crArr)
        WriteTextFile statusPath, jsonSnapshot
        MsgBox "No changes detected since last update.", vbInformation, "Suivi Update"
        GoTo Cleanup
    End If

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
    Set wsTmp = ThisWorkbook.Sheets(SH_TMP)
    livArr = LoadSheetData(wsLiv)
    lastTmpCol = wsTmp.UsedRange.Column + wsTmp.UsedRange.Columns.Count - 1
    Set sprintMap = BuildSprintRangeMap(wsTmp)
    lastBorderCol = lastTmpCol
    If lastBorderCol < COL_Y Then lastBorderCol = COL_Y

    ' ---------------------------------------------------------
    '  CLASSIFY: collect unique STRs, split into insert vs update
    ' ---------------------------------------------------------
    Set uniqueSTRs = CreateObject("Scripting.Dictionary")
    For Each iterItem In newSTRs
        Set entry = iterItem
        strVal = entry("STR")
        If Not uniqueSTRs.Exists(strVal) Then uniqueSTRs(strVal) = True
    Next iterItem
    For Each iterItem In modifiedRows
        Set entry = iterItem
        strVal = entry("STR")
        If Not uniqueSTRs.Exists(strVal) Then uniqueSTRs(strVal) = True
    Next iterItem

    Set strsToInsert = CreateObject("Scripting.Dictionary")
    Set strsToUpdate = CreateObject("Scripting.Dictionary")
    For Each strKey In uniqueSTRs.Keys
        If FindRowBySTR(livArr, CStr(strKey)) = 0 Then
            strsToInsert(strKey) = True
        Else
            strsToUpdate(strKey) = True
        End If
    Next strKey

    ' ---------------------------------------------------------
    '  INSERT: dynamic sprint blocks per STR
    ' ---------------------------------------------------------
    insertedCount = 0
    totalInsertedRows = 0
    Set yellowRanges = New Collection
    Set uvrColMap = BuildUVRColumnMap(wsLiv, uvrArr)
    Set maxSprintMap = BuildMaxSprintMapVHST(vhstArr)

    If strsToInsert.Count > 0 Then
        Application.StatusBar = "Suivi Update: Inserting " & strsToInsert.Count & " STR block(s)..."

        Set blockInfo = CreateObject("Scripting.Dictionary")

        firstNewRow = GetLastDataRow(wsLiv, COL_B) + 1
        If firstNewRow < LIV_FIRST_ROW Then firstNewRow = LIV_FIRST_ROW
        insertRow = firstNewRow

        For Each strKey In strsToInsert.Keys
            Set strSprints = GetSprintsForSTR(crArr, CStr(strKey))
            maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints, sprintMap)

            blockStartRow = insertRow

            For segIdx = 1 To 2
                If segIdx = 1 Then adl1Start = insertRow
                If segIdx = 2 Then swdsStart = insertRow

                For Each sp In strSprints
                    If sprintMap.Exists(CStr(sp)) Then
                        Set rangesCol = sprintMap(CStr(sp))
                        If segIdx <= rangesCol.Count Then
                            tmplPair = rangesCol(segIdx)
                            tmplStart = tmplPair(0)
                            tmplEnd = tmplPair(1)
                            nrows = tmplEnd - tmplStart + 1

                            wsTmp.Range(wsTmp.Cells(tmplStart, 1), wsTmp.Cells(tmplEnd, lastTmpCol)).Copy
                            wsLiv.Cells(insertRow, 1).PasteSpecial Paste:=xlPasteFormats
                            Application.CutCopyMode = False

                            wsTmp.Range(wsTmp.Cells(tmplStart, COL_C), wsTmp.Cells(tmplEnd, COL_E)).Copy
                            wsLiv.Cells(insertRow, COL_C).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False

                            For rr = insertRow To insertRow + nrows - 1
                                wsLiv.Cells(rr, COL_B).value = CStr(strKey)
                            Next rr

                            If maxSprintKey <> "" And CStr(sp) = maxSprintKey And sprintMap.Exists("3") Then
                                Set ycol = sprintMap("3")
                                If segIdx <= ycol.Count Then
                                    yp = ycol(segIdx)
                                    ApplyYellowSectionUtoX wsLiv, insertRow, insertRow + nrows - 1, wsTmp, yp(0), yp(1)
                                    yellowRanges.Add Array(insertRow, insertRow + nrows - 1)
                                End If
                            End If

                            insertRow = insertRow + nrows
                            totalInsertedRows = totalInsertedRows + nrows
                        End If
                    End If
                Next sp

                If segIdx = 1 Then
                    If insertRow > adl1Start Then
                        ApplyLightOutlineBorder wsLiv, adl1Start, insertRow - 1, lastBorderCol
                    End If
                Else
                    If insertRow > swdsStart Then
                        ApplyLightOutlineBorder wsLiv, swdsStart, insertRow - 1, lastBorderCol
                    End If
                End If
            Next segIdx

            If insertRow > blockStartRow Then
                blockEnd = insertRow - 1
                ApplyHardOutlineBorder wsLiv, blockStartRow, blockEnd, lastBorderCol
                blockInfo.Add CStr(strKey), Array(blockStartRow, blockEnd)
                insertedCount = insertedCount + 1
            End If
        Next strKey

        livArr = LoadSheetData(wsLiv)

        For Each strKey In strsToInsert.Keys
            If blockInfo.Exists(CStr(strKey)) Then
                br = blockInfo(CStr(strKey))
                For rr = br(0) To br(1)
                    bv = CStr(livArr(rr, COL_B) & "")
                    cv = CStr(livArr(rr, COL_C) & "")
                    dv = CStr(livArr(rr, COL_D) & "")
                    ev = CStr(livArr(rr, COL_E) & "")

                    wsLiv.Cells(rr, COL_F).value = ComputeColF(bv, cv, dv, ev, crArr)
                    wsLiv.Cells(rr, COL_G).value = ComputeColG(bv, cv, dv, ev, crArr)
                    wsLiv.Cells(rr, COL_H).value = ComputeColH(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_I).value = ComputeColI(bv, cv, dv, ev, powqArr, finRefCol)
                    wsLiv.Cells(rr, COL_J).value = ComputeColJ(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_M).value = ComputeColM(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_K).value = ComputeColK(bv, cv, dv, ev, crArr)

                    wsLiv.Cells(rr, COL_O).value = ComputeColO(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_T).value = ComputeColT(bv, cv, dv, ev, powqArr)

                    wsLiv.Cells(rr, COL_A).value = ComputeColA(bv, cv, dv, ev)
                Next rr
            End If
        Next strKey

        For Each yr In yellowRanges
            WriteYellowValuesUtoX wsLiv, yr(0), yr(1), uvrArr, uvrColMap, livArr
        Next yr
    End If

    ' ---------------------------------------------------------
    '  UPDATE: recompute formula columns for modified STRs
    ' ---------------------------------------------------------
    updatedCount = 0

    If strsToUpdate.Count > 0 Then
        Application.StatusBar = "Suivi Update: Recomputing " & strsToUpdate.Count & " STR(s)..."

        For Each strKey In strsToUpdate.Keys
            Set matchRows = FindAllRowsBySTR(livArr, CStr(strKey))

            ' If new sprints were added in Suivi_CR for an existing STR,
            ' insert the missing sprint segments from the template.
            Set desiredSprints = GetSprintsForSTR(crArr, CStr(strKey))
            maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, desiredSprints, sprintMap)
            Set existingSprints = CreateObject("Scripting.Dictionary")
            Set desiredSprintSet = CreateObject("Scripting.Dictionary")
            Set extraSprints = CreateObject("Scripting.Dictionary")
            minRow = 0: maxRow = 0
            For Each spKey In desiredSprints
                desiredSprintSet(CStr(spKey)) = True
            Next spKey
            For Each matchItem In matchRows
                rr = CLng(matchItem)
                If minRow = 0 Or rr < minRow Then minRow = rr
                If maxRow = 0 Or rr > maxRow Then maxRow = rr
                spKey = NormalizeSprintKey(livArr(rr, COL_D))
                If spKey <> "" Then existingSprints(CStr(spKey)) = True
            Next matchItem

            ' Delete sprint rows that no longer exist in Suivi_CR for this STR.
            If minRow > 0 Then
                For Each spKey In existingSprints.Keys
                    If Not desiredSprintSet.Exists(CStr(spKey)) Then
                        If sprintMap.Exists(CStr(spKey)) Then extraSprints(CStr(spKey)) = True
                    End If
                Next spKey

                If extraSprints.Count > 0 Then
                    delCount = 0
                    For rr = maxRow To minRow Step -1
                        spKey = NormalizeSprintKey(wsLiv.Cells(rr, COL_D).value)
                        If spKey <> "" Then
                            If extraSprints.Exists(CStr(spKey)) Then
                                wsLiv.Rows(rr).Delete
                                delCount = delCount + 1
                            End If
                        End If
                    Next rr

                    If delCount > 0 Then
                        livArr = LoadSheetData(wsLiv)
                        Set matchRows = FindAllRowsBySTR(livArr, CStr(strKey))
                        Set existingSprints = CreateObject("Scripting.Dictionary")
                        minRow = 0: maxRow = 0
                        For Each matchItem In matchRows
                            rr = CLng(matchItem)
                            If minRow = 0 Or rr < minRow Then minRow = rr
                            If maxRow = 0 Or rr > maxRow Then maxRow = rr
                            spKey = NormalizeSprintKey(livArr(rr, COL_D))
                            If spKey <> "" Then existingSprints(CStr(spKey)) = True
                        Next matchItem
                    End If
                End If
            End If

            Set missingSprints = New Collection
            For Each spKey In desiredSprints
                If Not existingSprints.Exists(CStr(spKey)) Then
                    If sprintMap.Exists(CStr(spKey)) Then missingSprints.Add CStr(spKey)
                End If
            Next spKey

            If missingSprints.Count > 0 And minRow > 0 Then
                ' Determine SwDS marker from template (col C of first SwDS row for a known sprint).
                swdsMarker = ""
                For Each spKey In desiredSprints
                    If sprintMap.Exists(CStr(spKey)) Then
                        Set rangesCol = sprintMap(CStr(spKey))
                        If rangesCol.Count >= 2 Then
                            pair = rangesCol(2)
                            swdsMarker = CStr(wsTmp.Cells(CLng(pair(0)), COL_C).value & "")
                            Exit For
                        End If
                    End If
                Next spKey

                ' Find SwDS start row inside the STR block.
                swdsStartRow = 0
                If swdsMarker <> "" Then
                    For rr = minRow To maxRow
                        If CStr(wsLiv.Cells(rr, COL_C).value & "") = swdsMarker Then
                            swdsStartRow = rr
                            Exit For
                        End If
                    Next rr
                End If
                If swdsStartRow = 0 Then swdsStartRow = maxRow + 1

                ' Clear previous yellow section colors (U-X) for this STR block.
                wsLiv.Range(wsLiv.Cells(minRow, COL_U), wsLiv.Cells(maxRow, COL_X)).Interior.ColorIndex = xlNone

                ' Insert missing ADL1 segments just before SwDS block.
                insRow = swdsStartRow
                For Each spKey In missingSprints
                    Set rangesCol = sprintMap(CStr(spKey))
                    pair = rangesCol(1) ' ADL1
                    segRows = CLng(pair(1)) - CLng(pair(0)) + 1
                    wsLiv.Rows(insRow & ":" & (insRow + segRows - 1)).Insert Shift:=xlDown

                    wsTmp.Range(wsTmp.Cells(CLng(pair(0)), 1), wsTmp.Cells(CLng(pair(1)), lastTmpCol)).Copy
                    wsLiv.Cells(insRow, 1).PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False

                    wsTmp.Range(wsTmp.Cells(CLng(pair(0)), COL_C), wsTmp.Cells(CLng(pair(1)), COL_E)).Copy
                    wsLiv.Cells(insRow, COL_C).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False

                    For rr = insRow To insRow + segRows - 1
                        wsLiv.Cells(rr, COL_B).value = CStr(strKey)
                    Next rr

                    insRow = insRow + segRows
                    maxRow = maxRow + segRows
                Next spKey

                ' Insert missing SwDS segments at the end of the STR block.
                insRow = maxRow + 1
                For Each spKey In missingSprints
                    Set rangesCol = sprintMap(CStr(spKey))
                    pair = rangesCol(2) ' SwDS
                    segRows = CLng(pair(1)) - CLng(pair(0)) + 1
                    wsLiv.Rows(insRow & ":" & (insRow + segRows - 1)).Insert Shift:=xlDown

                    wsTmp.Range(wsTmp.Cells(CLng(pair(0)), 1), wsTmp.Cells(CLng(pair(1)), lastTmpCol)).Copy
                    wsLiv.Cells(insRow, 1).PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False

                    wsTmp.Range(wsTmp.Cells(CLng(pair(0)), COL_C), wsTmp.Cells(CLng(pair(1)), COL_E)).Copy
                    wsLiv.Cells(insRow, COL_C).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False

                    For rr = insRow To insRow + segRows - 1
                        wsLiv.Cells(rr, COL_B).value = CStr(strKey)
                    Next rr

                    insRow = insRow + segRows
                    maxRow = maxRow + segRows
                Next spKey

                ' Reapply borders for the expanded STR block.
                ' ADL1: from minRow to (new) SwDS start - 1
                ' SwDS: from (new) SwDS start to maxRow
                swdsStartRow = 0
                If swdsMarker <> "" Then
                    For rr = minRow To maxRow
                        If CStr(wsLiv.Cells(rr, COL_C).value & "") = swdsMarker Then
                            swdsStartRow = rr
                            Exit For
                        End If
                    Next rr
                End If
                If swdsStartRow = 0 Then swdsStartRow = maxRow + 1
                If swdsStartRow > minRow Then ApplyLightOutlineBorder wsLiv, minRow, swdsStartRow - 1, lastBorderCol
                If swdsStartRow <= maxRow Then ApplyLightOutlineBorder wsLiv, swdsStartRow, maxRow, lastBorderCol
                ApplyHardOutlineBorder wsLiv, minRow, maxRow, lastBorderCol

                ' Recompute yellow sprint and reapply yellow background + UVR values on that sprint rows only.
                livArr = LoadSheetData(wsLiv)
                Set matchRows = FindAllRowsBySTR(livArr, CStr(strKey))
                Set desiredSprints = GetSprintsForSTR(crArr, CStr(strKey))
                maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, desiredSprints, sprintMap)
                If maxSprintKey <> "" And sprintMap.Exists("3") Then
                    Set ycol = sprintMap("3")
                    For segIdx = 1 To 2
                        If segIdx <= ycol.Count Then
                            yp = ycol(segIdx)
                            For Each matchItem In matchRows
                                rr = CLng(matchItem)
                                If NormalizeSprintKey(livArr(rr, COL_D)) = maxSprintKey Then
                                    ApplyYellowSectionUtoX wsLiv, rr, rr, wsTmp, CLng(yp(0)), CLng(yp(0))
                                End If
                            Next matchItem
                        End If
                    Next segIdx
                End If
            End If

            For Each matchItem In matchRows
                rr = CLng(matchItem)

                bv = CStr(livArr(rr, COL_B) & "")
                cv = CStr(livArr(rr, COL_C) & "")
                dv = CStr(livArr(rr, COL_D) & "")
                ev = CStr(livArr(rr, COL_E) & "")

                wsLiv.Cells(rr, COL_F).value = ComputeColF(bv, cv, dv, ev, crArr)
                wsLiv.Cells(rr, COL_G).value = ComputeColG(bv, cv, dv, ev, crArr)
                wsLiv.Cells(rr, COL_H).value = ComputeColH(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_I).value = ComputeColI(bv, cv, dv, ev, powqArr, finRefCol)
                wsLiv.Cells(rr, COL_J).value = ComputeColJ(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_M).value = ComputeColM(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_K).value = ComputeColK(bv, cv, dv, ev, crArr)

                wsLiv.Cells(rr, COL_O).value = ComputeColO(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_T).value = ComputeColT(bv, cv, dv, ev, powqArr)

                If maxSprintKey <> "" And NormalizeSprintKey(dv) = maxSprintKey Then
                    For colI2 = COL_U To COL_X
                        If uvrColMap.Exists(colI2) Then
                            wsLiv.Cells(rr, colI2).value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colI2)))
                        End If
                    Next colI2
                End If

                wsLiv.Cells(rr, COL_A).value = ComputeColA(bv, cv, dv, ev)
            Next matchItem
            updatedCount = updatedCount + 1
        Next strKey
    End If

    ' Borders can get impacted by row insertions; rebuild from scratch.
    RebuildSuiviLivrablesBorders wsLiv, wsTmp, sprintMap, lastBorderCol

    ' ---------------------------------------------------------
    '  SAVE NEW SNAPSHOT
    ' ---------------------------------------------------------
    Application.StatusBar = "Suivi Update: Saving snapshot..."
    jsonSnapshot = SerializeSnapshotToJson(crArr)
    WriteTextFile statusPath, jsonSnapshot

    msg = "Update completed successfully." & vbCrLf & vbCrLf & _
          "Actions on Suivi_Livrables:" & vbCrLf & _
          "  - " & insertedCount & " new STR blocks inserted (" & totalInsertedRows & " rows)" & vbCrLf & _
          "  - " & updatedCount & " existing STRs recomputed"
    MsgBox msg, vbInformation, "Suivi Update"
    GoTo Cleanup

ErrHandler:
    errNumber = Err.Number
    errMessage = Err.Description
    errSource = Err.Source

    On Error Resume Next
    AppendTextFile configDir & "error_logs.txt", _
        Format$(Now, "YYYY-MM-DD HH:NN:SS") & _
        " | user=" & Environ$("USERNAME") & _
        " | err=" & errNumber & _
        " | src=" & errSource & _
        " | " & errMessage

    MsgBox "Update failed: " & errMessage & _
           " (Error " & errNumber & ")", vbCritical, "Suivi Update"
    Resume Cleanup

Cleanup:
    On Error Resume Next
    If lockCreated Then
        wsCR.Unprotect Password:="suivi_update"
        wsCR.Range("I1").ClearContents
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


