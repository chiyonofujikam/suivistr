Option Explicit

Public Sub UpdateSuiviLivrable()
    Dim lockPath As String
    Dim statusPath As String
    Dim lockCreated As Boolean
    Dim crArr As Variant
    Dim powqArr As Variant
    Dim livArr As Variant
    Dim finRefCol As Long
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
    Dim bvK As String, cvK As String, dvK As String, evK As String
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
    Dim msg As String

    On Error GoTo ErrHandler

    lockPath = SHARED_FOLDER_PATH & "LOCK.txt"
    statusPath = SHARED_FOLDER_PATH & "status.json"
    lockCreated = False

    ' ---------------------------------------------------------
    '  LOCK CHECK
    ' ---------------------------------------------------------
    If FileExists(lockPath) Then
        MsgBox "An update is already in progress by another user." & vbCrLf & _
               "Please wait and try again.", vbExclamation, "Suivi Update"
        Exit Sub
    End If
    WriteTextFile lockPath, "LOCKED by: " & Environ("USERNAME") & " at " & Now()
    lockCreated = True

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
    finRefCol = FindFinRefColumn(powqArr)

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
        MsgBox "Initial snapshot created. The sheet is now tracked." & vbCrLf & _
               "Run Update again to perform the first synchronization.", _
               vbInformation, "Suivi Update"
        GoTo Cleanup
    End If

    ' ---------------------------------------------------------
    '  COMPUTE DIFF (new STR values + modified rows)
    ' ---------------------------------------------------------
    Application.StatusBar = "Suivi Update: Computing changes..."
    oldJson = ReadTextFile(statusPath)
    Set oldSnapshot = ParseSnapshotFromJson(oldJson)

    Set newSTRs = New Collection
    Set modifiedRows = New Collection

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        strVal = CStr(crArr(r, COL_B) & "")
        If strVal = "" Then GoTo NextCrRow

        If Not oldSnapshot.Exists(strVal) Then
            Set entry = CreateObject("Scripting.Dictionary")
            entry("STR") = strVal
            entry("row") = r
            newSTRs.Add entry
        Else
            Set oldCells = oldSnapshot(strVal)
            changed = False
            For c = 1 To UBound(crArr, 2)
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

    If strsToInsert.Count > 0 Then
        Application.StatusBar = "Suivi Update: Inserting " & strsToInsert.Count & " STR block(s)..."

        lastTmpCol = wsTmp.UsedRange.Column + wsTmp.UsedRange.Columns.Count - 1
        Set sprintMap = BuildSprintRangeMap(wsTmp)
        Set blockInfo = CreateObject("Scripting.Dictionary")

        firstNewRow = GetLastDataRow(wsLiv, COL_B) + 1
        If firstNewRow < LIV_FIRST_ROW Then firstNewRow = LIV_FIRST_ROW
        insertRow = firstNewRow

        For Each strKey In strsToInsert.Keys
            Set strSprints = GetSprintsForSTR(crArr, CStr(strKey))
            maxSprintKey = ""
            For i = strSprints.Count To 1 Step -1
                If sprintMap.Exists(CStr(strSprints(i))) Then
                    maxSprintKey = CStr(strSprints(i))
                    Exit For
                End If
            Next i

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
                                wsLiv.Cells(rr, COL_B).Value = CStr(strKey)
                            Next rr

                            If maxSprintKey <> "" And CStr(sp) = maxSprintKey And sprintMap.Exists("3") Then
                                Set ycol = sprintMap("3")
                                If segIdx <= ycol.Count Then
                                    yp = ycol(segIdx)
                                    ApplyYellowSectionUtoX wsLiv, insertRow, insertRow + nrows - 1, wsTmp, yp(0), yp(1)
                                End If
                            End If

                            insertRow = insertRow + nrows
                            totalInsertedRows = totalInsertedRows + nrows
                        End If
                    End If
                Next sp

                If segIdx = 1 Then
                    If insertRow > adl1Start Then
                        ApplyLightOutlineBorder wsLiv, adl1Start, insertRow - 1, lastTmpCol
                    End If
                Else
                    If insertRow > swdsStart Then
                        ApplyLightOutlineBorder wsLiv, swdsStart, insertRow - 1, lastTmpCol
                    End If
                End If
            Next segIdx

            If insertRow > blockStartRow Then
                blockEnd = insertRow - 1
                ApplyHardOutlineBorder wsLiv, blockStartRow, blockEnd, lastTmpCol
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

                    wsLiv.Cells(rr, COL_F).Value = ComputeColF(bv, cv, dv, ev, crArr)
                    wsLiv.Cells(rr, COL_G).Value = ComputeColG(bv, cv, dv, ev, crArr)
                    wsLiv.Cells(rr, COL_H).Value = ComputeColH(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_I).Value = ComputeColI(bv, cv, dv, ev, powqArr, finRefCol)
                    wsLiv.Cells(rr, COL_J).Value = ComputeColJ(bv, cv, dv, ev, powqArr)

                    If rr + 1 <= UBound(livArr, 1) Then
                        bvK = CStr(livArr(rr + 1, COL_B) & "")
                        cvK = CStr(livArr(rr + 1, COL_C) & "")
                        dvK = CStr(livArr(rr + 1, COL_D) & "")
                        evK = CStr(livArr(rr + 1, COL_E) & "")
                    Else
                        bvK = "": cvK = "": dvK = "": evK = ""
                    End If
                    wsLiv.Cells(rr, COL_K).Value = ComputeColK(bvK, cvK, dvK, evK, crArr)

                    wsLiv.Cells(rr, COL_O).Value = ComputeColO(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_T).Value = ComputeColT(bv, cv, dv, ev, powqArr)
                Next rr
            End If
        Next strKey
    End If

    ' ---------------------------------------------------------
    '  UPDATE: recompute formula columns for modified STRs
    ' ---------------------------------------------------------
    updatedCount = 0

    If strsToUpdate.Count > 0 Then
        Application.StatusBar = "Suivi Update: Recomputing " & strsToUpdate.Count & " STR(s)..."

        For Each strKey In strsToUpdate.Keys
            Set matchRows = FindAllRowsBySTR(livArr, CStr(strKey))

            For Each matchItem In matchRows
                rr = CLng(matchItem)

                bv = CStr(livArr(rr, COL_B) & "")
                cv = CStr(livArr(rr, COL_C) & "")
                dv = CStr(livArr(rr, COL_D) & "")
                ev = CStr(livArr(rr, COL_E) & "")

                wsLiv.Cells(rr, COL_F).Value = ComputeColF(bv, cv, dv, ev, crArr)
                wsLiv.Cells(rr, COL_G).Value = ComputeColG(bv, cv, dv, ev, crArr)
                wsLiv.Cells(rr, COL_H).Value = ComputeColH(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_I).Value = ComputeColI(bv, cv, dv, ev, powqArr, finRefCol)
                wsLiv.Cells(rr, COL_J).Value = ComputeColJ(bv, cv, dv, ev, powqArr)

                If rr + 1 <= UBound(livArr, 1) Then
                    bvK = CStr(livArr(rr + 1, COL_B) & "")
                    cvK = CStr(livArr(rr + 1, COL_C) & "")
                    dvK = CStr(livArr(rr + 1, COL_D) & "")
                    evK = CStr(livArr(rr + 1, COL_E) & "")
                Else
                    bvK = "": cvK = "": dvK = "": evK = ""
                End If
                wsLiv.Cells(rr, COL_K).Value = ComputeColK(bvK, cvK, dvK, evK, crArr)

                wsLiv.Cells(rr, COL_O).Value = ComputeColO(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_T).Value = ComputeColT(bv, cv, dv, ev, powqArr)
            Next matchItem
            updatedCount = updatedCount + 1
        Next strKey
    End If

    ' ---------------------------------------------------------
    '  SAVE NEW SNAPSHOT
    ' ---------------------------------------------------------
    Application.StatusBar = "Suivi Update: Saving snapshot..."
    jsonSnapshot = SerializeSnapshotToJson(crArr)
    WriteTextFile statusPath, jsonSnapshot

    msg = "Update completed successfully." & vbCrLf & _
          insertedCount & " new STR block(s) inserted (" & totalInsertedRows & " rows total)." & vbCrLf & _
          updatedCount & " existing STR(s) recomputed."
    MsgBox msg, vbInformation, "Suivi Update"
    GoTo Cleanup

ErrHandler:
    MsgBox "Update failed: " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical, "Suivi Update"
    Resume Cleanup

Cleanup:
    On Error Resume Next
    If lockCreated Then
        If FileExists(SHARED_FOLDER_PATH & "LOCK.txt") Then
            Kill SHARED_FOLDER_PATH & "LOCK.txt"
        End If
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
