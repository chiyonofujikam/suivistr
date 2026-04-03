Option Explicit

Public Sub UpdateSuiviLivrable()
    ' -- Variable declarations --
    Dim lockPath As String
    Dim statusPath As String
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
    Dim livRow As Long
    Dim srcRow As Long
    Dim bv As String
    Dim cv As String
    Dim dv As String
    Dim ev As String
    Dim bvK As String
    Dim cvK As String
    Dim dvK As String
    Dim evK As String
    Dim iterItem As Variant
    Dim lockCreated As Boolean
    Dim insertedCount As Long
    Dim updatedCount As Long
    Dim uniqueSTRs As Object
    Dim strsToInsert As Object
    Dim strsToUpdate As Object
    Dim strKey As Variant
    Dim blockSize As Long
    Dim blockEnd As Long
    Dim rr As Long
    Dim matchRows As Collection
    Dim lastTmpCol As Long
    Dim matchItem As Variant

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
    '  LOAD SHEET DATA INTO ARRAYS (one-time reads)
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
    '  COMPUTE DIFF  (new STR values + modified rows)
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
    blockSize = TMP_LAST_ROW - TMP_FIRST_ROW + 1

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
    '  INSERT: copy full template block for each new unique STR
    ' ---------------------------------------------------------
    insertedCount = 0

    If strsToInsert.Count > 0 Then
        Application.StatusBar = "Suivi Update: Inserting " & strsToInsert.Count & " STR block(s)..."

        lastTmpCol = wsTmp.UsedRange.Column + wsTmp.UsedRange.Columns.Count - 1

        firstNewRow = GetLastDataRow(wsLiv, COL_B) + 1
        If firstNewRow < LIV_FIRST_ROW Then firstNewRow = LIV_FIRST_ROW
        insertRow = firstNewRow

        ' -- Pass 1: copy template formatting + values, set STR, add borders --
        For Each strKey In strsToInsert.Keys
            blockEnd = insertRow + blockSize - 1

            ' Copy formatting for all columns from template
            wsTmp.Range(wsTmp.Cells(TMP_FIRST_ROW, 1), _
                        wsTmp.Cells(TMP_LAST_ROW, lastTmpCol)).Copy
            wsLiv.Cells(insertRow, 1).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False

            ' Copy values for cols C, D, E from template
            wsTmp.Range(wsTmp.Cells(TMP_FIRST_ROW, COL_C), _
                        wsTmp.Cells(TMP_LAST_ROW, COL_E)).Copy
            wsLiv.Cells(insertRow, COL_C).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False

            ' Set column B to actual STR value
            For rr = insertRow To blockEnd
                wsLiv.Cells(rr, COL_B).Value = CStr(strKey)
            Next rr

            ' Thick border at the bottom of the block across all columns
            With wsLiv.Range(wsLiv.Cells(blockEnd, 1), _
                             wsLiv.Cells(blockEnd, lastTmpCol)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With

            insertRow = insertRow + blockSize
            insertedCount = insertedCount + 1
        Next strKey

        ' Reload array so pass 2 reads the fresh values
        livArr = LoadSheetData(wsLiv)

        ' -- Pass 2: compute formula columns for every inserted row --
        insertRow = firstNewRow
        For Each strKey In strsToInsert.Keys
            blockEnd = insertRow + blockSize - 1
            For rr = insertRow To blockEnd
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
            insertRow = insertRow + blockSize
        Next strKey
    End If

    ' ---------------------------------------------------------
    '  UPDATE: recompute formula columns for all Livrables rows of modified STRs
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

    Dim msg As String
    msg = "Update completed successfully." & vbCrLf & _
          insertedCount & " new STR block(s) inserted (" & blockSize & " rows each)." & vbCrLf & _
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
