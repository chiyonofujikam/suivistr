Option Explicit

Public Sub UpdateSuiviLivrable()
    ' Main orchestration: lock, detect CR changes, update Suivi_Livrables, save snapshot.
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
    Dim insertRow As Long
    Dim iterItem As Variant
    Dim bv As String, cv As String, dv As String, ev As String
    Dim rr As Long
    Dim insertedCount As Long
    Dim updatedCount As Long
    Dim totalInsertedRows As Long
    Dim uniqueSTRs As Object
    Dim strsToInsert As Object
    Dim strsToUpdate As Object
    Dim strKey As Variant
    Dim matchRows As Collection
    Dim matchItem As Variant
    Dim lastLivCol As Long
    Dim blockInfo As Object
    Dim strSprints As Collection
    Dim maxSprintKey As String
    Dim nrows As Long
    Dim br As Variant
    Dim colI2 As Long
    Dim lastBorderCol As Long
    Dim msg As String
    Dim errNumber As Long
    Dim errMessage As String
    Dim errSource As String
    Dim rowKey As String
    Dim vhstSTRMap As Object
    Dim vhstKey As Variant
    Dim fonctions As Collection
    Dim minRow As Long

    On Error GoTo ErrHandler
    lockCreated = False
    Set wsCR = ThisWorkbook.Sheets(SH_CR)

    ' Acquire workbook-level lock to avoid concurrent runs.
    If Trim$(CStr(wsCR.Range("I1").Value & "")) <> "" Then
        WaitWhileLocked wsCR, "I1"
    End If
    lockValue = "LOCKED by: " & Environ$("USERNAME") & " at " & Format$(Now, "YYYY-MM-DD HH:NN:SS")
    wsCR.Range("I1").Value = lockValue
    lockCreated = True

    If wsCR.AutoFilterMode Then wsCR.AutoFilterMode = False
    If SheetExists(SH_LIV) Then
        If ThisWorkbook.Sheets(SH_LIV).AutoFilterMode Then ThisWorkbook.Sheets(SH_LIV).AutoFilterMode = False
    End If

    wsCR.Range("I1").Locked = False
    wsCR.Protect Password:="suivi_update", UserInterfaceOnly:=False

    configDir = SHARED_FOLDER_PATH & "config\"
    If Dir$(configDir, vbDirectory) = "" Then MkDir configDir

    statusPath = configDir & "status.json"

    ' Validate setup and load source arrays.
    ValidateRequiredSheets

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Mise a jour Suivi : chargement des donnees..."

    crArr = LoadSheetData(ThisWorkbook.Sheets(SH_CR))
    powqArr = LoadSheetData(ThisWorkbook.Sheets(SH_EXTRACT))
    uvrArr = LoadSheetData(ThisWorkbook.Sheets(SH_UVR))
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))
    finRefCol = FindFinRefColumn(powqArr)

    Call CheckAndOfferUpdateVHSTMaxSprints(ThisWorkbook.Sheets(SH_VHST), crArr, vhstArr)
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))

    isFirstRun = True
    If FileExists(statusPath) Then
        If FileLen(statusPath) > 0 Then isFirstRun = False
    End If

    If isFirstRun Then
        Application.StatusBar = "Mise a jour Suivi : creation du snapshot initial..."
        jsonSnapshot = SerializeSnapshotToJson(crArr)
        WriteTextFile statusPath, jsonSnapshot
    End If

    Application.StatusBar = "Mise a jour Suivi : calcul des changements..."
    If isFirstRun Then
        Set oldSnapshot = CreateObject("Scripting.Dictionary")
    Else
        oldJson = ReadTextFile(statusPath)
        Set oldSnapshot = ParseSnapshotFromJson(oldJson)
    End If

    Set newSTRs = New Collection
    Set modifiedRows = New Collection

    ' Compare current CR rows to last snapshot.
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

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
    livArr = LoadSheetData(wsLiv)
    lastLivCol = wsLiv.UsedRange.Column + wsLiv.UsedRange.Columns.Count - 1
    If lastLivCol < 1 Then lastLivCol = COL_Y
    lastBorderCol = lastLivCol
    If lastBorderCol < COL_Y Then lastBorderCol = COL_Y

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

    Set vhstSTRMap = BuildSTRMapVHST(vhstArr)
    For Each vhstKey In vhstSTRMap.Keys
        If FindRowBySTR(livArr, CStr(vhstKey)) = 0 Then
            strsToInsert(CStr(vhstKey)) = True
        End If
    Next vhstKey

    If strsToInsert.Count = 0 And strsToUpdate.Count = 0 Then
        jsonSnapshot = SerializeSnapshotToJson(crArr)
        WriteTextFile statusPath, jsonSnapshot
        MsgBox "Aucun changement detecte depuis la derniere mise a jour.", vbInformation, "Mise a jour Suivi"
        GoTo Cleanup
    End If

    insertedCount = 0
    totalInsertedRows = 0
    Set uvrColMap = BuildUVRColumnMap(wsLiv, uvrArr)
    Set maxSprintMap = BuildMaxSprintMapVHST(vhstArr)
    Set fonctions = BuildFonctionsFromVHST(vhstArr)
    If fonctions.Count = 0 Then
        Err.Raise vbObjectError + 2001, "UpdateSuiviLivrable", _
                  "Aucune fonction disponible dans " & SH_VHST & " (colonne F)."
    End If

    ' Insert new STR blocks from generated structure.
    If strsToInsert.Count > 0 Then
        Application.StatusBar = "Mise a jour Suivi : insertion de " & strsToInsert.Count & " bloc(s) STR..."

        Set blockInfo = CreateObject("Scripting.Dictionary")

        For Each strKey In strsToInsert.Keys
            Set strSprints = GetTargetSprintsForSTR(crArr, CStr(strKey), maxSprintMap)
            maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints)
            nrows = GeneratedBlockRowCount(strSprints, fonctions)
            If nrows > 0 Then
                insertRow = GetLastDataRow(wsLiv, COL_B) + 1
                If insertRow < LIV_FIRST_ROW Then insertRow = LIV_FIRST_ROW
                wsLiv.Rows(insertRow & ":" & (insertRow + nrows - 1)).Insert Shift:=xlDown
                br = InsertGeneratedSTRBlock(wsLiv, insertRow, CStr(strKey), strSprints, fonctions, lastBorderCol, maxSprintKey)
                blockInfo(CStr(strKey)) = br
                insertedCount = insertedCount + 1
                totalInsertedRows = totalInsertedRows + nrows
            End If
        Next strKey

        livArr = LoadSheetData(wsLiv)

        For Each strKey In strsToInsert.Keys
            If blockInfo.Exists(CStr(strKey)) Then
                Set strSprints = GetTargetSprintsForSTR(crArr, CStr(strKey), maxSprintMap)
                maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints)
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
                    wsLiv.Cells(rr, COL_M).Value = ComputeColM(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_K).Value = ComputeColK(bv, cv, dv, ev, crArr)
                    wsLiv.Cells(rr, COL_O).Value = ComputeColO(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_T).Value = ComputeColT(bv, cv, dv, ev, powqArr)
                    wsLiv.Cells(rr, COL_A).Value = ComputeColA(bv, cv, dv, ev)

                    If maxSprintKey <> "" And NormalizeSprintKey(dv) = maxSprintKey Then
                        For colI2 = COL_U To COL_X
                            If uvrColMap.Exists(colI2) Then
                                wsLiv.Cells(rr, colI2).Value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colI2)), colI2)
                            End If
                        Next colI2
                    End If
                Next rr
            End If
        Next strKey
    End If

    updatedCount = 0

    ' Rebuild changed STR blocks from generated structure.
    If strsToUpdate.Count > 0 Then
        Application.StatusBar = "Mise a jour Suivi : recalcul de " & strsToUpdate.Count & " STR..."

        For Each strKey In strsToUpdate.Keys
            livArr = LoadSheetData(wsLiv)
            Set matchRows = FindAllRowsBySTR(livArr, CStr(strKey))
            If matchRows.Count = 0 Then GoTo NextUpdateStr

            minRow = CLng(matchRows(1))
            For Each matchItem In matchRows
                rr = CLng(matchItem)
                If rr < minRow Then minRow = rr
            Next matchItem

            For rr = matchRows.Count To 1 Step -1
                wsLiv.Rows(CLng(matchRows(rr))).Delete
            Next rr

            Set strSprints = GetTargetSprintsForSTR(crArr, CStr(strKey), maxSprintMap)
            maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints)
            nrows = GeneratedBlockRowCount(strSprints, fonctions)
            If nrows <= 0 Then GoTo NextUpdateStr

            wsLiv.Rows(minRow & ":" & (minRow + nrows - 1)).Insert Shift:=xlDown
            br = InsertGeneratedSTRBlock(wsLiv, minRow, CStr(strKey), strSprints, fonctions, lastBorderCol, maxSprintKey)

            livArr = LoadSheetData(wsLiv)
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
                wsLiv.Cells(rr, COL_M).Value = ComputeColM(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_K).Value = ComputeColK(bv, cv, dv, ev, crArr)
                wsLiv.Cells(rr, COL_O).Value = ComputeColO(bv, cv, dv, ev, powqArr)
                wsLiv.Cells(rr, COL_T).Value = ComputeColT(bv, cv, dv, ev, powqArr)

                If maxSprintKey <> "" And NormalizeSprintKey(dv) = maxSprintKey Then
                    For colI2 = COL_U To COL_X
                        If uvrColMap.Exists(colI2) Then
                            wsLiv.Cells(rr, colI2).Value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colI2)), colI2)
                        End If
                    Next colI2
                End If

                wsLiv.Cells(rr, COL_A).Value = ComputeColA(bv, cv, dv, ev)
            Next rr
NextUpdateStr:
            updatedCount = updatedCount + 1
        Next strKey
    End If

    ' Rebuild borders and persist new snapshot state.
    RebuildSuiviLivrablesBorders wsLiv, lastBorderCol
    ApplySuiviLivrablesColumnFormats wsLiv

    Application.StatusBar = "Mise a jour Suivi : enregistrement du snapshot..."
    jsonSnapshot = SerializeSnapshotToJson(crArr)
    WriteTextFile statusPath, jsonSnapshot

    msg = "Mise a jour terminee avec succes." & vbCrLf & vbCrLf & _
          "Changements detectes dans Suivi_CR :" & vbCrLf & _
          "  - " & newSTRs.Count & " nouvelle(s) ligne(s) CR" & vbCrLf & _
          "Actions sur Suivi_Livrables :" & vbCrLf & _
          "  - " & insertedCount & " nouveau(x) bloc(s) STR insere(s) (" & totalInsertedRows & " lignes)" & vbCrLf & _
          "  - " & updatedCount & " STR existant(s) recalcule(s)"
    MsgBox msg, vbInformation, "Mise a jour Suivi"
    GoTo Cleanup

ErrHandler:
    ' Log runtime errors and show user-facing message.
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

    MsgBox "Echec de la mise a jour : " & errMessage & _
           " (Erreur " & errNumber & ")", vbCritical, "Mise a jour Suivi"
    Resume Cleanup

Cleanup:
    ' Always release lock and restore application settings.
    On Error Resume Next
    If lockCreated Then
        If Left$(CStr(wsCR.Range("I1").Value & ""), Len("LOCKED by: " & Environ$("USERNAME"))) = "LOCKED by: " & Environ$("USERNAME") Then
            wsCR.Unprotect Password:="suivi_update"
            wsCR.Range("I1").ClearContents
        End If
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
