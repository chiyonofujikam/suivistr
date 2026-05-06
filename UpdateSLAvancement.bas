Option Explicit

Public Sub UpdateSLAvancement()
    Dim lockCreated As Boolean
    Dim wsCR As Worksheet
    Dim lockValue As String
    Dim wsLiv As Worksheet
    Dim crArr As Variant
    Dim vhstArr As Variant
    Dim configArr As Variant
    Dim crIndex As Object
    Dim nrows As Long
    Dim outK() As Variant
    Dim rr As Long
    Dim bv As String, cv As String, dv As String, ev As String
    Dim oldLivArr As Variant
    Dim oldLastRow As Long
    Dim oldRowsByKey As Object
    Dim rowKey As String
    Dim lastLivCol As Long
    Dim lastBorderCol As Long
    Dim manualColsSnapshot As Object
    Dim vhstSTRMap As Object
    Dim maxSprintMap As Object
    Dim fonctions As Collection
    Dim typeLivrables As Collection
    Dim typeLivrableFallbackResp As VbMsgBoxResult
    Dim strPlans As Collection
    Dim plan As Variant
    Dim strKey As Variant
    Dim strSprints As Collection
    Dim maxSprintKey As String
    Dim insertRow As Long
    Dim totalRows As Long
    Dim bcdeArr As Variant
    Dim newLastRow As Long
    Dim restoreAJ() As Variant
    Dim restoreLY() As Variant
    Dim oldR As Long
    Dim c As Long
    Dim errNumber As Long
    Dim errMessage As String
    Dim errSource As String

    On Error GoTo ErrHandler
    lockCreated = False

    ValidateRequiredSheets
    Set wsCR = ThisWorkbook.Sheets(SH_CR)
    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)

    ' Lock (prevent concurrent runs).
    If Trim$(CStr(wsCR.Range(LOCK_CELL_ADDR).Value & "")) <> "" Then
        WaitWhileLocked wsCR, LOCK_CELL_ADDR
    End If
    lockValue = LOCK_PREFIX & Environ$("USERNAME") & LOCK_SEPARATOR & Format$(Now, LOCK_DATE_FORMAT)
    wsCR.Range(LOCK_CELL_ADDR).Value = lockValue
    lockCreated = True

    If wsCR.AutoFilterMode Then wsCR.AutoFilterMode = False
    If wsLiv.AutoFilterMode Then wsLiv.AutoFilterMode = False

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "MAJ Avancement : chargement des donnees..."

    crArr = LoadSheetData(wsCR)
    Set crIndex = BuildCRIndex(crArr)

    ' Snapshot existing row values keyed by B/C/D/E so we can rebuild structure without losing data.
    oldLastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    Set oldRowsByKey = CreateObject("Scripting.Dictionary")
    oldRowsByKey.CompareMode = vbBinaryCompare
    If oldLastRow >= LIV_FIRST_ROW Then
        oldLivArr = wsLiv.Range(wsLiv.Cells(1, 1), wsLiv.Cells(oldLastRow, COL_Y)).Value
        For rr = LIV_FIRST_ROW To oldLastRow
            rowKey = Trim$(CStr(oldLivArr(rr, COL_B) & "")) & vbTab & _
                     Trim$(CStr(oldLivArr(rr, COL_C) & "")) & vbTab & _
                     Trim$(CStr(oldLivArr(rr, COL_D) & "")) & vbTab & _
                     Trim$(CStr(oldLivArr(rr, COL_E) & ""))
            If Replace(rowKey, vbTab, "") <> "" Then
                If Not oldRowsByKey.Exists(rowKey) Then oldRowsByKey.Add rowKey, rr
            End If
        Next rr
    End If

    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))
    configArr = LoadSheetData(ThisWorkbook.Sheets(SH_CONFIG))

    Call CheckAndOfferUpdateVHSTMaxSprints(ThisWorkbook.Sheets(SH_VHST), crArr, vhstArr)
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))

    Set manualColsSnapshot = CaptureSuiviLivrableManualValues(wsLiv)
    lastLivCol = wsLiv.UsedRange.Column + wsLiv.UsedRange.Columns.Count - 1
    If lastLivCol < 1 Then lastLivCol = COL_Y
    lastBorderCol = lastLivCol
    If lastBorderCol < COL_Y Then lastBorderCol = COL_Y

    Set vhstSTRMap = BuildSTRMapVHST(vhstArr)
    Set maxSprintMap = BuildMaxSprintMapVHST(vhstArr)
    Set fonctions = BuildFonctionsFromConfig(configArr)
    Set typeLivrables = BuildTypeLivrablesFromConfig(configArr)
    If fonctions.Count = 0 Then Err.Raise vbObjectError + 2101, "UpdateSLAvancement", "Aucune fonction disponible dans " & SH_CONFIG & "."
    If typeLivrables.Count = 0 Then
        typeLivrableFallbackResp = MsgBox( _
            "Aucun type livrable disponible dans " & SH_CONFIG & " (colonne '" & HDR_TYPE_LIVRABLE & "')." & vbCrLf & vbCrLf & _
            "Voulez-vous utiliser les types de livrables par defaut ADL1 et SwDS ?", _
            vbYesNo + vbQuestion, "MAJ Avancement")
        If typeLivrableFallbackResp = vbYes Then
            EnsureDefaultTypeLivrablesInConfig ThisWorkbook.Sheets(SH_CONFIG)
            configArr = LoadSheetData(ThisWorkbook.Sheets(SH_CONFIG))
            Set typeLivrables = BuildTypeLivrablesFromConfig(configArr)
        Else
            Err.Raise vbObjectError + 2102, "UpdateSLAvancement", "Aucun type livrable disponible dans " & SH_CONFIG & "."
        End If
    End If
    If vhstSTRMap.Count = 0 Then GoTo Cleanup

    If wsLiv.AutoFilterMode Then wsLiv.AutoFilterMode = False
    insertRow = GetLastDataRow(wsLiv, COL_B)
    If insertRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & insertRow).Delete Shift:=xlUp
    End If

    Set strPlans = New Collection
    totalRows = 0
    For Each strKey In vhstSTRMap.Keys
        Set strSprints = GetTargetSprintsForSTR(crArr, CStr(strKey), maxSprintMap)
        maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints)
        nrows = GeneratedBlockRowCount(strSprints, fonctions, typeLivrables)
        If nrows > 0 Then
            strPlans.Add Array(CStr(strKey), strSprints, maxSprintKey, nrows)
            totalRows = totalRows + nrows
        End If
    Next strKey

    If totalRows > 0 Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & (LIV_FIRST_ROW + totalRows - 1)).Insert Shift:=xlDown
    End If

    insertRow = LIV_FIRST_ROW
    For Each plan In strPlans
        Set strSprints = plan(1)
        maxSprintKey = CStr(plan(2))
        nrows = CLng(plan(3))

        bcdeArr = BuildSTRBlockBCDEMatrix(CStr(plan(0)), strSprints, fonctions, typeLivrables)
        Call InsertGeneratedSTRBlock(wsLiv, insertRow, CStr(plan(0)), strSprints, fonctions, typeLivrables, lastBorderCol, maxSprintKey)
        insertRow = insertRow + nrows
    Next plan

    ' Restore old values (all columns except K), then recompute K on the rebuilt structure.
    newLastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If newLastRow >= LIV_FIRST_ROW Then
        nrows = newLastRow - LIV_FIRST_ROW + 1
        ReDim restoreAJ(1 To nrows, 1 To 10) ' A..J
        ReDim restoreLY(1 To nrows, 1 To (COL_Y - 12 + 1)) ' L..Y
        ReDim outK(1 To nrows, 1 To 1)

        bcdeArr = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_B), wsLiv.Cells(newLastRow, COL_E)).Value
        For rr = 1 To nrows
            bv = CStr(bcdeArr(rr, 1) & "")
            cv = CStr(bcdeArr(rr, 2) & "")
            dv = CStr(bcdeArr(rr, 3) & "")
            ev = CStr(bcdeArr(rr, 4) & "")

            rowKey = Trim$(bv) & vbTab & Trim$(cv) & vbTab & Trim$(dv) & vbTab & Trim$(ev)
            If oldRowsByKey.Exists(rowKey) Then
                oldR = CLng(oldRowsByKey(rowKey))
                For c = COL_A To COL_J
                    restoreAJ(rr, c) = oldLivArr(oldR, c)
                Next c
                For c = 12 To COL_Y
                    restoreLY(rr, c - 12 + 1) = oldLivArr(oldR, c)
                Next c
            End If

            outK(rr, 1) = ComputeColKFast(bv, cv, dv, ev, crArr, crIndex)
        Next rr

        wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_A), wsLiv.Cells(newLastRow, COL_J)).Value = restoreAJ
        wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 12), wsLiv.Cells(newLastRow, COL_Y)).Value = restoreLY
        wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_K), wsLiv.Cells(newLastRow, COL_K)).Value = outK
    End If

    RestoreSuiviLivrableManualValues wsLiv, manualColsSnapshot
    RebuildSuiviLivrablesBorders wsLiv, lastBorderCol
    ApplySuiviLivrablesColumnFormats wsLiv
    MsgBox "MAJ Avancement terminee (colonne K).", vbInformation, "MAJ Avancement"
    GoTo Cleanup

ErrHandler:
    errNumber = Err.Number
    errMessage = Err.Description
    errSource = Err.Source
    On Error Resume Next
    LogErrorToSheet errNumber, errSource, errMessage, Now
    MsgBox "Echec de la mise a jour Avancement : " & errMessage & _
           " (Erreur " & errNumber & ")", vbCritical, "MAJ Avancement"
    Resume Cleanup

Cleanup:
    On Error Resume Next
    If lockCreated Then
        If Left$(CStr(wsCR.Range(LOCK_CELL_ADDR).Value & ""), Len(LOCK_PREFIX & Environ$("USERNAME"))) = LOCK_PREFIX & Environ$("USERNAME") Then
            wsCR.Unprotect Password:=PROTECT_PASSWORD
            wsCR.Range(LOCK_CELL_ADDR).ClearContents
        End If
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

