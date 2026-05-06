Option Explicit

Public Sub UpdateSuiviLivrable()
    ' Rebuilds `Suivi_Livrables` from current sources.
    Dim lockCreated As Boolean
    Dim wsCR As Worksheet
    Dim lockValue As String
    Dim crArr As Variant
    Dim powqArr As Variant
    Dim uvrArr As Variant
    Dim vhstArr As Variant
    Dim configArr As Variant
    Dim finRefCol As Long
    Dim uvrColMap As Object
    Dim maxSprintMap As Object
    Dim wsLiv As Worksheet
    Dim insertRow As Long
    Dim insertedCount As Long
    Dim totalInsertedRows As Long
    Dim strKey As Variant
    Dim lastLivCol As Long
    Dim strSprints As Collection
    Dim maxSprintKey As String
    Dim nrows As Long
    Dim br As Variant
    Dim lastBorderCol As Long
    Dim msg As String
    Dim errNumber As Long
    Dim errMessage As String
    Dim errSource As String
    Dim vhstSTRMap As Object
    Dim fonctions As Collection
    Dim typeLivrables As Collection
    Dim typeLivrableFallbackResp As VbMsgBoxResult
    Dim manualColsSnapshot As Object
    Dim crIndex As Object
    Dim powqCompIndex As Object
    Dim powqAIndex As Object
    Dim uvrIndex As Object
    Dim strPlans As Collection
    Dim plan As Variant
    Dim totalRows As Long
    Dim bcdeArr As Variant
    Dim blockTop As Long
    Dim blockBottom As Long

    On Error GoTo ErrHandler
    lockCreated = False
    Set wsCR = ThisWorkbook.Sheets(SH_CR)

    ' Lock (prevent concurrent runs).
    If Trim$(CStr(wsCR.Range(LOCK_CELL_ADDR).Value & "")) <> "" Then
        WaitWhileLocked wsCR, LOCK_CELL_ADDR
    End If
    lockValue = LOCK_PREFIX & Environ$("USERNAME") & LOCK_SEPARATOR & Format$(Now, LOCK_DATE_FORMAT)
    wsCR.Range(LOCK_CELL_ADDR).Value = lockValue
    lockCreated = True

    If wsCR.AutoFilterMode Then wsCR.AutoFilterMode = False
    If SheetExists(SH_LIV) Then
        If ThisWorkbook.Sheets(SH_LIV).AutoFilterMode Then ThisWorkbook.Sheets(SH_LIV).AutoFilterMode = False
    End If

    wsCR.Range(LOCK_CELL_ADDR).Locked = False
    wsCR.Protect Password:=PROTECT_PASSWORD, UserInterfaceOnly:=False

    ' Load sources.
    ValidateRequiredSheets

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Mise a jour Suivi : chargement des donnees..."

    crArr = LoadSheetData(ThisWorkbook.Sheets(SH_CR))
    powqArr = LoadSheetData(ThisWorkbook.Sheets(SH_EXTRACT))
    uvrArr = LoadSheetData(ThisWorkbook.Sheets(SH_UVR))
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))
    configArr = LoadSheetData(ThisWorkbook.Sheets(SH_CONFIG))
    finRefCol = FindFinRefColumn(powqArr)

    Call CheckAndOfferUpdateVHSTMaxSprints(ThisWorkbook.Sheets(SH_VHST), crArr, vhstArr)
    vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
    Set manualColsSnapshot = CaptureSuiviLivrableManualValues(wsLiv)
    lastLivCol = wsLiv.UsedRange.Column + wsLiv.UsedRange.Columns.Count - 1
    If lastLivCol < 1 Then lastLivCol = COL_Y
    lastBorderCol = lastLivCol
    If lastBorderCol < COL_Y Then lastBorderCol = COL_Y

    Set vhstSTRMap = BuildSTRMapVHST(vhstArr)
    Set uvrColMap = BuildUVRColumnMap(wsLiv, uvrArr)
    Set maxSprintMap = BuildMaxSprintMapVHST(vhstArr)
    Set fonctions = BuildFonctionsFromConfig(configArr)
    Set typeLivrables = BuildTypeLivrablesFromConfig(configArr)
    If fonctions.Count = 0 Then
        Err.Raise vbObjectError + 2001, "UpdateSuiviLivrable", _
                  "Aucune fonction disponible dans " & SH_CONFIG & " (colonne '" & HDR_FONCTIONS & "')."
    End If
    If typeLivrables.Count = 0 Then
        typeLivrableFallbackResp = MsgBox( _
            "Aucun type livrable disponible dans " & SH_CONFIG & " (colonne '" & HDR_TYPE_LIVRABLE & "')." & vbCrLf & vbCrLf & _
            "Voulez-vous utiliser les types de livrables par defaut ADL1 et SwDS ?", _
            vbYesNo + vbQuestion, "Mise a jour Suivi")
        If typeLivrableFallbackResp = vbYes Then
            EnsureDefaultTypeLivrablesInConfig ThisWorkbook.Sheets(SH_CONFIG)
            configArr = LoadSheetData(ThisWorkbook.Sheets(SH_CONFIG))
            Set typeLivrables = BuildTypeLivrablesFromConfig(configArr)
        Else
            Err.Raise vbObjectError + 2002, "UpdateSuiviLivrable", _
                      "Aucun type livrable disponible dans " & SH_CONFIG & " (colonne '" & HDR_TYPE_LIVRABLE & "')."
        End If
    End If

    If vhstSTRMap.Count = 0 Then
        MsgBox "Aucune STR disponible dans " & SH_VHST & ".", vbExclamation, "Mise a jour Suivi"
        GoTo Cleanup
    End If

    ' Build lookup indexes.
    Application.StatusBar = "Mise a jour Suivi : indexation des sources..."
    Set crIndex = BuildCRIndex(crArr)
    Set powqCompIndex = BuildPowQCompositeIndex(powqArr)
    Set powqAIndex = BuildPowQAIndex(powqArr)
    Set uvrIndex = BuildUVRIndex(uvrArr)

    Application.StatusBar = "Mise a jour Suivi : reconstruction complete de " & SH_LIV & "..."

    insertedCount = 0
    totalInsertedRows = 0

    If wsLiv.AutoFilterMode Then wsLiv.AutoFilterMode = False
    insertRow = GetLastDataRow(wsLiv, COL_B)
    If insertRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & insertRow).Delete Shift:=xlUp
    End If

    ' Pass 1: plan blocks and total rows.
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

    ' Single insert for the whole rebuild.
    If totalRows > 0 Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & (LIV_FIRST_ROW + totalRows - 1)).Insert Shift:=xlDown
    End If

    ' Pass 2: generate blocks and compute columns.
    insertRow = LIV_FIRST_ROW
    For Each plan In strPlans
        Set strSprints = plan(1)
        maxSprintKey = CStr(plan(2))
        nrows = CLng(plan(3))

        bcdeArr = BuildSTRBlockBCDEMatrix(CStr(plan(0)), strSprints, fonctions, typeLivrables)

        br = InsertGeneratedSTRBlock(wsLiv, insertRow, CStr(plan(0)), strSprints, _
                                     fonctions, typeLivrables, lastBorderCol, maxSprintKey)
        blockTop = CLng(br(0))
        blockBottom = CLng(br(1))

        WriteSTRBlockComputedColumns wsLiv, blockTop, blockBottom, bcdeArr, _
                                     crArr, powqArr, uvrArr, _
                                     crIndex, powqCompIndex, powqAIndex, uvrIndex, _
                                     finRefCol, uvrColMap, maxSprintKey

        insertedCount = insertedCount + 1
        totalInsertedRows = totalInsertedRows + nrows
        insertRow = insertRow + nrows
    Next plan

    ' Finalize formatting and restore manual columns.
    RestoreSuiviLivrableManualValues wsLiv, manualColsSnapshot
    RebuildSuiviLivrablesBorders wsLiv, lastBorderCol
    ApplySuiviLivrablesColumnFormats wsLiv

    msg = "Mise a jour terminee avec succes." & vbCrLf & vbCrLf & _
          "Reconstruction complete de " & SH_LIV & " :" & vbCrLf & _
          "  - " & insertedCount & " bloc(s) STR genere(s)" & vbCrLf & _
          "  - " & totalInsertedRows & " ligne(s) regeneree(s)"
    MsgBox msg, vbInformation, "Mise a jour Suivi"
    GoTo Cleanup

ErrHandler:
    ' Log and report runtime errors.
    errNumber = Err.Number
    errMessage = Err.Description
    errSource = Err.Source

    On Error Resume Next
    LogErrorToSheet errNumber, errSource, errMessage, Now

    MsgBox "Echec de la mise a jour : " & errMessage & _
           " (Erreur " & errNumber & ")", vbCritical, "Mise a jour Suivi"
    Resume Cleanup

Cleanup:
    ' Release lock and restore application settings.
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

' Computes derived columns for one generated block.
Private Sub WriteSTRBlockComputedColumns(ByVal wsLiv As Worksheet, _
                                         ByVal blockTop As Long, ByVal blockBottom As Long, _
                                         bcdeArr As Variant, _
                                         crArr As Variant, powqArr As Variant, uvrArr As Variant, _
                                         crIndex As Object, powqCompIndex As Object, _
                                         powqAIndex As Object, uvrIndex As Object, _
                                         ByVal finRefCol As Long, _
                                         uvrColMap As Object, _
                                         ByVal maxSprintKey As String)
    Dim nrows As Long
    Dim rr As Long
    Dim bv As String, cv As String, dv As String, ev As String
    Dim arrA() As Variant
    Dim arrBK() As Variant
    Dim arrM() As Variant
    Dim arrO() As Variant
    Dim arrT() As Variant
    Dim arrUX() As Variant
    Dim hasUVR As Boolean
    Dim isMaxSprint As Boolean
    Dim colIdx As Long

    nrows = blockBottom - blockTop + 1
    If nrows <= 0 Then Exit Sub

    ReDim arrA(1 To nrows, 1 To 1)
    ReDim arrBK(1 To nrows, 1 To 10) ' covers B(2) .. K(11)
    ReDim arrM(1 To nrows, 1 To 1)
    ReDim arrO(1 To nrows, 1 To 1)
    ReDim arrT(1 To nrows, 1 To 1)
    ReDim arrUX(1 To nrows, 1 To 4) ' covers U(21) .. X(24)

    hasUVR = False
    If Not uvrColMap Is Nothing Then
        hasUVR = (uvrColMap.Count > 0)
    End If

    For rr = 1 To nrows
        bv = CStr(bcdeArr(rr, 1) & "")
        cv = CStr(bcdeArr(rr, 2) & "")
        dv = CStr(bcdeArr(rr, 3) & "")
        ev = CStr(bcdeArr(rr, 4) & "")

        arrA(rr, 1) = ComputeColA(bv, cv, dv, ev)

        ' B/C/D/E (stable keys).
        arrBK(rr, 1) = bv
        arrBK(rr, 2) = cv
        arrBK(rr, 3) = dv
        arrBK(rr, 4) = ev
        arrBK(rr, 5) = ComputeColFFast(bv, cv, dv, ev, crArr, crIndex)
        arrBK(rr, 6) = ComputeColGFast(bv, cv, dv, ev, crArr, crIndex)
        arrBK(rr, 7) = ComputeColHFast(bv, cv, dv, ev, powqArr, powqCompIndex)
        arrBK(rr, 8) = ComputeColIFast(bv, cv, dv, ev, powqArr, finRefCol, powqAIndex)
        arrBK(rr, 9) = ComputeColJFast(bv, cv, dv, ev, powqArr, powqCompIndex)
        arrBK(rr, 10) = ComputeColKFast(bv, cv, dv, ev, crArr, crIndex)

        arrM(rr, 1) = ComputeColMFast(bv, cv, dv, ev, powqArr, powqAIndex)
        arrO(rr, 1) = ComputeColOFast(bv, cv, dv, ev, powqArr, powqAIndex)
        arrT(rr, 1) = ComputeColTFast(bv, cv, dv, ev, powqArr, powqAIndex)

        If hasUVR And maxSprintKey <> "" Then
            isMaxSprint = (NormalizeSprintKey(dv) = maxSprintKey)
            If isMaxSprint Then
                For colIdx = COL_U To COL_X
                    If uvrColMap.Exists(colIdx) Then
                        arrUX(rr, colIdx - COL_U + 1) = ComputeUVRCellFast( _
                            bv, cv, ev, uvrArr, CLng(uvrColMap(colIdx)), colIdx, uvrIndex)
                    End If
                Next colIdx
            End If
        End If
    Next rr

    ' Bulk writes (leave L/N/P/Q/R/S/Y to snapshot restore).
    wsLiv.Range(wsLiv.Cells(blockTop, COL_A), wsLiv.Cells(blockBottom, COL_A)).Value = arrA
    wsLiv.Range(wsLiv.Cells(blockTop, COL_B), wsLiv.Cells(blockBottom, COL_K)).Value = arrBK
    wsLiv.Range(wsLiv.Cells(blockTop, COL_M), wsLiv.Cells(blockBottom, COL_M)).Value = arrM
    wsLiv.Range(wsLiv.Cells(blockTop, COL_O), wsLiv.Cells(blockBottom, COL_O)).Value = arrO
    wsLiv.Range(wsLiv.Cells(blockTop, COL_T), wsLiv.Cells(blockBottom, COL_T)).Value = arrT
    If hasUVR And maxSprintKey <> "" Then
        wsLiv.Range(wsLiv.Cells(blockTop, COL_U), wsLiv.Cells(blockBottom, COL_X)).Value = arrUX
    End If
End Sub
