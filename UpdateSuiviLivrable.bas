Option Explicit

Public Sub UpdateSuiviLivrable()
    ' Main orchestration: lock, rebuild Suivi_Livrables from current source sheets.
    Dim lockCreated As Boolean
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
    Dim wsLiv As Worksheet
    Dim insertRow As Long
    Dim bv As String, cv As String, dv As String, ev As String
    Dim rr As Long
    Dim insertedCount As Long
    Dim totalInsertedRows As Long
    Dim strKey As Variant
    Dim lastLivCol As Long
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
    Dim vhstSTRMap As Object
    Dim fonctions As Collection
    Dim typeLivrables As Collection
    Dim typeLivrableFallbackResp As VbMsgBoxResult
    Dim logPath As String

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

    On Error Resume Next
    logPath = SHARED_FOLDER_PATH(False)
    If Err.Number <> 0 Or Trim$(logPath) = "" Then
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "La selection du dossier partage n'a pas ete finalisee correctement." & vbCrLf & _
               "La mise a jour est annulee.", vbExclamation, "Mise a jour Suivi"
        GoTo Cleanup
    End If
    On Error GoTo ErrHandler
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logPath = logPath & "error_logs.txt"

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

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
    lastLivCol = wsLiv.UsedRange.Column + wsLiv.UsedRange.Columns.Count - 1
    If lastLivCol < 1 Then lastLivCol = COL_Y
    lastBorderCol = lastLivCol
    If lastBorderCol < COL_Y Then lastBorderCol = COL_Y

    Set vhstSTRMap = BuildSTRMapVHST(vhstArr)
    Set uvrColMap = BuildUVRColumnMap(wsLiv, uvrArr)
    Set maxSprintMap = BuildMaxSprintMapVHST(vhstArr)
    Set fonctions = BuildFonctionsFromVHST(vhstArr)
    Set typeLivrables = BuildTypeLivrablesFromVHST(vhstArr)
    If fonctions.Count = 0 Then
        Err.Raise vbObjectError + 2001, "UpdateSuiviLivrable", _
                  "Aucune fonction disponible dans " & SH_VHST & " (colonne 'Fonctions')."
    End If
    If typeLivrables.Count = 0 Then
        typeLivrableFallbackResp = MsgBox( _
            "Aucun type livrable disponible dans " & SH_VHST & " (colonne 'Type de livrable')." & vbCrLf & vbCrLf & _
            "Voulez-vous utiliser les types de livrables par defaut ADL1 et SwDS ?", _
            vbYesNo + vbQuestion, "Mise a jour Suivi")
        If typeLivrableFallbackResp = vbYes Then
            EnsureDefaultTypeLivrablesInVHST ThisWorkbook.Sheets(SH_VHST)
            vhstArr = LoadSheetData(ThisWorkbook.Sheets(SH_VHST))
            Set typeLivrables = BuildTypeLivrablesFromVHST(vhstArr)
        Else
            Err.Raise vbObjectError + 2002, "UpdateSuiviLivrable", _
                      "Aucun type livrable disponible dans " & SH_VHST & " (colonne 'Type de livrable')."
        End If
    End If

    If vhstSTRMap.Count = 0 Then
        MsgBox "Aucune STR disponible dans " & SH_VHST & ".", vbExclamation, "Mise a jour Suivi"
        GoTo Cleanup
    End If

    Application.StatusBar = "Mise a jour Suivi : reconstruction complete de " & SH_LIV & "..."

    insertedCount = 0
    totalInsertedRows = 0

    If wsLiv.AutoFilterMode Then wsLiv.AutoFilterMode = False
    insertRow = GetLastDataRow(wsLiv, COL_B)
    If insertRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & insertRow).Delete Shift:=xlUp
    End If

    For Each strKey In vhstSTRMap.Keys
        Set strSprints = GetTargetSprintsForSTR(crArr, CStr(strKey), maxSprintMap)
        maxSprintKey = GetYellowSprintKeyForSTR(CStr(strKey), maxSprintMap, strSprints)
        nrows = GeneratedBlockRowCount(strSprints, fonctions, typeLivrables)
        If nrows <= 0 Then GoTo NextStr

        insertRow = GetLastDataRow(wsLiv, COL_B) + 1
        If insertRow < LIV_FIRST_ROW Then insertRow = LIV_FIRST_ROW
        wsLiv.Rows(insertRow & ":" & (insertRow + nrows - 1)).Insert Shift:=xlDown
        br = InsertGeneratedSTRBlock(wsLiv, insertRow, CStr(strKey), strSprints, fonctions, typeLivrables, lastBorderCol, maxSprintKey)

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
            wsLiv.Cells(rr, COL_A).Value = ComputeColA(bv, cv, dv, ev)

            If maxSprintKey <> "" And NormalizeSprintKey(dv) = maxSprintKey Then
                For colI2 = COL_U To COL_X
                    If uvrColMap.Exists(colI2) Then
                        wsLiv.Cells(rr, colI2).Value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colI2)), colI2)
                    End If
                Next colI2
            End If
        Next rr

        insertedCount = insertedCount + 1
        totalInsertedRows = totalInsertedRows + nrows
NextStr:
    Next strKey

    ' Rebuild borders and persist new snapshot state.
    RebuildSuiviLivrablesBorders wsLiv, lastBorderCol
    ApplySuiviLivrablesColumnFormats wsLiv

    msg = "Mise a jour terminee avec succes." & vbCrLf & vbCrLf & _
          "Reconstruction complete de " & SH_LIV & " :" & vbCrLf & _
          "  - " & insertedCount & " bloc(s) STR genere(s)" & vbCrLf & _
          "  - " & totalInsertedRows & " ligne(s) regeneree(s)"
    MsgBox msg, vbInformation, "Mise a jour Suivi"
    GoTo Cleanup

ErrHandler:
    ' Log runtime errors and show user-facing message.
    errNumber = Err.Number
    errMessage = Err.Description
    errSource = Err.Source

    On Error Resume Next
    AppendTextFile logPath, _
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
