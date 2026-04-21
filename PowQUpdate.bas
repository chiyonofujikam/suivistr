Option Explicit

Private m_PowQBatchMode As Boolean
Private m_PowQBatchStatus As Object
Private m_PowQCurrentProcess As String

' Returns True when value is empty or numeric zero.
Private Function IsZeroOrEmpty(v As Variant) As Boolean
    If IsEmpty(v) Then
        IsZeroOrEmpty = True
    ElseIf VarType(v) = vbString Then
        IsZeroOrEmpty = (Len(v) = 0)
    ElseIf IsNumeric(v) Then
        IsZeroOrEmpty = (CDbl(v) = 0)
    Else
        IsZeroOrEmpty = False
    End If
End Function

' Converts mixed string/number values to Double.
Private Function ToNum(v As Variant) As Double
    If IsEmpty(v) Then ToNum = 0: Exit Function
    If IsNumeric(v) Then ToNum = CDbl(v): Exit Function
    Dim s As String
    s = Trim(CStr(v))
    s = Replace(s, " ", "")
    s = Replace(s, ",", ".")
    If Len(s) = 0 Then ToNum = 0: Exit Function
    On Error Resume Next
    ToNum = Val(s)
    On Error GoTo 0
End Function

' Floors a numeric value safely, returns empty on error.
Private Function SafeFloor(v As Variant) As Variant
    On Error GoTo ErrH
    SafeFloor = Fix(ToNum(v))
    Exit Function
ErrH:
    SafeFloor = ""
End Function

' Converts input value to a Date serial when possible.
Private Function ParseDateValue(v As Variant) As Variant
    On Error GoTo ErrH

    If IsZeroOrEmpty(v) Then
        ParseDateValue = ""
        Exit Function
    End If

    If IsDate(v) Then
        ParseDateValue = CDate(v)
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) = 0 Then
            ParseDateValue = ""
        Else
            ParseDateValue = CDate(CDbl(v))
        End If
        Exit Function
    End If

    ParseDateValue = ""
    Exit Function

ErrH:
    ParseDateValue = ""
End Function

' Parses hour strings like "2h" or "2 heures".
Private Function ParseHours(v As Variant) As Variant
    If IsZeroOrEmpty(v) Then
        ParseHours = ""
        Exit Function
    End If
    Dim s As String
    s = Trim(CStr(v))
    If Right(s, 1) = "h" Then
        ParseHours = ToNum(Replace(s, "h", ""))
    Else
        s = Replace(s, "heure", "")
        s = Replace(s, "s", "")
        ParseHours = ToNum(s)
    End If
End Function

' Normalizes header text for reliable comparisons.
Private Function NormalizeHeaderText(ByVal s As String) As String
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, ChrW$(8203), "")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeHeaderText = LCase$(s)
End Function

' Makes table headers unique for Excel without changing visible label.
Private Function BuildExcelSafeHeaders(ByVal headers As Variant) As Variant
    Dim result() As Variant
    Dim seen As Object
    Dim normKey As String
    Dim label As String
    Dim i As Long
    Dim dupIdx As Long

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    ReDim result(LBound(headers) To UBound(headers))
    For i = LBound(headers) To UBound(headers)
        label = CStr(headers(i))
        If Len(label) = 0 Then label = "Col_" & CStr(i)
        normKey = NormalizeHeaderText(label)
        If Len(normKey) = 0 Then normKey = "COL_" & CStr(i)

        If Not seen.Exists(normKey) Then
            seen.Add normKey, 0
            result(i) = label
        Else
            dupIdx = CLng(seen(normKey)) + 1
            seen(normKey) = dupIdx
            result(i) = label & String$(dupIdx, ChrW$(8203))
        End If
    Next i

    BuildExcelSafeHeaders = result
End Function

' Returns True when UVR destination column is a date field.
Private Function IsUVRDateColumnIndex(ByVal colIdx As Long) As Boolean
    Select Case colIdx
        Case 8, 9, 11, 12, 15, 17, 19 ' H, I, K, L, O, Q, S
            IsUVRDateColumnIndex = True
        Case Else
            IsUVRDateColumnIndex = False
    End Select
End Function

' Sanitizes imported UVR values: Excel errors and #N/A become blank.
Private Function SanitizeUVRImportedValue(ByVal v As Variant) As Variant
    If IsError(v) Then
        SanitizeUVRImportedValue = ""
        Exit Function
    End If

    If VarType(v) = vbString Then
        If UCase$(Trim$(CStr(v))) = "#N/A" Then
            SanitizeUVRImportedValue = ""
            Exit Function
        End If
    End If

    SanitizeUVRImportedValue = v
End Function

' Removes brackets and digits from a label.
Private Function StripBracketsAndDigits(ByVal s As String) As String
    s = Replace(s, "[", "")
    s = Replace(s, "%]", "")
    Dim d As Long
    For d = 0 To 9
        s = Replace(s, CStr(d), "")
    Next d
    StripBracketsAndDigits = s
End Function

' Returns True when the sheet contains no values.
Private Function IsWorksheetEmpty(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    IsWorksheetEmpty = (Application.WorksheetFunction.CountA(ws.Cells) = 0)
    On Error GoTo 0
End Function

' Gets worksheet by name and returns True when found.
Private Function TryGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    TryGetWorksheet = Not ws Is Nothing
End Function

Private Sub PowQBatchStart()
    m_PowQBatchMode = True
    Set m_PowQBatchStatus = CreateObject("Scripting.Dictionary")
    m_PowQCurrentProcess = ""
End Sub

Private Sub PowQBatchSetProcess(ByVal processName As String)
    m_PowQCurrentProcess = processName
    If m_PowQBatchStatus Is Nothing Then Set m_PowQBatchStatus = CreateObject("Scripting.Dictionary")
    If Not m_PowQBatchStatus.Exists(processName) Then
        m_PowQBatchStatus.Add processName, "PENDING"
    End If
End Sub

Private Sub PowQBatchMarkSuccess()
    If Not m_PowQBatchMode Then Exit Sub
    If Len(m_PowQCurrentProcess) = 0 Then Exit Sub
    m_PowQBatchStatus(m_PowQCurrentProcess) = "OK"
End Sub

Private Sub PowQBatchMarkError(ByVal messageText As String)
    If Not m_PowQBatchMode Then Exit Sub
    If Len(m_PowQCurrentProcess) = 0 Then Exit Sub
    If Not m_PowQBatchStatus.Exists(m_PowQCurrentProcess) Then
        m_PowQBatchStatus.Add m_PowQCurrentProcess, "ERROR: " & messageText
        Exit Sub
    End If
    If Left$(CStr(m_PowQBatchStatus(m_PowQCurrentProcess)), 2) <> "OK" Then
        m_PowQBatchStatus(m_PowQCurrentProcess) = "ERROR: " & messageText
    End If
End Sub

Private Sub PowQBatchFinish()
    Dim p As Variant
    Dim statusText As String
    Dim allOk As Boolean
    Dim summary As String

    If m_PowQBatchStatus Is Nothing Then Exit Sub

    allOk = True
    summary = "Statut PowQ Tout :" & vbCrLf & vbCrLf
    For Each p In m_PowQBatchStatus.Keys
        statusText = CStr(m_PowQBatchStatus(p))
        summary = summary & "- " & CStr(p) & " : " & statusText & vbCrLf
        If Left$(statusText, 2) <> "OK" Then allOk = False
    Next p

    If allOk Then
        VBA.Interaction.MsgBox summary, vbInformation, "PowQ Tout"
    Else
        VBA.Interaction.MsgBox summary, vbExclamation, "PowQ Tout"
    End If

    m_PowQBatchMode = False
    Set m_PowQBatchStatus = Nothing
    m_PowQCurrentProcess = ""
End Sub

' Local wrapper to suppress popups during PowQ Tout batch mode.
Private Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "") As VbMsgBoxResult
    Dim iconPart As Long
    If Not m_PowQBatchMode Then
        MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
        Exit Function
    End If

    iconPart = (Buttons And (vbCritical Or vbExclamation Or vbInformation Or vbQuestion))
    If iconPart = vbCritical Or iconPart = vbExclamation Then
        PowQBatchMarkError Prompt
    End If

    If (Buttons And vbYesNo) = vbYesNo Then
        MsgBox = vbYes
    Else
        MsgBox = vbOK
    End If
End Function


' Rebuilds PowQ_Extract from a selected input workbook.
Sub Update_PowQ_Exract(Optional ByVal externalWorkbookPath As String = "", Optional ByVal inputSheetName As String = "")
    Dim inputFilePath As Variant
    Dim wbInput As Workbook
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowInput As Long
    Dim lastRowOutput As Long
    Dim dataCount As Long
    Dim inp As Variant
    Dim out() As Variant
    Dim i As Long
    Dim j As Long
    Dim col As Long
    Dim outB As Variant, outC As Variant, outF As Variant
    Dim outG As Variant, outT As Variant, outU As Variant
    Dim tbl As ListObject
    Dim lo As ListObject
    Dim tblRange As Range
    Dim targetRange As Range
    Dim filtered() As Variant
    Dim filteredCount As Long
    Dim headers As Variant
    Dim existingTableNames As String
    Dim userChoice As VbMsgBoxResult
    Dim shouldFocusOutput As Boolean

    ' Ask user for the source workbook when no external path is provided.
    If Len(externalWorkbookPath) > 0 Then
        inputFilePath = externalWorkbookPath
    Else
        inputFilePath = Application.GetOpenFilename( _
            FileFilter:="Fichiers Excel (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
            Title:="Sélectionner le fichier d'entrée PowQ")
    End If

    If inputFilePath = False Or Len(CStr(inputFilePath)) = 0 Then Exit Sub

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(SH_EXTRACT)
    On Error GoTo 0

    If wsOutput Is Nothing Then
        MsgBox "La feuille '" & SH_EXTRACT & "' est introuvable dans ce classeur.", vbCritical, "Erreur"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' Load source data (columns A:X).
    Set wbInput = Workbooks.Open(CStr(inputFilePath), ReadOnly:=True, UpdateLinks:=0)
    If Len(inputSheetName) > 0 Then
        If Not TryGetWorksheet(wbInput, inputSheetName, wsInput) Then
            MsgBox "La feuille '" & inputSheetName & "' est introuvable dans le fichier d'entrée.", vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If
    Else
        Set wsInput = wbInput.Worksheets(1)
    End If

    If IsWorksheetEmpty(wsInput) Then
        If Len(inputSheetName) > 0 Then
            MsgBox "La feuille source '" & wsInput.Name & "' est vide : la mise à jour PowQ_Extract n'est pas correctement faite.", vbExclamation, "Attention"
        Else
            MsgBox "La feuille source du fichier d'entrée est vide : la mise à jour PowQ_Extract n'est pas correctement faite.", vbExclamation, "Attention"
        End If
        wbInput.Close False
        GoTo Cleanup
    End If

    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    If lastRowInput < 2 Then
        If Len(inputSheetName) > 0 Then
            MsgBox "La feuille source '" & wsInput.Name & "' ne contient pas de données exploitables : la mise à jour PowQ_Extract n'est pas correctement faite.", vbExclamation, "Attention"
        Else
            MsgBox "La feuille source du fichier d'entrée ne contient pas de données exploitables : la mise à jour PowQ_Extract n'est pas correctement faite.", vbExclamation, "Attention"
        End If
        wbInput.Close False
        GoTo Cleanup
    End If

    dataCount = lastRowInput - 1
    inp = wsInput.Range("A2:X" & lastRowInput).Value2

    wbInput.Close False
    Set wbInput = Nothing

    ' Build output rows (columns A:Y).
    ReDim out(1 To dataCount, 1 To 25)

    For i = 1 To dataCount
        If IsZeroOrEmpty(inp(i, 13)) Then out(i, 2) = "" Else out(i, 2) = inp(i, 13)
        If IsZeroOrEmpty(inp(i, 2)) Then out(i, 3) = "" Else out(i, 3) = inp(i, 2)
        If IsZeroOrEmpty(inp(i, 3)) Then out(i, 4) = "" Else out(i, 4) = inp(i, 3)
        If IsZeroOrEmpty(inp(i, 1)) Then out(i, 5) = "" Else out(i, 5) = inp(i, 1)
        If IsZeroOrEmpty(inp(i, 14)) Then out(i, 6) = "" Else out(i, 6) = inp(i, 14)
        If IsZeroOrEmpty(inp(i, 4)) Then out(i, 7) = "" Else out(i, 7) = inp(i, 4)
        out(i, 8) = ParseDateValue(inp(i, 5))
        out(i, 9) = ParseDateValue(inp(i, 6))
        out(i, 10) = ParseHours(inp(i, 7))
        If IsZeroOrEmpty(inp(i, 9)) Then out(i, 11) = "" Else out(i, 11) = ToNum(inp(i, 9)) * 100
        If IsZeroOrEmpty(inp(i, 8)) Then out(i, 12) = "" Else out(i, 12) = ToNum(inp(i, 8)) * 100
        out(i, 13) = ParseDateValue(inp(i, 10))
        out(i, 14) = ParseDateValue(inp(i, 11))
        out(i, 15) = ParseHours(inp(i, 12))
        If IsZeroOrEmpty(inp(i, 17)) Then out(i, 16) = "" Else out(i, 16) = inp(i, 17)
        If IsZeroOrEmpty(inp(i, 18)) Then out(i, 17) = "" Else out(i, 17) = inp(i, 18)
        If IsZeroOrEmpty(inp(i, 17)) Then out(i, 18) = "" Else out(i, 18) = SafeFloor(inp(i, 17))
        If IsZeroOrEmpty(inp(i, 18)) Then out(i, 19) = "" Else out(i, 19) = SafeFloor(inp(i, 18))
        out(i, 20) = inp(i, 19)

        If IsZeroOrEmpty(inp(i, 23)) Then
            out(i, 21) = ""
        ElseIf CStr(inp(i, 23)) = "/" Then
            out(i, 21) = ""
        Else
            out(i, 21) = inp(i, 23)
        End If

        If IsZeroOrEmpty(inp(i, 24)) Then out(i, 25) = "" Else out(i, 25) = ToNum(inp(i, 24))

        outF = out(i, 6)
        outG = out(i, 7)

        If CStr(outF) = "Reprise suite valid" Then out(i, 22) = "x" Else out(i, 22) = ""
        out(i, 23) = ""

        If IsZeroOrEmpty(outG) Then
            out(i, 24) = ""
        Else
            out(i, 24) = StripBracketsAndDigits(CStr(outG))
        End If

        outB = out(i, 2)
        outC = out(i, 3)
        outT = out(i, 20)
        outU = out(i, 21)

        If CStr(outB) = "" Then
            out(i, 1) = ""
        ElseIf CStr(outC) = "" Then
            out(i, 1) = ""
        ElseIf CStr(outT) = "Non" Then
            out(i, 1) = ""
        ElseIf CStr(outU) = "" Then
            out(i, 1) = CStr(outB) & "/" & CStr(outC) & "/" & CStr(outF)
        Else
            out(i, 1) = CStr(outB) & "/" & CStr(outC) & "/" & CStr(outF) & "/Sprint " & CStr(outU)
        End If
    Next i

    ' Keep only valid rows: column A not empty and column U numeric/empty.
    filteredCount = 0
    For i = 1 To dataCount
        If CStr(out(i, 1)) = "" Then GoTo SkipRow
        If CStr(out(i, 21)) <> "" And Not IsNumeric(out(i, 21)) Then GoTo SkipRow
        filteredCount = filteredCount + 1
SkipRow:
    Next i

    If filteredCount = 0 Then
        MsgBox "Aucune ligne valide trouvée.", vbExclamation, "Attention"
        GoTo Cleanup
    End If

    ReDim filtered(1 To filteredCount, 1 To 25)
    j = 0
    For i = 1 To dataCount
        If CStr(out(i, 1)) = "" Then GoTo SkipRow2
        If CStr(out(i, 21)) <> "" And Not IsNumeric(out(i, 21)) Then GoTo SkipRow2
        j = j + 1
        For col = 1 To 25
            filtered(j, col) = out(i, col)
        Next col
SkipRow2:
    Next i

    ' Force expected headers/order for PowQ_Extract table.
    headers = GetPowQExtractHeaders()
    For col = 0 To UBound(headers)
        wsOutput.Cells(1, col + 1).Value = headers(col)
    Next col

    ' Replace existing output data.
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    If lastRowOutput >= 2 Then
        wsOutput.Range("A2:Y" & lastRowOutput).ClearContents
    End If

    wsOutput.Range("A2:Y" & (filteredCount + 1)).Value = filtered

    With wsOutput
        .Range("H2:H" & (filteredCount + 1)).NumberFormat = "dd/mm/yyyy"
        .Range("I2:I" & (filteredCount + 1)).NumberFormat = "dd/mm/yyyy"
        .Range("M2:M" & (filteredCount + 1)).NumberFormat = "dd/mm/yyyy"
        .Range("N2:N" & (filteredCount + 1)).NumberFormat = "dd/mm/yyyy"
    End With

    ' Recreate output table after removing any existing table on PowQ_Extract.
    Set targetRange = wsOutput.Range("A1:Y" & (filteredCount + 1))

    If wsOutput.ListObjects.Count > 0 Then
        existingTableNames = ""
        For Each lo In wsOutput.ListObjects
            If Len(existingTableNames) > 0 Then existingTableNames = existingTableNames & vbCrLf
            existingTableNames = existingTableNames & "- " & lo.Name
        Next lo

        userChoice = MsgBox( _
            "Les tableaux suivants seront supprimés de la feuille '" & SH_EXTRACT & "' :" & vbCrLf & vbCrLf & _
            existingTableNames & vbCrLf & vbCrLf & _
            "Continuer et reconstruire le tableau '" & TBL_EXTRACT & "' ?", _
            vbYesNo + vbExclamation + vbDefaultButton2, _
            "Confirmation suppression tableaux")

        If userChoice <> vbYes Then
            MsgBox "Le processus est arrêté. Cette mise à jour ne peut pas fonctionner tant que le tableau n'est pas supprimé " & _
                   "ou renommé en '" & TBL_EXTRACT & "'.", vbCritical, "Arrêt du traitement"
            GoTo Cleanup
        End If
    End If

    Do While wsOutput.ListObjects.Count > 0
        Set lo = wsOutput.ListObjects(wsOutput.ListObjects.Count)
        lo.Unlist
    Loop

    Set tblRange = targetRange
    Set tbl = wsOutput.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = TBL_EXTRACT

    PowQBatchMarkSuccess
    shouldFocusOutput = True
    MsgBox "Mise à jour de PowQ_Extract terminée." & vbCrLf & _
           filteredCount & " lignes écrites (" & (dataCount - filteredCount) & " lignes ignorées).", vbInformation, "Terminé"
    GoTo Cleanup

ErrHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
    If Not wbInput Is Nothing Then wbInput.Close False

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If shouldFocusOutput Then
        On Error Resume Next
        wsOutput.Activate
        wsOutput.Cells(1, 1).Select
        On Error GoTo 0
    End If
End Sub


' Finds a header in a worksheet and returns its row/column.
Private Function FindHeaderPosition(ByVal ws As Worksheet, ByVal headerName As String, ByRef headerRow As Long, ByRef headerCol As Long) As Boolean
    Dim firstCell As Range

    Set firstCell = ws.Cells.Find(What:=headerName, After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If firstCell Is Nothing Then
        FindHeaderPosition = False
        Exit Function
    End If

    headerRow = firstCell.Row
    headerCol = firstCell.Column
    FindHeaderPosition = True
End Function

' Returns the row count below a header until the last non-empty cell.
Private Function GetColumnDataCount(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerCol As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, headerCol).End(xlUp).Row

    If lastRow <= headerRow Then
        GetColumnDataCount = 0
    Else
        GetColumnDataCount = lastRow - headerRow
    End If
End Function

' Returns the target table from PowQ_EDU_CE_VHST sheet.
Private Function GetEduTable(ByVal ws As Worksheet) As ListObject
    On Error Resume Next
    Set GetEduTable = ws.ListObjects(TBL_EDU)
    On Error GoTo 0

    If GetEduTable Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set GetEduTable = ws.ListObjects(1)
        End If
    End If
End Function

' Updates only selected columns in PowQ_EDU_CE_VHST from a selected workbook.
Sub Update_PowQ_EDU_CE_VHST(Optional ByVal externalWorkbookPath As String = "", Optional ByVal inputSheetName As String = "")
    Dim inputFilePath As Variant
    Dim wbInput As Workbook
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim tbl As ListObject
    Dim requiredHeaders As Variant
    Dim headerRows() As Long
    Dim headerCols() As Long
    Dim colDataCounts() As Long
    Dim headerRowRef As Long
    Dim tableRowCount As Long
    Dim rowsToCopy As Long
    Dim sourceCount As Long
    Dim ignoredRows As Long
    Dim i As Long
    Dim r As Long
    Dim outCol As Long
    Dim rngTarget As Range
    Dim srcData As Variant
    Dim tgtData As Variant
    Dim sourceMaxCount As Long
    Dim tableColCount As Long
    Dim createRange As Range
    Dim shouldFocusOutput As Boolean

    requiredHeaders = GetPowQEduHeaders()
    ReDim headerRows(LBound(requiredHeaders) To UBound(requiredHeaders))
    ReDim headerCols(LBound(requiredHeaders) To UBound(requiredHeaders))
    ReDim colDataCounts(LBound(requiredHeaders) To UBound(requiredHeaders))

    If Len(externalWorkbookPath) > 0 Then
        inputFilePath = externalWorkbookPath
    Else
        inputFilePath = Application.GetOpenFilename( _
            FileFilter:="Fichiers Excel (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
            Title:="Sélectionner le fichier d'entrée EDU_CE_VHST")
    End If

    If inputFilePath = False Then Exit Sub

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(SH_VHST)
    On Error GoTo 0
    If wsOutput Is Nothing Then
        MsgBox "La feuille '" & SH_VHST & "' est introuvable dans ce classeur.", vbCritical, "Erreur"
        Exit Sub
    End If

    Set tbl = GetEduTable(wsOutput)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    Set wbInput = Workbooks.Open(CStr(inputFilePath), ReadOnly:=True, UpdateLinks:=0)
    If Len(inputSheetName) > 0 Then
        If Not TryGetWorksheet(wbInput, inputSheetName, wsInput) Then
            MsgBox "La feuille '" & inputSheetName & "' est introuvable dans le fichier d'entrée.", vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If
    Else
        If Not TryGetWorksheet(wbInput, SH_IN_VHST, wsInput) Then
            MsgBox "La feuille '" & SH_IN_VHST & "' est introuvable dans le fichier d'entrée.", vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If
    End If

    If IsWorksheetEmpty(wsInput) Then
        If Len(inputSheetName) > 0 Then
            MsgBox "La feuille source '" & wsInput.Name & "' est vide : la mise à jour " & SH_VHST & " n'est pas correctement faite.", vbExclamation, "Attention"
        Else
            MsgBox "La feuille source du fichier d'entrée est vide : la mise à jour " & SH_VHST & " n'est pas correctement faite.", vbExclamation, "Attention"
        End If
        wbInput.Close False
        GoTo Cleanup
    End If

    headerRowRef = -1
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not FindHeaderPosition(wsInput, CStr(requiredHeaders(i)), headerRows(i), headerCols(i)) Then
            MsgBox "Colonne obligatoire introuvable dans le fichier d'entrée : " & CStr(requiredHeaders(i)), vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If

        If headerRowRef = -1 Then
            headerRowRef = headerRows(i)
        ElseIf headerRows(i) <> headerRowRef Then
            MsgBox "Les en-têtes requis ne sont pas sur la même ligne dans le fichier d'entrée.", vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If

        colDataCounts(i) = GetColumnDataCount(wsInput, headerRows(i), headerCols(i))
    Next i

    sourceMaxCount = 0
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If colDataCounts(i) > sourceMaxCount Then sourceMaxCount = colDataCounts(i)
    Next i

    tableColCount = UBound(requiredHeaders) - LBound(requiredHeaders) + 1
    If sourceMaxCount < 1 Then sourceMaxCount = 1

    If tbl Is Nothing Then
        For i = LBound(requiredHeaders) To UBound(requiredHeaders)
            wsOutput.Cells(1, i - LBound(requiredHeaders) + 1).Value = CStr(requiredHeaders(i))
        Next i
        Set createRange = wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(sourceMaxCount + 1, tableColCount))
        Set tbl = wsOutput.ListObjects.Add(xlSrcRange, createRange, , xlYes)
        tbl.Name = TBL_EDU
    ElseIf tbl.ListRows.Count <> sourceMaxCount Then
        Set createRange = wsOutput.Range(tbl.HeaderRowRange.Cells(1, 1), tbl.HeaderRowRange.Cells(1, tableColCount).Offset(sourceMaxCount, 0))
        tbl.Resize createRange
    End If

    tableRowCount = tbl.ListRows.Count

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        On Error Resume Next
        outCol = tbl.ListColumns(CStr(requiredHeaders(i))).Index
        On Error GoTo ErrHandler
        If outCol = 0 Then
            MsgBox "La colonne '" & CStr(requiredHeaders(i)) & "' est absente du tableau de sortie.", vbCritical, "Erreur"
            wbInput.Close False
            GoTo Cleanup
        End If

        Set rngTarget = tbl.ListColumns(outCol).DataBodyRange

        sourceCount = colDataCounts(i)
        ignoredRows = 0

        ReDim tgtData(1 To tableRowCount, 1 To 1)
        If sourceCount > 0 Then
            srcData = wsInput.Range(wsInput.Cells(headerRows(i) + 1, headerCols(i)), wsInput.Cells(headerRows(i) + sourceCount, headerCols(i))).Value
            rowsToCopy = sourceCount
            If rowsToCopy > tableRowCount Then
                rowsToCopy = tableRowCount
                ignoredRows = sourceCount - tableRowCount
            End If

            For r = 1 To rowsToCopy
                tgtData(r, 1) = srcData(r, 1)
            Next r
        End If

        ' Fill remaining rows with blanks when source column is shorter.
        If rowsToCopy < tableRowCount Then
            For r = rowsToCopy + 1 To tableRowCount
                tgtData(r, 1) = ""
            Next r
        End If

        rngTarget.Value = tgtData

        If ignoredRows > 0 Then
            MsgBox "La colonne '" & CStr(requiredHeaders(i)) & "' contient " & sourceCount & _
                   " lignes dans le fichier d'entrée." & vbCrLf & _
                   ignoredRows & " ligne(s) ont été ignorée(s) car le tableau de destination contient " & tableRowCount & " lignes.", _
                   vbExclamation, "Attention - Lignes ignorées"
        End If

        outCol = 0
        rowsToCopy = 0
    Next i

    wbInput.Close False
    Set wbInput = Nothing

    PowQBatchMarkSuccess
    shouldFocusOutput = True
    MsgBox "Mise à jour de PowQ_EDU_CE_VHST terminée." & vbCrLf & _
           "Colonnes mises à jour : Nom_STR, Sprint, Collaborateurs, Sociétés, Info_Complet.", vbInformation, "Terminé"
    GoTo Cleanup

ErrHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
    If Not wbInput Is Nothing Then wbInput.Close False

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If shouldFocusOutput Then
        On Error Resume Next
        wsOutput.Activate
        wsOutput.Cells(1, 1).Select
        On Error GoTo 0
    End If
End Sub


Sub Update_PowQ_Suivi_UVR(Optional ByVal externalWorkbookPath As String = "", Optional ByVal inputSheetName As String = "")
    Dim inputFilePath As Variant
    Dim wbInput As Workbook
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim openedWindow As Window
    Dim wasAlreadyOpen As Boolean
    Dim tbl As ListObject
    Dim lo As ListObject
    Dim existingTableNames As String
    Dim userChoice As VbMsgBoxResult
    Dim requiredHeaders As Variant
    Dim srcHeaderCols() As Long
    Dim outData() As Variant
    Dim sourceData As Variant
    Dim targetRange As Range
    Dim tblRange As Range
    Dim firstDataRow As Long
    Dim headerRow As Long
    Dim dataLastRow As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim r As Long
    Dim srcCol As Long
    Dim outHeader As String
    Dim foundHeader As Range
    Dim headerColMap As Object
    Dim headerNextIdx As Object
    Dim colList As Collection
    Dim headerKey As String
    Dim c As Long
    Dim useInputHeaders As Boolean
    Dim shouldFocusOutput As Boolean
    Dim targetTableName As String
    Dim availableHeaders As String
    Dim pos As Long
    Dim outputHeaders As Variant

    targetTableName = TBL_UVR
    headerRow = UVR_HEADER_ROW
    firstDataRow = UVR_FIRST_DATA_ROW

    If Len(externalWorkbookPath) > 0 Then
        inputFilePath = externalWorkbookPath
    Else
        inputFilePath = Application.GetOpenFilename( _
            FileFilter:="Fichiers Excel (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
            Title:="Sélectionner le fichier d'entrée UVR")
    End If

    If inputFilePath = False Or Len(CStr(inputFilePath)) = 0 Then Exit Sub

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(SH_UVR)
    On Error GoTo 0
    If wsOutput Is Nothing Then
        MsgBox "La feuille '" & SH_UVR & "' est introuvable dans ce classeur.", vbCritical, "Erreur"
        Exit Sub
    End If

    On Error Resume Next
    Set tbl = wsOutput.ListObjects(targetTableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        If wsOutput.ListObjects.Count = 0 Then
            useInputHeaders = True
        Else
            existingTableNames = ""
            For Each lo In wsOutput.ListObjects
                If Len(existingTableNames) > 0 Then existingTableNames = existingTableNames & vbCrLf
                existingTableNames = existingTableNames & "- " & lo.Name
            Next lo

            userChoice = MsgBox( _
                "Le tableau '" & targetTableName & "' est introuvable sur '" & SH_UVR & "'." & vbCrLf & vbCrLf & _
                "Tableau(x) trouvé(s) :" & vbCrLf & existingTableNames & vbCrLf & vbCrLf & _
                "Voulez-vous supprimer ces tableaux et créer '" & targetTableName & "' ?", _
                vbYesNo + vbExclamation + vbDefaultButton2, _
                "Confirmation suppression tableaux")

            If userChoice <> vbYes Then
                MsgBox "Le processus est arrêté. Cette mise à jour ne fonctionnera pas tant que le tableau n'est pas supprimé " & _
                       "ou renommé en '" & targetTableName & "'.", vbCritical, "Arrêt du traitement"
                Exit Sub
            End If

            Set tbl = wsOutput.ListObjects(1)
        End If
    End If

    If Not tbl Is Nothing Then
        colCount = tbl.ListColumns.Count
        ReDim requiredHeaders(1 To colCount)
        For i = 1 To colCount
            requiredHeaders(i) = CStr(tbl.ListColumns(i).Name)
        Next i
    End If

    On Error GoTo ErrHandler
    wasAlreadyOpen = False
    For Each wb In Workbooks
        If StrComp(CStr(wb.FullName), CStr(inputFilePath), vbTextCompare) = 0 Then
            Set wbInput = wb
            wasAlreadyOpen = True
            Exit For
        End If
    Next wb

    If wbInput Is Nothing Then
        Set wbInput = Workbooks.Open(CStr(inputFilePath), ReadOnly:=True, UpdateLinks:=0)
        On Error Resume Next
        Set openedWindow = wbInput.Windows(1)
        If Not openedWindow Is Nothing Then openedWindow.Visible = False
        On Error GoTo ErrHandler
    End If

    If Not TryGetWorksheet(wbInput, SH_IN_UVR, wsInput) Then
        MsgBox "La feuille '" & SH_IN_UVR & "' est introuvable dans le fichier d'entrée.", vbCritical, "Erreur"
        If Not wasAlreadyOpen Then wbInput.Close False
        Exit Sub
    End If

    If IsWorksheetEmpty(wsInput) Then
        MsgBox "La feuille source '" & wsInput.Name & "' est vide : la mise à jour " & SH_UVR & " n'est pas correctement faite.", vbExclamation, "Attention"
        If Not wasAlreadyOpen Then wbInput.Close False
        Exit Sub
    End If

    Set headerColMap = CreateObject("Scripting.Dictionary")
    headerColMap.CompareMode = vbTextCompare
    Set headerNextIdx = CreateObject("Scripting.Dictionary")
    headerNextIdx.CompareMode = vbTextCompare
    For c = 1 To 23 ' A:W
        headerKey = NormalizeHeaderText(CStr(wsInput.Cells(headerRow, c).Value2 & ""))
        If Len(headerKey) > 0 Then
            If Not headerColMap.Exists(headerKey) Then
                Set colList = New Collection
                colList.Add c
                headerColMap.Add headerKey, colList
                headerNextIdx.Add headerKey, 1
            Else
                Set colList = headerColMap(headerKey)
                colList.Add c
            End If
        End If
    Next c

    If useInputHeaders Then
        colCount = 0
        For c = 1 To 23 ' A:W
            If Len(Trim$(CStr(wsInput.Cells(headerRow, c).Value2 & ""))) > 0 Then
                colCount = colCount + 1
            End If
        Next c
        If colCount = 0 Then
            MsgBox "La feuille '" & SH_IN_UVR & "' ne contient aucun en-tête exploitable sur A:W (ligne " & headerRow & ").", vbCritical, "Erreur"
            If Not wasAlreadyOpen Then wbInput.Close False
            Exit Sub
        End If
        ReDim requiredHeaders(1 To colCount)
        i = 0
        For c = 1 To 23 ' A:W
            outHeader = Trim$(CStr(wsInput.Cells(headerRow, c).Value2 & ""))
            If Len(outHeader) > 0 Then
                i = i + 1
                requiredHeaders(i) = outHeader
            End If
        Next c
    End If

    ReDim srcHeaderCols(1 To colCount)
    availableHeaders = ""
    For i = 1 To colCount
        outHeader = CStr(requiredHeaders(i))
        headerKey = NormalizeHeaderText(outHeader)
        If Not headerColMap.Exists(headerKey) Then
            MsgBox "Le fichier source (feuille '" & SH_IN_UVR & "') n'a pas les mêmes colonnes que le tableau '" & _
                   targetTableName & "' de '" & SH_UVR & "'." & vbCrLf & _
                   "Colonne manquante : " & outHeader, vbCritical, "Erreur de correspondance colonnes"
            If Not wasAlreadyOpen Then wbInput.Close False
            Exit Sub
        End If
        Set colList = headerColMap(headerKey)
        pos = CLng(headerNextIdx(headerKey))
        If pos > colList.Count Then pos = colList.Count
        srcHeaderCols(i) = CLng(colList(pos))
        If pos < colList.Count Then headerNextIdx(headerKey) = pos + 1
        If Len(availableHeaders) > 0 Then availableHeaders = availableHeaders & ", "
        availableHeaders = availableHeaders & outHeader
    Next i

    dataLastRow = headerRow
    For i = 1 To colCount
        srcCol = srcHeaderCols(i)
        r = wsInput.Cells(wsInput.Rows.Count, srcCol).End(xlUp).Row
        If r > dataLastRow Then dataLastRow = r
    Next i

    If dataLastRow < firstDataRow Then
        rowCount = 0
    Else
        rowCount = dataLastRow - firstDataRow + 1
    End If

    If rowCount > 0 Then
        ReDim outData(1 To rowCount, 1 To colCount)
        For i = 1 To colCount
            srcCol = srcHeaderCols(i)
            sourceData = wsInput.Range(wsInput.Cells(firstDataRow, srcCol), wsInput.Cells(dataLastRow, srcCol)).Value2
            For r = 1 To rowCount
                sourceData(r, 1) = SanitizeUVRImportedValue(sourceData(r, 1))
                If IsUVRDateColumnIndex(i) Then
                    outData(r, i) = ParseDateValue(sourceData(r, 1))
                Else
                    outData(r, i) = sourceData(r, 1)
                End If
            Next r
        Next i
    End If

    If Not wasAlreadyOpen Then wbInput.Close False
    Set wbInput = Nothing

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    If userChoice = vbYes Then
        Do While wsOutput.ListObjects.Count > 0
            Set lo = wsOutput.ListObjects(wsOutput.ListObjects.Count)
            lo.Unlist
        Loop
    ElseIf Not tbl Is Nothing Then
        tbl.Unlist
    End If

    wsOutput.Cells.ClearContents

    outputHeaders = BuildExcelSafeHeaders(requiredHeaders)
    For i = 1 To colCount
        wsOutput.Cells(1, i).Value = outputHeaders(i)
    Next i

    If rowCount > 0 Then
        wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(rowCount + 1, colCount)).Value2 = outData
        wsOutput.Range("H2:H" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("I2:I" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("K2:K" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("L2:L" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("O2:O" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("Q2:Q" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        wsOutput.Range("S2:S" & (rowCount + 1)).NumberFormat = "dd/mm/yyyy"
        Set targetRange = wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(rowCount + 1, colCount))
    Else
        Set targetRange = wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, colCount))
    End If

    Set tblRange = targetRange
    Set tbl = wsOutput.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = targetTableName

    PowQBatchMarkSuccess
    shouldFocusOutput = True
    MsgBox "Mise à jour de " & SH_UVR & " terminée depuis la feuille '" & SH_IN_UVR & "'." & vbCrLf & _
           rowCount & " ligne(s) chargée(s).", vbInformation, "Terminé"
    GoTo Cleanup

ErrHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
    If Not wbInput Is Nothing Then
        If Not wasAlreadyOpen Then wbInput.Close False
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If shouldFocusOutput Then
        On Error Resume Next
        wsOutput.Activate
        wsOutput.Cells(1, 1).Select
        On Error GoTo 0
    End If
End Sub

' Runs all PowQ updates using 3 file dialogs.
Sub Update_PowQ_All()
    Dim inputFileEDU As String
    Dim inputFileExtract As String
    Dim inputFileUVR As String
    inputFileEDU = PickPowQInputFile("PowQ Tout - Sélectionner le fichier EDU_CE_VHST")
    If Len(inputFileEDU) = 0 Then Exit Sub

    inputFileUVR = PickPowQInputFile("PowQ Tout - Sélectionner le fichier UVR (Global)")
    If Len(inputFileUVR) = 0 Then Exit Sub

    inputFileExtract = PickPowQInputFile("PowQ Tout - Sélectionner le fichier Extract")
    If Len(inputFileExtract) = 0 Then Exit Sub
    If Not ConfirmPowQAllFiles(inputFileEDU, inputFileUVR, inputFileExtract) Then Exit Sub

    PowQBatchStart
    PowQBatchSetProcess "EDU_CE_VHST"
    Update_PowQ_EDU_CE_VHST inputFileEDU, SH_IN_VHST
    PowQBatchSetProcess "UVR"
    Update_PowQ_Suivi_UVR inputFileUVR, SH_IN_UVR
    PowQBatchSetProcess "Extract"
    Update_PowQ_Exract inputFileExtract
    PowQBatchFinish
End Sub

Private Function PickPowQInputFile(ByVal dialogTitle As String) As String
    Dim picked As Variant
    picked = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        Title:=dialogTitle)
    If VarType(picked) = vbBoolean Then
        If CBool(picked) = False Then
            PickPowQInputFile = ""
            Exit Function
        End If
    End If
    PickPowQInputFile = CStr(picked)
End Function

Private Function ConfirmPowQAllFiles(ByVal fileEDU As String, ByVal fileUVR As String, ByVal fileExtract As String) As Boolean
    Dim resp As VbMsgBoxResult
    resp = VBA.Interaction.MsgBox( _
        "Confirmer les fichiers PowQ Tout :" & vbCrLf & vbCrLf & _
        "EDU_CE_VHST :" & vbCrLf & Dir$(fileEDU) & vbCrLf & vbCrLf & _
        "UVR (Global) :" & vbCrLf & Dir$(fileUVR) & vbCrLf & vbCrLf & _
        "Extract :" & vbCrLf & Dir$(fileExtract), _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Confirmation PowQ Tout")
    ConfirmPowQAllFiles = (resp = vbYes)
End Function
