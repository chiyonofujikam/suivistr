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


' Rebuilds PowQ_Extract from a selected input workbook.
Sub Update_PowQ_Exract()
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
    Dim tblRange As Range
    Dim filtered() As Variant
    Dim filteredCount As Long

    ' Ask user for the source workbook.
    inputFilePath = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        Title:="Sélectionner le fichier d'entrée PowQ")

    If inputFilePath = False Then Exit Sub

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("PowQ_Extract")
    On Error GoTo 0

    If wsOutput Is Nothing Then
        MsgBox "La feuille 'PowQ_Extract' est introuvable dans ce classeur.", vbCritical, "Erreur"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' Load source data (columns A:X).
    Set wbInput = Workbooks.Open(CStr(inputFilePath), ReadOnly:=True, UpdateLinks:=0)
    Set wsInput = wbInput.Sheets(1)

    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    If lastRowInput < 2 Then
        MsgBox "Le fichier d'entrée ne contient pas de données.", vbExclamation, "Attention"
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
        If IsZeroOrEmpty(inp(i, 5)) Then out(i, 8) = "" Else out(i, 8) = SafeFloor(inp(i, 5))
        If IsZeroOrEmpty(inp(i, 6)) Then out(i, 9) = "" Else out(i, 9) = SafeFloor(inp(i, 6))
        out(i, 10) = ParseHours(inp(i, 7))
        If IsZeroOrEmpty(inp(i, 9)) Then out(i, 11) = "" Else out(i, 11) = ToNum(inp(i, 9)) * 100
        If IsZeroOrEmpty(inp(i, 8)) Then out(i, 12) = "" Else out(i, 12) = ToNum(inp(i, 8)) * 100
        If IsZeroOrEmpty(inp(i, 10)) Then out(i, 13) = "" Else out(i, 13) = SafeFloor(inp(i, 10))
        If IsZeroOrEmpty(inp(i, 11)) Then out(i, 14) = "" Else out(i, 14) = SafeFloor(inp(i, 11))
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

    ' Replace existing output data.
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    If lastRowOutput >= 2 Then
        wsOutput.Range("A2:Y" & lastRowOutput).ClearContents
    End If

    wsOutput.Range("A2:Y" & (filteredCount + 1)).Value = filtered

    With wsOutput
        .Range("H2:H" & (filteredCount + 1)).NumberFormat = "0"
        .Range("I2:I" & (filteredCount + 1)).NumberFormat = "0"
        .Range("M2:M" & (filteredCount + 1)).NumberFormat = "0"
        .Range("N2:N" & (filteredCount + 1)).NumberFormat = "0"
    End With

    ' Recreate output table.
    For Each tbl In wsOutput.ListObjects
        If tbl.Name = "Extract_MSP" Then
            tbl.Unlist
            Exit For
        End If
    Next tbl

    Set tblRange = wsOutput.Range("A1:Y" & (filteredCount + 1))
    Set tbl = wsOutput.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "Extract_MSP"

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
End Sub


' Updates PowQ_Suivi_UVR (placeholder until implemented).
Sub Update_PowQ_Suivi_UVR()
    MsgBox "Cette fonctionnalité est encore en cours de développement et n'est pas encore finalisée.", vbExclamation, "En cours de développement"
End Sub
