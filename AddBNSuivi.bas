Option Explicit

' Syncs the BN_Suivi sheet from VHST (STR/max sprints/fonctions) and Suivi_CR (joined AA text).
Public Sub AddBNSuivi()
    Dim wsVHST As Worksheet
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

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsVHST = ThisWorkbook.Worksheets(SH_VHST)
    Set wsCR = ThisWorkbook.Worksheets(SH_CR)
    Set wsBN = ThisWorkbook.Worksheets("BN_Suivi dossier Safety")

    lastVHSTRow = GetLastDataRow(wsVHST, COL_A)
    lastCRRow = GetLastDataRow(wsCR, COL_B)
    lastBNRow = wsBN.Cells(wsBN.Rows.Count, COL_B).End(xlUp).Row
    If lastBNRow < 3 Then lastBNRow = 2

    firstDataRowBN = 3

    Set dictCombos = CreateObject("Scripting.Dictionary")
    dictCombos.CompareMode = vbTextCompare

    Set fonctions = New Collection
    LoadAllFonctions wsVHST, lastVHSTRow, fonctions

    For r = 2 To lastVHSTRow
        strVal = Trim$(CStr(wsVHST.Cells(r, COL_A).Value))
        If strVal <> "" Then
            If IsNumeric(wsVHST.Cells(r, COL_B).Value) Then
                maxSprint = CLng(wsVHST.Cells(r, COL_B).Value)
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
               addedCount & " ligne(s) ajoutee(s) dans ""BN_Suivi dossier Safety"".", _
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
    MsgBox "Erreur lors de la mise a jour de 'BN_Suivi dossier Safety' : " & Err.Description, _
           vbExclamation, "BN_Suivi"
    Resume CleanExit
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
           StrComp(Trim$(CStr(wsCR.Cells(r, COL_O).Value)), "OUI", vbTextCompare) = 0 Then

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

Private Sub LoadAllFonctions(wsVHST As Worksheet, ByVal lastVHSTRow As Long, ByRef fonctions As Collection)
    Dim seen As Object
    Dim r As Long
    Dim rawFonctions As String
    Dim oneFonctions As Collection
    Dim fonctionItem As Variant

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For r = 2 To lastVHSTRow
        rawFonctions = Trim$(CStr(wsVHST.Cells(r, COL_F).Value))
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

