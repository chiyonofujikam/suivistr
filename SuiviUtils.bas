Option Explicit

' DEPENDENCY: JsonConverter.bas (https://github.com/VBA-tools/VBA-JSON)
' Also requires: Tools > References > Microsoft Scripting Runtime (Dictionary)
' Constants are in Globals.bas

' ---------------------------------------------------------
'  CONFIGURATION
' ---------------------------------------------------------

Private m_SharedFolder As String

Public Function SHARED_FOLDER_PATH() As String
    Dim dlg As Object
    Dim p As String

    If m_SharedFolder <> "" Then
        SHARED_FOLDER_PATH = m_SharedFolder
        Exit Function
    End If

    Set dlg = Application.FileDialog(4)
    With dlg
        .Title = "Select the shared folder for Suivi files (LOCK.txt, status.json)"
        .ButtonName = "Select"
        If .Show = -1 Then
            p = .SelectedItems(1)
            If Right$(p, 1) <> "\" Then p = p & "\"
            m_SharedFolder = p
        Else
            Err.Raise vbObjectError + 1, "SHARED_FOLDER_PATH", _
                "No shared folder selected. Update cancelled."
        End If
    End With

    SHARED_FOLDER_PATH = m_SharedFolder
End Function

' ---------------------------------------------------------
'  SHEET VALIDATION
' ---------------------------------------------------------

Public Sub ValidateRequiredSheets()
    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim missing As String
    Dim found As Boolean
    Dim sheetList As String

    names = Array(SH_CR, SH_LIV, SH_EXTRACT, SH_TMP)
    missing = ""

    For i = LBound(names) To UBound(names)
        found = False
        For Each ws In ThisWorkbook.Worksheets
            If LCase(ws.Name) = LCase(CStr(names(i))) Then
                found = True
                Exit For
            End If
        Next ws
        If Not found Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & """" & names(i) & """"
        End If
    Next i

    If missing <> "" Then
        sheetList = ""
        For Each ws In ThisWorkbook.Worksheets
            If sheetList <> "" Then sheetList = sheetList & ", "
            sheetList = sheetList & """" & ws.Name & """"
        Next ws
        Err.Raise 9, "ValidateRequiredSheets", _
            "Sheet(s) not found: " & missing & vbCrLf & _
            "Existing sheets: " & sheetList & vbCrLf & _
            "Please rename your sheets or update the constants in SuiviUtils."
    End If
End Sub

' ---------------------------------------------------------
'  LOW-LEVEL UTILITIES
' ---------------------------------------------------------

Public Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function

Public Function ReadTextFile(path As String) As String
    Dim fNum As Integer
    fNum = FreeFile
    Open path For Input As #fNum
    If LOF(fNum) > 0 Then
        ReadTextFile = Input$(LOF(fNum), fNum)
    Else
        ReadTextFile = ""
    End If
    Close #fNum
End Function

Public Sub WriteTextFile(path As String, content As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open path For Output As #fNum
    Print #fNum, content
    Close #fNum
End Sub

Public Function ColNumToLetter(colNum As Long) As String
    Dim n As Long
    Dim result As String
    n = colNum
    result = ""
    Do While n > 0
        n = n - 1
        result = Chr(65 + (n Mod 26)) & result
        n = n \ 26
    Loop
    ColNumToLetter = result
End Function

Public Function NormalizeValue(v As Variant) As String
    If IsEmpty(v) Then
        NormalizeValue = ""
    ElseIf IsNull(v) Then
        NormalizeValue = ""
    ElseIf IsError(v) Then
        NormalizeValue = ""
    ElseIf VarType(v) = vbDate Then
        NormalizeValue = Format$(v, "YYYY-MM-DD HH:NN:SS")
    ElseIf VarType(v) = vbBoolean Then
        If CBool(v) Then NormalizeValue = "TRUE" Else NormalizeValue = "FALSE"
    ElseIf IsNumeric(v) Then
        NormalizeValue = Str$(CDbl(v))
    Else
        NormalizeValue = Trim$(CStr(v & ""))
    End If
End Function

Private Function CDblSafe(v As Variant) As Double
    If IsNumeric(v) Then
        CDblSafe = CDbl(v)
    Else
        CDblSafe = 0
    End If
End Function

Private Function IsValidPowQValue(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        IsValidPowQValue = False
    ElseIf LCase(CStr(v & "")) = "nan" Or CStr(v & "") = "" Then
        IsValidPowQValue = False
    Else
        IsValidPowQValue = True
    End If
End Function

' ---------------------------------------------------------
'  SHEET DATA LOADER
' ---------------------------------------------------------

Public Function LoadSheetData(ws As Worksheet) As Variant
    Dim ur As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rng As Range
    Dim tmpArr() As Variant

    Set ur = ws.UsedRange
    If ur Is Nothing Then
        ReDim tmpArr(1 To 1, 1 To 1)
        LoadSheetData = tmpArr
        Exit Function
    End If

    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    If rng.Cells.Count = 1 Then
        ReDim tmpArr(1 To 1, 1 To 1)
        tmpArr(1, 1) = rng.Value
        LoadSheetData = tmpArr
    Else
        LoadSheetData = rng.Value
    End If
End Function

' ---------------------------------------------------------
'  JSON SNAPSHOT -- SERIALIZE
' ---------------------------------------------------------

Public Function SerializeSnapshotToJson(crArr As Variant) As String
    Dim coll As Collection
    Dim r As Long
    Dim c As Long
    Dim numCols As Long
    Dim strVal As String
    Dim rowDict As Object
    Dim cellsDict As Object

    Set coll = New Collection
    numCols = UBound(crArr, 2)

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        strVal = CStr(crArr(r, COL_B) & "")
        If strVal = "" Then GoTo NextSnapRow

        Set rowDict = CreateObject("Scripting.Dictionary")
        rowDict("STR") = strVal
        rowDict("row") = r

        Set cellsDict = CreateObject("Scripting.Dictionary")
        For c = 1 To numCols
            cellsDict(ColNumToLetter(c)) = NormalizeValue(crArr(r, c))
        Next c
        Set rowDict("cells") = cellsDict

        coll.Add rowDict
NextSnapRow:
    Next r

    SerializeSnapshotToJson = JsonConverter.ConvertToJson(coll, Whitespace:=2)
End Function

' ---------------------------------------------------------
'  JSON SNAPSHOT -- PARSE
' ---------------------------------------------------------

Public Function ParseSnapshotFromJson(jsonStr As String) As Object
    Dim result As Object
    Dim coll As Object
    Dim item As Variant
    Dim strKey As String
    Dim cellsObj As Object
    Dim cellsDict As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set coll = JsonConverter.ParseJson(jsonStr)

    For Each item In coll
        strKey = CStr(item("STR"))
        Set cellsObj = item("cells")
        Set cellsDict = CreateObject("Scripting.Dictionary")

        For Each key In cellsObj.Keys
            cellsDict(CStr(key)) = NormalizeValue(cellsObj(key))
        Next key

        Set result(strKey) = cellsDict
    Next item

    Set ParseSnapshotFromJson = result
End Function

' ---------------------------------------------------------
'  LOOKUP HELPERS
' ---------------------------------------------------------

Public Function FindFinRefColumn(powqArr As Variant) As Long
    Dim c As Long
    For c = 1 To UBound(powqArr, 2)
        If LCase(CStr(powqArr(1, c) & "")) = "fin ref" Then
            FindFinRefColumn = c
            Exit Function
        End If
    Next c
    FindFinRefColumn = 0
End Function

Public Function GetLastDataRow(ws As Worksheet, keyCol As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < LIV_FIRST_ROW Then lastRow = 0
    GetLastDataRow = lastRow
End Function

Public Function FindRowBySTR(livArr As Variant, strVal As String) As Long
    Dim r As Long
    Dim lStr As String
    lStr = LCase(strVal)
    For r = LIV_FIRST_ROW To UBound(livArr, 1)
        If LCase(CStr(livArr(r, COL_B) & "")) = lStr Then
            FindRowBySTR = r
            Exit Function
        End If
    Next r
    FindRowBySTR = 0
End Function

Public Function FindAllRowsBySTR(livArr As Variant, strVal As String) As Collection
    Dim result As New Collection
    Dim r As Long
    Dim lStr As String
    lStr = LCase(strVal)
    For r = LIV_FIRST_ROW To UBound(livArr, 1)
        If LCase(CStr(livArr(r, COL_B) & "")) = lStr Then
            result.Add r
        End If
    Next r
    Set FindAllRowsBySTR = result
End Function

' ---------------------------------------------------------
'  SPRINT HELPERS
' ---------------------------------------------------------

' Extracts a numeric sprint key from any value: 1, "1", "Sprint 1" -> "1"
Public Function NormalizeSprintKey(v As Variant) As String
    Dim s As String
    Dim i As Long
    Dim j As Long

    If IsEmpty(v) Or IsNull(v) Then
        NormalizeSprintKey = ""
        Exit Function
    End If
    If IsNumeric(v) And Not VarType(v) = vbString Then
        NormalizeSprintKey = CStr(CLng(CDbl(v)))
        Exit Function
    End If

    s = Trim$(CStr(v & ""))
    If s = "" Then
        NormalizeSprintKey = ""
        Exit Function
    End If
    If IsNumeric(s) Then
        NormalizeSprintKey = CStr(CLng(CDbl(s)))
        Exit Function
    End If

    For i = 1 To Len(s)
        If Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9" Then
            j = i
            Do While j <= Len(s) And Mid$(s, j, 1) >= "0" And Mid$(s, j, 1) <= "9"
                j = j + 1
            Loop
            NormalizeSprintKey = Mid$(s, i, j - i)
            Exit Function
        End If
    Next i
    NormalizeSprintKey = s
End Function

' Returns sorted unique sprint keys from Suivi_CR col C for a given STR (col B).
' Falls back to {"1","2","3"} if no sprints found.
Public Function GetSprintsForSTR(crArr As Variant, strVal As String) As Collection
    Dim seen As Object
    Dim result As New Collection
    Dim r As Long
    Dim lStr As String
    Dim sp As String
    Dim arr() As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim k As Variant

    Set seen = CreateObject("Scripting.Dictionary")
    lStr = LCase(strVal)

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        If LCase(CStr(crArr(r, COL_B) & "")) = lStr Then
            sp = NormalizeSprintKey(crArr(r, COL_C))
            If sp <> "" And Not seen.Exists(sp) Then seen(sp) = True
        End If
    Next r

    If seen.Count = 0 Then
        result.Add "1"
        result.Add "2"
        result.Add "3"
        Set GetSprintsForSTR = result
        Exit Function
    End If

    ReDim arr(0 To seen.Count - 1)
    n = 0
    For Each k In seen.Keys
        arr(n) = CStr(k)
        n = n + 1
    Next k

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If SprintSortKey(arr(j)) < SprintSortKey(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    For i = LBound(arr) To UBound(arr)
        result.Add arr(i)
    Next i
    Set GetSprintsForSTR = result
End Function

Private Function SprintSortKey(s As String) As Double
    If IsNumeric(s) Then
        SprintSortKey = CDbl(s)
    Else
        SprintSortKey = 1E+30
    End If
End Function

' Maps each sprint key to its template row ranges: key -> Collection of Array(startRow, endRow).
' Segment 1 = ADL1, segment 2 = SwDS.
Public Function BuildSprintRangeMap(ws As Worksheet) As Object
    Dim map As Object
    Dim r As Long
    Dim startR As Long
    Dim curKey As String
    Dim nk As String

    Set map = CreateObject("Scripting.Dictionary")
    startR = TMP_FIRST_ROW
    curKey = NormalizeSprintKey(ws.Cells(TMP_FIRST_ROW, COL_D).Value)

    For r = TMP_FIRST_ROW + 1 To TMP_LAST_ROW
        nk = NormalizeSprintKey(ws.Cells(r, COL_D).Value)
        If nk <> curKey Then
            SprintMapAddRange map, curKey, startR, r - 1
            startR = r
            curKey = nk
        End If
    Next r
    SprintMapAddRange map, curKey, startR, TMP_LAST_ROW

    Set BuildSprintRangeMap = map
End Function

Private Sub SprintMapAddRange(map As Object, key As String, startR As Long, endR As Long)
    Dim col As Collection
    If key = "" Then Exit Sub
    If map.Exists(key) Then
        Set col = map(key)
    Else
        Set col = New Collection
        map.Add key, col
    End If
    col.Add Array(startR, endR)
End Sub

' ---------------------------------------------------------
'  FORMATTING HELPERS
' ---------------------------------------------------------

' Copies yellow background for cols U-X from template rows to destination rows.
Public Sub ApplyYellowSectionUtoX(wsLiv As Worksheet, ByVal destStart As Long, ByVal destEnd As Long, _
                                 wsTmp As Worksheet, ByVal srcStart As Long, ByVal srcEnd As Long)
    Dim i As Long
    Dim n As Long
    n = destEnd - destStart + 1
    If srcEnd - srcStart + 1 <> n Then Exit Sub
    For i = 0 To n - 1
        wsLiv.Range(wsLiv.Cells(destStart + i, COL_U), wsLiv.Cells(destStart + i, COL_X)).Interior.Color = _
            wsTmp.Range(wsTmp.Cells(srcStart + i, COL_U), wsTmp.Cells(srcStart + i, COL_X)).Interior.Color
    Next i
End Sub

' Thin gray outline for ADL1 / SwDS sub-blocks.
Public Sub ApplyLightOutlineBorder(ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, _
                                  ByVal lastCol As Long)
    Dim rng As Range
    If topRow > bottomRow Then Exit Sub
    Set rng = ws.Range(ws.Cells(topRow, 1), ws.Cells(bottomRow, lastCol))
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
End Sub

' Thick black outline for a full STR block.
Public Sub ApplyHardOutlineBorder(ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, _
                                  ByVal lastCol As Long)
    Dim rng As Range
    If topRow > bottomRow Then Exit Sub
    Set rng = ws.Range(ws.Cells(topRow, 1), ws.Cells(bottomRow, lastCol))
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
End Sub

' ---------------------------------------------------------
'  COLUMN F -- COUNTIFS(Suivi_CR: B=$B, C=$D, D=$E, J<>"")
' ---------------------------------------------------------

Public Function ComputeColF(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As Long
    Dim cnt As Long
    Dim r As Long
    Dim lB As String, lD As String, lE As String

    cnt = 0
    lB = LCase(B): lD = LCase(D): lE = LCase(E)

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
           LCase(CStr(crArr(r, COL_C) & "")) = lD And _
           LCase(CStr(crArr(r, COL_D) & "")) = lE And _
           CStr(crArr(r, COL_J) & "") <> "" Then
            cnt = cnt + 1
        End If
    Next r

    ComputeColF = cnt
End Function

' ---------------------------------------------------------
'  COLUMN G -- COUNTIFS(same as F) + Suivi_CR!G = "Bloque"
' ---------------------------------------------------------

Public Function ComputeColG(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As Long
    Dim cnt As Long
    Dim r As Long
    Dim lB As String, lD As String, lE As String
    Dim lBloque As String

    cnt = 0
    lB = LCase(B): lD = LCase(D): lE = LCase(E)
    lBloque = LCase("Bloqu" & ChrW(233))

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
           LCase(CStr(crArr(r, COL_C) & "")) = lD And _
           LCase(CStr(crArr(r, COL_D) & "")) = lE And _
           CStr(crArr(r, COL_J) & "") <> "" And _
           LCase(CStr(crArr(r, COL_G) & "")) = lBloque Then
            cnt = cnt + 1
        End If
    Next r

    ComputeColG = cnt
End Function

' ---------------------------------------------------------
'  COLUMN H -- SUMIFS(PowQ!Y, B=$B, C=$E, F=$C, U=$D)
' ---------------------------------------------------------

Public Function ComputeColH(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As Variant
    Dim total As Double
    Dim found As Boolean
    Dim r As Long
    Dim lB As String, lC As String, lD As String, lE As String

    total = 0
    found = False
    lB = LCase(B): lC = LCase(C): lD = LCase(D): lE = LCase(E)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_B) & "")) = lB And _
           LCase(CStr(powqArr(r, COL_C) & "")) = lE And _
           LCase(CStr(powqArr(r, COL_F) & "")) = lC And _
           LCase(CStr(powqArr(r, COL_U) & "")) = lD Then
            If IsValidPowQValue(powqArr(r, COL_Y)) And IsNumeric(powqArr(r, COL_Y)) Then
                total = total + CDbl(powqArr(r, COL_Y))
                found = True
            End If
        End If
    Next r

    If found Then
        ComputeColH = total
    Else
        ComputeColH = ""
    End If
End Function

' ---------------------------------------------------------
'  COLUMN I -- VLOOKUP PowQ_Extract col A -> "Fin Ref"
' ---------------------------------------------------------

Public Function ComputeColI(B As String, C As String, D As String, _
                            E As String, powqArr As Variant, _
                            finRefCol As Long) As String
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    If finRefCol = 0 Then
        ComputeColI = ""
        Exit Function
    End If

    lookupKey = B & "/" & E & "/" & C & "/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, finRefCol)) Then
                ComputeColI = CStr(powqArr(r, finRefCol) & "")
            Else
                ComputeColI = ""
            End If
            Exit Function
        End If
    Next r

    ComputeColI = ""
End Function

' ---------------------------------------------------------
'  COLUMN J -- MAX(PowQ!I) where B=$B, C=$E, U=$D, F=$C
' ---------------------------------------------------------

Public Function ComputeColJ(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As String
    Dim maxVal As Double
    Dim found As Boolean
    Dim r As Long
    Dim cellVal As String
    Dim numVal As Double
    Dim lB As String, lC As String, lD As String, lE As String

    maxVal = 0
    found = False
    lB = LCase(B): lC = LCase(C): lD = LCase(D): lE = LCase(E)

    For r = 2 To UBound(powqArr, 1)
        cellVal = CStr(powqArr(r, COL_I) & "")
        If cellVal <> "" And _
           LCase(CStr(powqArr(r, COL_B) & "")) = lB And _
           LCase(CStr(powqArr(r, COL_C) & "")) = lE And _
           LCase(CStr(powqArr(r, COL_U) & "")) = lD And _
           LCase(CStr(powqArr(r, COL_F) & "")) = lC Then
            If IsNumeric(powqArr(r, COL_I)) Then
                numVal = CDbl(powqArr(r, COL_I))
                If (Not found) Or numVal > maxVal Then
                    maxVal = numVal
                    found = True
                End If
            End If
        End If
    Next r

    If found Then
        ComputeColJ = CStr(maxVal)
    Else
        ComputeColJ = ""
    End If
End Function

' ---------------------------------------------------------
'  COLUMN K -- Weighted average (IFS-based, row N+1 offset)
' ---------------------------------------------------------

Public Function ComputeColK(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As Variant
    Dim numerator As Double
    Dim denominator As Double
    Dim r As Long
    Dim zVal As Double
    Dim wVal As Double
    Dim lB As String, lD As String, lE As String

    If B = "" Then
        ComputeColK = ""
        Exit Function
    End If

    lB = LCase(B): lD = LCase(D): lE = LCase(E)

    If LCase(C) = "adl1" Then
        numerator = 0: denominator = 0
        For r = CR_FIRST_ROW To UBound(crArr, 1)
            If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
               LCase(CStr(crArr(r, COL_C) & "")) = lD And _
               LCase(CStr(crArr(r, COL_D) & "")) = lE And _
               CStr(crArr(r, COL_Z) & "") <> "" And _
               CStr(crArr(r, COL_J) & "") <> "" Then
                zVal = CDblSafe(crArr(r, COL_Z))
                wVal = CDblSafe(crArr(r, COL_J))
                numerator = numerator + wVal * zVal
                denominator = denominator + zVal
            End If
        Next r
        If denominator = 0 Then
            ComputeColK = ""
        Else
            ComputeColK = numerator / denominator
        End If

    ElseIf LCase(C) = "swds" Or LCase(C) = "reprise suite valid" Then
        numerator = 0: denominator = 0
        For r = CR_FIRST_ROW To UBound(crArr, 1)
            If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
               LCase(CStr(crArr(r, COL_C) & "")) = lD And _
               LCase(CStr(crArr(r, COL_D) & "")) = lE And _
               CStr(crArr(r, COL_Z) & "") <> "" And _
               CStr(crArr(r, COL_L) & "") <> "" Then
                zVal = CDblSafe(crArr(r, COL_Z))
                wVal = CDblSafe(crArr(r, COL_L))
                numerator = numerator + wVal * zVal
                denominator = denominator + zVal
            End If
        Next r
        If denominator = 0 Then
            ComputeColK = ""
        Else
            ComputeColK = numerator / denominator
        End If

    Else
        ComputeColK = ""
    End If
End Function

' ---------------------------------------------------------
'  COLUMN O -- VLOOKUP PowQ col A -> col I
' ---------------------------------------------------------

Public Function ComputeColO(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As String
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & "/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, COL_I)) Then
                ComputeColO = CStr(powqArr(r, COL_I) & "")
            Else
                ComputeColO = ""
            End If
            Exit Function
        End If
    Next r

    ComputeColO = ""
End Function

' ---------------------------------------------------------
'  COLUMN T -- VLOOKUP PowQ col A -> col I (with " OK")
' ---------------------------------------------------------

Public Function ComputeColT(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As String
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & " OK/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, COL_I)) Then
                ComputeColT = CStr(powqArr(r, COL_I) & "")
            Else
                ComputeColT = ""
            End If
            Exit Function
        End If
    Next r

    ComputeColT = ""
End Function

' ---------------------------------------------------------
'  ARCHIVE SUIVI_LIVRABLES
' ---------------------------------------------------------

Public Sub ArchiveSuiviLivrable()
    Dim wsLiv As Worksheet
    Dim wbNew As Workbook
    Dim dlg As Object
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim ts As String

    On Error GoTo ErrHandler

    If Not SheetExists(SH_LIV) Then
        MsgBox "Sheet """ & SH_LIV & """ not found.", vbExclamation
        Exit Sub
    End If

    Set dlg = Application.FileDialog(4)  ' msoFileDialogFolderPicker
    dlg.Title = "Select archive destination folder"
    If dlg.Show <> -1 Then
        MsgBox "Archive cancelled.", vbInformation
        Exit Sub
    End If
    folderPath = dlg.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ts = Format(Now, "DDMMYYYY_HHMMSS")
    fileName = "Suivi_Livrable_" & ts & ".xlsx"
    fullPath = folderPath & fileName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
    wsLiv.Copy
    Set wbNew = ActiveWorkbook

    Dim shp As Shape
    For Each shp In wbNew.Sheets(1).Shapes
        shp.Delete
    Next shp

    wbNew.SaveAs fileName:=fullPath, _
                  FileFormat:=xlOpenXMLWorkbook, _
                  CreateBackup:=False
    wbNew.Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & lastRow).Delete Shift:=xlUp
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Archive saved & sheet reset:" & vbCrLf & fullPath, vbInformation

    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Archive failed: " & Err.Description, vbCritical
End Sub

Private Function SheetExists(shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
