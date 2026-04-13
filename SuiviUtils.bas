Option Explicit

' Module: SuiviUtils
' Purpose: Shared helpers used by UpdateSuiviLivrable.
' Inputs:
' - Workbook sheets named by constants in Globals.bas (e.g. SH_CR, SH_LIV, SH_EXTRACT, SH_UVR, SH_VHST, SH_TMP)
' - Arrays loaded via LoadSheetData (1-based 2D Variant arrays)

Private m_SharedFolder As String
Private m_ArchiveRunning As Boolean

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function SHARED_FOLDER_PATH() As String
    Dim dlg As Object
    Dim p As String

    If m_SharedFolder <> "" Then
        SHARED_FOLDER_PATH = m_SharedFolder
        Exit Function
    End If

    Set dlg = Application.FileDialog(4)
    With dlg
        .Title = "Select the shared folder for Suivi files (config\LOCK.txt, config\status.json)"
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

' Shows a waiting window while Suivi_CR!I1 indicates update is running.
' wsCR: Suivi_CR worksheet; lockCell: the lock cell (e.g. "I1").
Public Sub WaitWhileLocked(wsCR As Worksheet, ByVal lockCell As String)
    Dim started As Date
    Dim lockInfo As String
    Dim lockUser As String
    Dim lockStartedAt As String

    started = Now
    lockInfo = CStr(wsCR.Range(lockCell).Value & "")
    ParseLockInfo lockInfo, lockUser, lockStartedAt

    MsgBox "Update is already in progress." & vbCrLf & vbCrLf & _
           "User: " & lockUser & vbCrLf & _
           "Started at: " & lockStartedAt & vbCrLf & vbCrLf & _
           "This will wait until the update finishes.", _
           vbInformation, "Suivi Update"

    Do While Trim$(CStr(wsCR.Range(lockCell).Value & "")) <> ""
        Application.StatusBar = "Suivi Update: waiting for lock release... (" & Format$(Now - started, "hh:nn:ss") & ")"
        DoEvents
        Sleep 1000
    Loop
    Application.StatusBar = False
End Sub

Private Sub ParseLockInfo(lockInfo As String, ByRef lockUser As String, ByRef lockStartedAt As String)
    Dim s As String
    Dim pBy As Long
    Dim pAt As Long

    lockUser = "Unknown"
    lockStartedAt = "Unknown"

    s = Trim$(Replace(lockInfo, vbCr, ""))
    If s = "" Then Exit Sub

    pBy = InStr(1, s, "LOCKED by:", vbTextCompare)
    If pBy = 0 Then Exit Sub

    pAt = InStr(pBy + Len("LOCKED by:"), s, " at ", vbTextCompare)
    If pAt > 0 Then
        lockUser = Trim$(Mid$(s, pBy + Len("LOCKED by:"), pAt - (pBy + Len("LOCKED by:"))))
        lockStartedAt = Trim$(Mid$(s, pAt + Len(" at ")))
    Else
        lockUser = Trim$(Mid$(s, pBy + Len("LOCKED by:")))
    End If
End Sub

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

Public Sub AppendTextFile(path As String, content As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open path For Append As #fNum
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

Private Function IsValidPowQValue(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        IsValidPowQValue = False
    ElseIf LCase(CStr(v & "")) = "nan" Or CStr(v & "") = "" Then
        IsValidPowQValue = False
    Else
        IsValidPowQValue = True
    End If
End Function

Private Function CDblSafe(v As Variant) As Double
    If IsNumeric(v) Then
        CDblSafe = CDbl(v)
    Else
        CDblSafe = 0
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

Public Function SnapshotRowKey(crArr As Variant, r As Long) As String
    SnapshotRowKey = CStr(crArr(r, COL_B) & "") & "|" & _
                     CStr(crArr(r, COL_C) & "") & "|" & _
                     CStr(crArr(r, COL_D) & "")
End Function

Public Function SerializeSnapshotToJson(crArr As Variant) As String
    Dim coll As Collection
    Dim r As Long
    Dim c As Long
    Dim numCols As Long
    Dim strVal As String
    Dim rowDict As Object
    Dim cellsDict As Object
    Dim colLetter As String
    Dim v As Variant

    Set coll = New Collection
    numCols = UBound(crArr, 2)

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        strVal = CStr(crArr(r, COL_B) & "")
        If strVal = "" Then GoTo NextSnapRow

        Set rowDict = CreateObject("Scripting.Dictionary")
        rowDict("STR") = strVal
        rowDict("key") = SnapshotRowKey(crArr, r)
        rowDict("row") = r

        Set cellsDict = CreateObject("Scripting.Dictionary")
        ' Snapshot only B:Q (2..17) to keep status.json small/stable.
        For c = COL_B To 17
            If c > numCols Then Exit For
            colLetter = ColNumToLetter(c)
            v = crArr(r, c)
            If (colLetter = "J" Or colLetter = "L" Or colLetter = "N") Then
                If IsNumeric(v) Then
                    If CDbl(v) = 0 Then
                        cellsDict(colLetter) = ""
                    Else
                        cellsDict(colLetter) = NormalizeValue(v)
                    End If
                Else
                    cellsDict(colLetter) = NormalizeValue(v)
                End If
            Else
                cellsDict(colLetter) = NormalizeValue(v)
            End If
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
    Dim rowKey As String
    Dim cellsObj As Object
    Dim cellsDict As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set coll = JsonConverter.ParseJson(jsonStr)

    For Each item In coll
        If item.Exists("key") Then
            rowKey = CStr(item("key"))
        Else
            rowKey = CStr(item("STR"))
        End If
        Set cellsObj = item("cells")
        Set cellsDict = CreateObject("Scripting.Dictionary")

        For Each key In cellsObj.Keys
            cellsDict(CStr(key)) = NormalizeValue(cellsObj(key))
        Next key

        Set result(rowKey) = cellsDict
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

' Builds a lookup map from PowQ_EDU_CE_VHST:
'   Col 1 = Nom_STR, Col 2 = Max_Sprint
' Returns Dictionary: LCase(Nom_STR) -> normalized sprint key (e.g. "3")
Public Function BuildMaxSprintMapVHST(vhstArr As Variant) As Object
    Dim dict As Object
    Dim r As Long
    Dim k As String
    Dim sp As String

    Set dict = CreateObject("Scripting.Dictionary")

    If IsEmpty(vhstArr) Then
        Set BuildMaxSprintMapVHST = dict
        Exit Function
    End If
    If UBound(vhstArr, 1) < 2 Then
        Set BuildMaxSprintMapVHST = dict
        Exit Function
    End If

    For r = 2 To UBound(vhstArr, 1)
        k = LCase(Trim$(CStr(vhstArr(r, 1) & "")))
        If k <> "" Then
            sp = NormalizeSprintKey(vhstArr(r, 2))
            If sp <> "" Then dict(k) = sp
        End If
    Next r

    Set BuildMaxSprintMapVHST = dict
End Function

Private Function CollectionContains(col As Collection, ByVal value As String) As Boolean
    Dim v As Variant
    For Each v In col
        If CStr(v) = value Then
            CollectionContains = True
            Exit Function
        End If
    Next v
    CollectionContains = False
End Function

' Determines which sprint should receive the yellow section for a STR.
' Primary source is PowQ_EDU_CE_VHST max sprint; fallback is the last sprint
' present in Suivi_CR for that STR (and present in the template sprintMap).
Public Function GetYellowSprintKeyForSTR(strKey As String, maxSprintMap As Object, _
                                        strSprints As Collection, sprintMap As Object) As String
    Dim candidate As String
    Dim k As String
    Dim i As Long

    GetYellowSprintKeyForSTR = ""
    k = LCase(Trim$(CStr(strKey & "")))

    If Not maxSprintMap Is Nothing Then
        If maxSprintMap.Exists(k) Then
            candidate = CStr(maxSprintMap(k))
            If candidate <> "" Then
                If CollectionContains(strSprints, candidate) And sprintMap.Exists(candidate) Then
                    GetYellowSprintKeyForSTR = candidate
                    Exit Function
                End If
            End If
        End If
    End If

    For i = strSprints.Count To 1 Step -1
        candidate = CStr(strSprints(i))
        If sprintMap.Exists(candidate) Then
            GetYellowSprintKeyForSTR = candidate
            Exit Function
        End If
    Next i
End Function

' Computes actual maximum sprint per STR based on Suivi_CR.
' Returns Dictionary: LCase(STR) -> max sprint as string (e.g. "3")
Public Function BuildActualMaxSprintMapCR(crArr As Variant) As Object
    Dim dict As Object
    Dim r As Long
    Dim k As String
    Dim sp As String
    Dim n As Double
    Dim cur As Double

    Set dict = CreateObject("Scripting.Dictionary")

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        k = LCase(Trim$(CStr(crArr(r, COL_B) & "")))
        If k <> "" Then
            sp = NormalizeSprintKey(crArr(r, COL_C))
            If IsNumeric(sp) Then
                n = CDbl(sp)
                If dict.Exists(k) Then
                    cur = CDbl(dict(k))
                    If n > cur Then dict(k) = CStr(CLng(n))
                Else
                    dict(k) = CStr(CLng(n))
                End If
            End If
        End If
    Next r

    Set BuildActualMaxSprintMapCR = dict
End Function

' Checks CR max sprint > VHST max sprint. If user agrees, updates the VHST sheet (col 2),
' and adds missing STRs at the bottom. Returns True if execution should continue.
Public Function CheckAndOfferUpdateVHSTMaxSprints(wsVHST As Worksheet, crArr As Variant, vhstArr As Variant) As Boolean
    Dim actualMap As Object
    Dim vhstMap As Object
    Dim updates As Object
    Dim k As Variant
    Dim a As Double, v As Double
    Dim msg As String
    Dim shown As Long
    Dim resp As VbMsgBoxResult

    Set actualMap = BuildActualMaxSprintMapCR(crArr)
    Set vhstMap = BuildMaxSprintMapVHST(vhstArr)
    Set updates = CreateObject("Scripting.Dictionary")

    msg = "Some STR(s) have a sprint in Suivi_CR higher than the Max_Sprint in " & SH_VHST & ":" & vbCrLf & vbCrLf
    shown = 0

    For Each k In actualMap.Keys
        a = CDbl(actualMap(k))
        If vhstMap.Exists(k) Then
            If IsNumeric(vhstMap(k)) Then
                v = CDbl(vhstMap(k))
            Else
                v = 0
            End If
        Else
            v = 0
        End If

        If a > v Then
            updates(k) = CStr(CLng(a))
            If shown < 20 Then
                msg = msg & "- " & CStr(k) & " : VHST=" & CStr(CLng(v)) & " / CR=" & CStr(CLng(a)) & vbCrLf
                shown = shown + 1
            End If
        End If
    Next k

    If updates.Count = 0 Then
        CheckAndOfferUpdateVHSTMaxSprints = True
        Exit Function
    End If

    If updates.Count > shown Then
        msg = msg & vbCrLf & "(+" & (updates.Count - shown) & " more)"
    End If
    msg = msg & vbCrLf & vbCrLf & "Do you want to update " & SH_VHST & " Max_Sprint values now?"

    resp = MsgBox(msg, vbYesNo + vbExclamation, "Max Sprint mismatch")
    If resp = vbYes Then
        ApplyVHSTMaxSprintUpdates wsVHST, updates
    End If

    CheckAndOfferUpdateVHSTMaxSprints = True
End Function

Private Sub ApplyVHSTMaxSprintUpdates(wsVHST As Worksheet, updates As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Dim found As Boolean
    Dim u As Variant

    lastRow = wsVHST.Cells(wsVHST.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For Each u In updates.Keys
        k = CStr(u)
        found = False

        For r = 2 To lastRow
            If LCase(Trim$(CStr(wsVHST.Cells(r, 1).Value & ""))) = k Then
                wsVHST.Cells(r, 2).Value = updates(u)
                found = True
                Exit For
            End If
        Next r

        If Not found Then
            lastRow = lastRow + 1
            wsVHST.Cells(lastRow, 1).Value = u
            wsVHST.Cells(lastRow, 2).Value = updates(u)
        End If
    Next u
End Sub

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

' Builds a mapping from Suivi_Livrables header names (row 3, cols U-X)
' to PowQ_Suivi_UVR column indices (row 1).
' Returns a Dictionary: livColIndex -> uvrColIndex  (e.g. 21 -> 5)
Public Function BuildUVRColumnMap(wsLiv As Worksheet, uvrArr As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim colIdx As Long
    Dim headerName As String
    Dim uc As Long

    For colIdx = COL_U To COL_X
        headerName = LCase(Trim$(CStr(wsLiv.Cells(3, colIdx).Value & "")))
        If headerName <> "" Then
            For uc = 1 To UBound(uvrArr, 2)
                If LCase(Trim$(CStr(uvrArr(1, uc) & ""))) = headerName Then
                    dict(colIdx) = uc
                    Exit For
                End If
            Next uc
        End If
    Next colIdx

    Set BuildUVRColumnMap = dict
End Function

' Looks up a single UVR value. Key = B & " " & E & " " & C matched against col A.
' Returns the value from uvrColIdx, or "" if not found / zero.
Public Function ComputeUVRCell(B As String, C As String, E As String, _
                               uvrArr As Variant, ByVal uvrColIdx As Long) As Variant
    Dim lookupKey As String
    Dim r As Long
    Dim v As Variant

    lookupKey = LCase(B & " " & E & " " & C)

    For r = 2 To UBound(uvrArr, 1)
        If LCase(CStr(uvrArr(r, 1) & "")) = lookupKey Then
            v = uvrArr(r, uvrColIdx)
            If IsValidPowQValue(v) Then
                If IsDate(v) Then
                    ComputeUVRCell = v
                ElseIf IsNumeric(v) Then
                    ComputeUVRCell = CDbl(v)
                Else
                    ComputeUVRCell = v
                End If
            Else
                ComputeUVRCell = 0
            End If
            Exit Function
        End If
    Next r

    ComputeUVRCell = 0
End Function

' Writes cols U-X values for a range of rows using the UVR data.
Public Sub WriteYellowValuesUtoX(wsLiv As Worksheet, ByVal destStart As Long, ByVal destEnd As Long, _
                                 uvrArr As Variant, uvrColMap As Object, livArr As Variant)
    Dim rr As Long
    Dim colIdx As Long
    Dim bv As String, cv As String, ev As String

    For rr = destStart To destEnd
        bv = CStr(livArr(rr, COL_B) & "")
        cv = CStr(livArr(rr, COL_C) & "")
        ev = CStr(livArr(rr, COL_E) & "")

        For colIdx = COL_U To COL_X
            If uvrColMap.Exists(colIdx) Then
                wsLiv.Cells(rr, colIdx).Value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colIdx)))
            End If
        Next colIdx
    Next rr
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

' Clears all borders in Suivi_Livrables (from LIV_FIRST_ROW) then rebuilds
' light borders for ADL1 / SwDS and hard borders per STR block.
Public Sub RebuildSuiviLivrablesBorders(wsLiv As Worksheet, wsTmp As Worksheet, sprintMap As Object, ByVal lastCol As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim curStr As String
    Dim nextStr As String
    Dim swdsMarker As String
    Dim swdsStartRow As Long
    Dim k As Variant
    Dim rangesCol As Collection
    Dim pair As Variant

    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow < LIV_FIRST_ROW Then Exit Sub

    ' Determine SwDS marker from template: col C of first SwDS row for any sprint.
    swdsMarker = ""
    If Not sprintMap Is Nothing Then
        For Each k In sprintMap.Keys
            Set rangesCol = sprintMap(CStr(k))
            If rangesCol.Count >= 2 Then
                pair = rangesCol(2)
                swdsMarker = CStr(wsTmp.Cells(CLng(pair(0)), COL_C).Value & "")
                Exit For
            End If
        Next k
    End If

    ' Clear all existing borders.
    With wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 1), wsLiv.Cells(lastRow, lastCol)).Borders
        .LineStyle = xlNone
    End With

    blockStart = LIV_FIRST_ROW
    Do While blockStart <= lastRow
        curStr = Trim$(CStr(wsLiv.Cells(blockStart, COL_B).Value & ""))
        If curStr = "" Then
            blockStart = blockStart + 1
            GoTo NextBlock
        End If

        blockEnd = blockStart
        Do While blockEnd < lastRow
            nextStr = Trim$(CStr(wsLiv.Cells(blockEnd + 1, COL_B).Value & ""))
            If nextStr <> curStr Then Exit Do
            blockEnd = blockEnd + 1
        Loop

        ' Find SwDS start row within this STR block.
        swdsStartRow = blockEnd + 1
        If swdsMarker <> "" Then
            For r = blockStart To blockEnd
                If CStr(wsLiv.Cells(r, COL_C).Value & "") = swdsMarker Then
                    swdsStartRow = r
                    Exit For
                End If
            Next r
        End If

        If swdsStartRow > blockStart Then
            ApplyLightOutlineBorder wsLiv, blockStart, swdsStartRow - 1, lastCol
        End If
        If swdsStartRow <= blockEnd Then
            ApplyLightOutlineBorder wsLiv, swdsStartRow, blockEnd, lastCol
        End If
        ApplyHardOutlineBorder wsLiv, blockStart, blockEnd, lastCol

        blockStart = blockEnd + 1
NextBlock:
    Loop
End Sub

' Compute column A for Suivi_Livrables.
' Inputs: B,C,D,E are the row values from columns B,C,D,E.
' Output: if B is empty => "" else B & E & D & C
Public Function ComputeColA(B As String, C As String, D As String, E As String) As String
    If Trim$(B) = "" Then
        ComputeColA = ""
    Else
        ComputeColA = B & E & D & C
    End If
End Function

' ---------------------------------------------------------
'  COLUMN F -- COUNTIFS(Suivi_CR: B=$B, C=$D, D=$E, J<>"")
' ---------------------------------------------------------

Public Function ComputeColF(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As Long
    Dim cnt As Long
    Dim r As Long
    Dim lB As String, lD As String, lE As String
    Dim valueCol As Long

    cnt = 0
    lB = LCase(B): lD = LCase(D): lE = LCase(E)
    valueCol = COL_J
    If LCase(Trim$(C)) = "swds" Then valueCol = COL_L

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
           LCase(CStr(crArr(r, COL_C) & "")) = lD And _
           LCase(CStr(crArr(r, COL_D) & "")) = lE And _
           CStr(crArr(r, valueCol) & "") <> "" Then
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
    Dim valueCol As Long

    cnt = 0
    lB = LCase(B): lD = LCase(D): lE = LCase(E)
    lBloque = LCase("Bloqu" & ChrW(233))
    valueCol = COL_J
    If LCase(Trim$(C)) = "swds" Then valueCol = COL_L

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
           LCase(CStr(crArr(r, COL_C) & "")) = lD And _
           LCase(CStr(crArr(r, COL_D) & "")) = lE And _
           CStr(crArr(r, valueCol) & "") <> "" And _
           (LCase(CStr(crArr(r, COL_G) & "")) = lBloque Or LCase(CStr(crArr(r, COL_G) & "")) = "bloque") Then
            cnt = cnt + 1
        End If
    Next r

    ComputeColG = cnt
End Function

' ---------------------------------------------------------
'  COLUMN H -- SUMIFS(PowQ!Y, B=$B, C=$E, F=$C, U=$D)
' ---------------------------------------------------------

Public Function ComputeColH(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As Double
    Dim total As Double
    Dim r As Long
    Dim lB As String, lC As String, lD As String, lE As String

    total = 0
    lB = LCase(B): lC = LCase(C): lD = LCase(D): lE = LCase(E)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_B) & "")) = lB And _
           LCase(CStr(powqArr(r, COL_C) & "")) = lE And _
           LCase(CStr(powqArr(r, COL_F) & "")) = lC And _
           LCase(CStr(powqArr(r, COL_U) & "")) = lD Then
            If IsValidPowQValue(powqArr(r, COL_Y)) And IsNumeric(powqArr(r, COL_Y)) Then
                total = total + CDbl(powqArr(r, COL_Y))
            End If
        End If
    Next r

    ComputeColH = total
End Function

' ---------------------------------------------------------
'  COLUMN I -- VLOOKUP PowQ_Extract col A -> "Fin Ref"
' ---------------------------------------------------------

Public Function ComputeColI(B As String, C As String, D As String, _
                            E As String, powqArr As Variant, _
                            finRefCol As Long) As Variant
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
                ComputeColI = powqArr(r, finRefCol)
            Else
                ComputeColI = ""
            End If
            Exit Function
        End If
    Next r

    ComputeColI = ""
End Function

' Column M -- IFERROR(VLOOKUP(B&"/"&E&"/Reprise suite valid/Sprint "&D; PowQ_Extract A:I; 9; 0), "")
Public Function ComputeColM(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As Variant
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/Reprise suite valid/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, COL_I)) Then
                ComputeColM = powqArr(r, COL_I)
            Else
                ComputeColM = ""
            End If
            Exit Function
        End If
    Next r

    ComputeColM = ""
End Function

' ---------------------------------------------------------
'  COLUMN J -- MAX(PowQ_Extract I) where B=$B, C=$E, U=$D, F=$C
'  ADL1:  I <> "" (numeric values only contribute to MAX; blank => "")
'  SwDS:  I <> "" and I <> 0 (same matches; excludes zero)
' ---------------------------------------------------------

Public Function ComputeColJ(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As Variant
    Dim maxVal As Double
    Dim found As Boolean
    Dim r As Long
    Dim numVal As Double
    Dim lB As String, lC As String, lD As String, lE As String
    Dim isSwds As Boolean
    Dim v As Variant

    maxVal = 0
    found = False
    lB = LCase(B): lC = LCase(C): lD = LCase(D): lE = LCase(E)
    isSwds = (LCase(Trim$(C)) = "swds")

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_B) & "")) = lB And _
           LCase(CStr(powqArr(r, COL_C) & "")) = lE And _
           LCase(CStr(powqArr(r, COL_U) & "")) = lD And _
           LCase(CStr(powqArr(r, COL_F) & "")) = lC Then

            v = powqArr(r, COL_I)
            If Not IsEmpty(v) And Not IsError(v) Then
                If IsDate(v) Or IsNumeric(v) Then
                    numVal = CDbl(v)
                    If isSwds Then
                        If numVal <> 0 Then
                            If (Not found) Or numVal > maxVal Then
                                maxVal = numVal
                                found = True
                            End If
                        End If
                    Else
                        If (Not found) Or numVal > maxVal Then
                            maxVal = numVal
                            found = True
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If found Then
        ComputeColJ = CDate(maxVal)
    Else
        ComputeColJ = ""
    End If
End Function

' ---------------------------------------------------------
'  COLUMN K -- Average of Suivi_CR!J (ADL1) or L (SwDS) for matching B/D/E
' ---------------------------------------------------------

Public Function ComputeColK(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As Double
    Dim total As Double
    Dim cnt As Long
    Dim r As Long
    Dim lB As String, lD As String, lE As String

    If B = "" Then
        ComputeColK = 0
        Exit Function
    End If

    lB = LCase(B): lD = LCase(D): lE = LCase(E)

    If LCase(C) = "adl1" Then
        total = 0: cnt = 0
        For r = CR_FIRST_ROW To UBound(crArr, 1)
            If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
               LCase(CStr(crArr(r, COL_C) & "")) = lD And _
               LCase(CStr(crArr(r, COL_D) & "")) = lE And _
               CStr(crArr(r, COL_J) & "") <> "" Then
                total = total + CDblSafe(crArr(r, COL_J))
                cnt = cnt + 1
            End If
        Next r

    ElseIf LCase(C) = "swds" Or LCase(C) = "reprise suite valid" Then
        total = 0: cnt = 0
        For r = CR_FIRST_ROW To UBound(crArr, 1)
            If LCase(CStr(crArr(r, COL_B) & "")) = lB And _
               LCase(CStr(crArr(r, COL_C) & "")) = lD And _
               LCase(CStr(crArr(r, COL_D) & "")) = lE And _
               CStr(crArr(r, COL_L) & "") <> "" Then
                total = total + CDblSafe(crArr(r, COL_L))
                cnt = cnt + 1
            End If
        Next r
    End If

    If cnt > 0 Then
        ComputeColK = total / cnt
    Else
        ComputeColK = 0
    End If
End Function

' ---------------------------------------------------------
'  COLUMN O -- VLOOKUP PowQ col A -> col I
' ---------------------------------------------------------

Public Function ComputeColO(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As Variant
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & "/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, COL_I)) Then
                ComputeColO = powqArr(r, COL_I)
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
                            E As String, powqArr As Variant) As Variant
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & " OK/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            If IsValidPowQValue(powqArr(r, COL_I)) Then
                ComputeColT = powqArr(r, COL_I)
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
    Dim wsNew As Worksheet
    Dim srcRng As Range
    Dim dstRng As Range
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim ts As String
    Dim dayFolder As String
    Dim resp As VbMsgBoxResult
    Dim shp As Shape
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long
    Dim r As Long
    Dim cfg As String
    Dim errLine As String

    If m_ArchiveRunning Then Exit Sub
    m_ArchiveRunning = True

    On Error GoTo ErrHandler

    If Not SheetExists(SH_LIV) Then
        MsgBox "Sheet """ & SH_LIV & """ not found.", vbExclamation
        Exit Sub
    End If

    folderPath = SHARED_FOLDER_PATH & "Archived\"
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath

    dayFolder = folderPath & Format$(Date, "DDMMYYYY") & "\"
    If Dir$(dayFolder, vbDirectory) = "" Then MkDir dayFolder

    ts = Format(Now, "DDMMYYYY_HHMMSS")
    fileName = "Suivi_Livrable_" & ts & ".xlsx"
    fullPath = dayFolder & fileName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)

    ' Remove any active AutoFilter so the archive covers ALL rows.
    If wsLiv.AutoFilterMode Then wsLiv.AutoFilterMode = False

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = wsLiv.Name

    lastRow = wsLiv.UsedRange.Row + wsLiv.UsedRange.Rows.Count - 1
    lastCol = wsLiv.UsedRange.Column + wsLiv.UsedRange.Columns.Count - 1
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    Set srcRng = wsLiv.Range(wsLiv.Cells(1, 1), wsLiv.Cells(lastRow, lastCol))
    Set dstRng = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lastRow, lastCol))

    srcRng.Copy
    dstRng.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    dstRng.Value = srcRng.Value

    For c = 1 To lastCol
        wsNew.Columns(c).ColumnWidth = wsLiv.Columns(c).ColumnWidth
    Next c
    For r = 1 To lastRow
        wsNew.Rows(r).RowHeight = wsLiv.Rows(r).RowHeight
    Next r

    For Each shp In wsNew.Shapes
        shp.Delete
    Next shp

    wbNew.SaveAs fileName:=fullPath, _
                  FileFormat:=xlOpenXMLWorkbook, _
                  CreateBackup:=False
    wbNew.Close SaveChanges:=False

    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & lastRow).Delete Shift:=xlUp
    End If

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    m_ArchiveRunning = False

    resp = MsgBox("Archive saved & sheet reset." & vbCrLf & vbCrLf & _
                  "Open the archived file now?" & vbCrLf & fullPath, _
                  vbYesNo + vbInformation, "Archive")
    If resp = vbYes Then
        ThisWorkbook.FollowHyperlink fullPath
    End If

    Exit Sub

ErrHandler:
    cfg = SHARED_FOLDER_PATH & "config\"
    If Dir$(cfg, vbDirectory) = "" Then MkDir cfg
    errLine = Format$(Now, "YYYY-MM-DD HH:NN:SS") & _
              " | user=" & Environ$("USERNAME") & _
              " | proc=ArchiveSuiviLivrable" & _
              " | err=" & Err.Number & _
              " | " & Err.Description
    On Error Resume Next
    AppendTextFile cfg & "error_logs.txt", errLine
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    m_ArchiveRunning = False
    MsgBox "Archive failed: " & Err.Description, vbCritical
End Sub

Public Function SheetExists(shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
