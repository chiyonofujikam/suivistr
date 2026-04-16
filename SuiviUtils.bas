Option Explicit

Private m_SharedFolder As String

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Gets and caches the shared folder selected by the user.
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

' Validates that required workbook sheets exist.
Public Sub ValidateRequiredSheets()
    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim missing As String
    Dim found As Boolean
    Dim sheetList As String

    names = Array(SH_CR, SH_LIV, SH_EXTRACT, SH_VHST)
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

' Checks whether a file exists on disk.
Public Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function

' Waits until lock cell is cleared by another running update.
Public Sub WaitWhileLocked(wsCR As Worksheet, ByVal lockCell As String)
    Dim lockInfo As String
    Dim lockUser As String
    Dim lockStartedAt As String
    Dim lockDate As Date
    Dim lockAgeMinutes As Double
    Dim resp As VbMsgBoxResult
    Const STALE_LOCK_MINUTES As Double = 30#

    lockInfo = CStr(wsCR.Range(lockCell).Value & "")
    ParseLockInfo lockInfo, lockUser, lockStartedAt

    If TryParseLockDateTime(lockStartedAt, lockDate) Then
        lockAgeMinutes = DateDiff("s", lockDate, Now) / 60#
        If lockAgeMinutes >= STALE_LOCK_MINUTES Then
            resp = MsgBox("Une mise a jour est indiquee en cours depuis " & Format$(lockDate, "dd/mm/yyyy hh:nn:ss") & "." & vbCrLf & vbCrLf & _
                          "Utilisateur : " & lockUser & vbCrLf & _
                          "Duree du verrou : " & Format$(lockAgeMinutes, "0") & " minute(s)" & vbCrLf & vbCrLf & _
                          "Le verrou semble ancien. Voulez-vous le supprimer maintenant ?", _
                          vbYesNo + vbExclamation, "Verrou potentiellement bloque")
            If resp = vbYes Then
                On Error Resume Next
                wsCR.Unprotect Password:="suivi_update"
                wsCR.Range(lockCell).ClearContents
                On Error GoTo 0
                Exit Sub
            End If
        End If
    End If

    MsgBox "Une mise a jour est deja en cours." & vbCrLf & vbCrLf & _
           "Utilisateur : " & lockUser & vbCrLf & _
           "Demarree a : " & lockStartedAt & vbCrLf & vbCrLf & _
           "Veuillez recliquer sur ""Mise a jour"" une fois le traitement termine.", _
           vbInformation, "Mise a jour Suivi"
    Err.Raise vbObjectError + 1002, "WaitWhileLocked", "Update already running. Retry later."
End Sub

' Parses lock datetime formatted as "YYYY-MM-DD HH:NN:SS".
Private Function TryParseLockDateTime(ByVal s As String, ByRef parsed As Date) As Boolean
    Dim y As Integer
    Dim m As Integer
    Dim d As Integer
    Dim hh As Integer
    Dim nn As Integer
    Dim ss As Integer

    On Error GoTo ParseFail
    If Len(Trim$(s)) < 19 Then GoTo ParseFail

    y = CInt(Mid$(s, 1, 4))
    m = CInt(Mid$(s, 6, 2))
    d = CInt(Mid$(s, 9, 2))
    hh = CInt(Mid$(s, 12, 2))
    nn = CInt(Mid$(s, 15, 2))
    ss = CInt(Mid$(s, 18, 2))

    parsed = DateSerial(y, m, d) + TimeSerial(hh, nn, ss)
    TryParseLockDateTime = True
    Exit Function

ParseFail:
    TryParseLockDateTime = False
End Function

' Parses lock metadata text into user and timestamp.
Private Sub ParseLockInfo(lockInfo As String, ByRef lockUser As String, ByRef lockStartedAt As String)
    Dim s As String
    Dim pBy As Long
    Dim pAt As Long

    lockUser = "Inconnu"
    lockStartedAt = "Inconnu"

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

' Reads a text file into a string.
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

' Overwrites a text file with content.
Public Sub WriteTextFile(path As String, content As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open path For Output As #fNum
    Print #fNum, content
    Close #fNum
End Sub

' Appends a line to a text file.
Public Sub AppendTextFile(path As String, content As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open path For Append As #fNum
    Print #fNum, content
    Close #fNum
End Sub

' Converts column number to Excel letter.
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

' Normalizes a value to stable string form for comparisons.
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

' Validates PowQ cell values (non-empty, non-NaN).
Private Function IsValidPowQValue(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        IsValidPowQValue = False
    ElseIf LCase(CStr(v & "")) = "nan" Or CStr(v & "") = "" Then
        IsValidPowQValue = False
    Else
        IsValidPowQValue = True
    End If
End Function

' Converts value to Double or returns 0.
Private Function CDblSafe(v As Variant) As Double
    If IsNumeric(v) Then
        CDblSafe = CDbl(v)
    Else
        CDblSafe = 0
    End If
End Function

' Loads used worksheet range into a 2D array.
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

' Builds unique snapshot key from CR row.
Public Function SnapshotRowKey(crArr As Variant, r As Long) As String
    SnapshotRowKey = CStr(crArr(r, COL_B) & "") & "|" & _
                     CStr(crArr(r, COL_C) & "") & "|" & _
                     CStr(crArr(r, COL_D) & "")
End Function

' Serializes CR tracking data to JSON.
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

' Parses snapshot JSON into a row-key dictionary.
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

' Finds "fin ref" column index in PowQ extract.
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

' Gets last used data row for a key column.
Public Function GetLastDataRow(ws As Worksheet, keyCol As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < LIV_FIRST_ROW Then lastRow = 0
    GetLastDataRow = lastRow
End Function

' Finds first Suivi_Livrables row for an STR.
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

' Finds all Suivi_Livrables rows for an STR.
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

' Normalizes sprint value to a comparable key.
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

' Returns sorted sprint keys for a given STR.
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

' Returns target sprint keys for an STR using VHST max sprint when available.
Public Function GetTargetSprintsForSTR(crArr As Variant, strVal As String, _
                                       maxSprintMap As Object) As Collection
    Dim maxSprint As String
    Dim maxVal As Long
    Dim i As Long
    Dim result As New Collection

    If Not maxSprintMap Is Nothing Then
        If maxSprintMap.Exists(Trim$(CStr(strVal & ""))) Then
            maxSprint = NormalizeSprintKey(maxSprintMap(Trim$(CStr(strVal & ""))))
            If IsNumeric(maxSprint) Then
                maxVal = CLng(CDbl(maxSprint))
                If maxVal > 0 Then
                    For i = 1 To maxVal
                        result.Add CStr(i)
                    Next i
                    Set GetTargetSprintsForSTR = result
                    Exit Function
                End If
            End If
        End If
    End If

    Set GetTargetSprintsForSTR = GetSprintsForSTR(crArr, strVal)
End Function

' Builds Max_Sprint map from VHST sheet data.
Public Function BuildMaxSprintMapVHST(vhstArr As Variant) As Object
    Dim dict As Object
    Dim r As Long
    Dim k As String
    Dim sp As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If IsEmpty(vhstArr) Then
        Set BuildMaxSprintMapVHST = dict
        Exit Function
    End If
    If UBound(vhstArr, 1) < 2 Then
        Set BuildMaxSprintMapVHST = dict
        Exit Function
    End If

    For r = 2 To UBound(vhstArr, 1)
        k = Trim$(CStr(vhstArr(r, 1) & ""))
        If k <> "" Then
            sp = NormalizeSprintKey(vhstArr(r, 2))
            If sp <> "" Then dict(k) = sp
        End If
    Next r

    Set BuildMaxSprintMapVHST = dict
End Function

' Builds unique STR set from VHST sheet (column A).
Public Function BuildSTRMapVHST(vhstArr As Variant) As Object
    Dim dict As Object
    Dim r As Long
    Dim k As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If IsEmpty(vhstArr) Then
        Set BuildSTRMapVHST = dict
        Exit Function
    End If
    If UBound(vhstArr, 1) < 2 Then
        Set BuildSTRMapVHST = dict
        Exit Function
    End If

    For r = 2 To UBound(vhstArr, 1)
        k = Trim$(CStr(vhstArr(r, 1) & ""))
        If k <> "" Then
            If Not dict.Exists(k) Then dict(k) = True
        End If
    Next r

    Set BuildSTRMapVHST = dict
End Function

' Builds unique global Fonction list from VHST sheet (column F).
Public Function BuildFonctionsFromVHST(vhstArr As Variant) As Collection
    Dim result As New Collection
    Dim seen As Object
    Dim r As Long
    Dim fn As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    If Not IsEmpty(vhstArr) Then
        If UBound(vhstArr, 1) >= 2 And UBound(vhstArr, 2) >= COL_F Then
            For r = 2 To UBound(vhstArr, 1)
                fn = Trim$(CStr(vhstArr(r, COL_F) & ""))
                If fn <> "" Then
                    If Not seen.Exists(fn) Then
                        seen(fn) = True
                        result.Add fn
                    End If
                End If
            Next r
        End If
    End If

    Set BuildFonctionsFromVHST = result
End Function


' Builds actual max sprint map from CR data.
Public Function BuildActualMaxSprintMapCR(crArr As Variant) As Object
    Dim dict As Object
    Dim r As Long
    Dim k As String
    Dim sp As String
    Dim n As Double
    Dim cur As Double

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For r = CR_FIRST_ROW To UBound(crArr, 1)
        k = Trim$(CStr(crArr(r, COL_B) & ""))
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

' Compares CR vs VHST max sprints and offers sync.
Public Function CheckAndOfferUpdateVHSTMaxSprints(wsVHST As Worksheet, crArr As Variant, vhstArr As Variant) As Boolean
    Dim actualMap As Object
    Dim vhstMap As Object
    Dim k As Variant
    Dim a As Double, v As Double
    Dim msg As String
    Dim resp As VbMsgBoxResult
    Dim oneUpdate As Object
    Dim isMissingInVHST As Boolean

    Set actualMap = BuildActualMaxSprintMapCR(crArr)
    Set vhstMap = BuildMaxSprintMapVHST(vhstArr)

    For Each k In actualMap.Keys
        a = CDbl(actualMap(k))
        isMissingInVHST = Not vhstMap.Exists(k)
        If Not isMissingInVHST Then
            If IsNumeric(vhstMap(k)) Then
                v = CDbl(vhstMap(k))
            Else
                v = 0
            End If
        Else
            v = 0
        End If

        If a > v Then
            If isMissingInVHST Then
                msg = "La STR suivante existe dans Suivi_CR mais est absente de " & SH_VHST & " :" & vbCrLf & vbCrLf & _
                      "STR : " & CStr(k) & vbCrLf & _
                      "Suivi_CR (Sprint max) : " & CStr(CLng(a)) & vbCrLf & vbCrLf & _
                      "Voulez-vous l'ajouter dans " & SH_VHST & " avec ce sprint max ?"
            Else
                msg = "La STR suivante a un sprint dans Suivi_CR superieur a " & SH_VHST & " :" & vbCrLf & vbCrLf & _
                      "STR : " & CStr(k) & vbCrLf & _
                      SH_VHST & " (Max_Sprint) : " & CStr(CLng(v)) & vbCrLf & _
                      "Suivi_CR (Sprint max) : " & CStr(CLng(a)) & vbCrLf & vbCrLf & _
                      "Voulez-vous mettre a jour " & SH_VHST & " pour cette STR ?"
            End If

            resp = MsgBox(msg, vbYesNoCancel + vbExclamation, "Synchronisation " & SH_VHST)
            If resp = vbYes Then
                Set oneUpdate = CreateObject("Scripting.Dictionary")
                oneUpdate(CStr(k)) = CStr(CLng(a))
                ApplyVHSTMaxSprintUpdates wsVHST, oneUpdate
            ElseIf resp = vbCancel Then
                Exit For
            End If
        End If
    Next k

    CheckAndOfferUpdateVHSTMaxSprints = True
End Function

' Applies max sprint updates to VHST sheet.
Private Sub ApplyVHSTMaxSprintUpdates(wsVHST As Worksheet, updates As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Dim found As Boolean
    Dim u As Variant
    Dim targetRow As Long

    lastRow = wsVHST.Cells(wsVHST.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For Each u In updates.Keys
        k = CStr(u)
        found = False

        For r = 2 To lastRow
            If StrComp(Trim$(CStr(wsVHST.Cells(r, 1).Value & "")), k, vbTextCompare) = 0 Then
                wsVHST.Cells(r, 2).Value = updates(u)
                found = True
                Exit For
            End If
        Next r

        If Not found Then
            targetRow = 0
            For r = 2 To lastRow
                If Trim$(CStr(wsVHST.Cells(r, 1).Value & "")) = "" Then
                    targetRow = r
                    Exit For
                End If
            Next r

            If targetRow = 0 Then
                lastRow = lastRow + 1
                targetRow = lastRow
            End If

            wsVHST.Cells(targetRow, 1).Value = u
            wsVHST.Cells(targetRow, 2).Value = updates(u)
        End If
    Next u
End Sub

' Returns numeric sort key for sprint ordering.
Private Function SprintSortKey(s As String) As Double
    If IsNumeric(s) Then
        SprintSortKey = CDbl(s)
    Else
        SprintSortKey = 1E+30
    End If
End Function

' Builds template row ranges by sprint key.
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

' Adds one row range entry into sprint range map.
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

' Maps Suivi U:X headers to UVR source columns.
Public Function BuildUVRColumnMap(wsLiv As Worksheet, uvrArr As Variant) As Object
    Dim dict As Object
    Dim colIdx As Long
    Dim headerName As String
    Dim uc As Long

    Set dict = CreateObject("Scripting.Dictionary")

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

' Computes one UVR value lookup for a target row.
Public Function ComputeUVRCell(B As String, C As String, E As String, _
                               uvrArr As Variant, ByVal uvrColIdx As Long, _
                               Optional ByVal destColIdx As Long = 0) As Variant
    Dim lookupKey As String
    Dim r As Long
    Dim v As Variant

    lookupKey = LCase(B & " " & E & " " & C)

    For r = 2 To UBound(uvrArr, 1)
        If LCase(CStr(uvrArr(r, 1) & "")) = lookupKey Then
            v = uvrArr(r, uvrColIdx)
            ComputeUVRCell = NormalizeUVRCellByDestCol(v, destColIdx)
            Exit Function
        End If
    Next r

    ComputeUVRCell = NormalizeUVRCellByDestCol(Empty, destColIdx)
End Function

' Normalizes UVR value by destination column expected type.
Private Function NormalizeUVRCellByDestCol(v As Variant, ByVal destColIdx As Long) As Variant
    Select Case destColIdx
        Case COL_U, COL_X
            ' U and X are date fields: keep blank when missing/0.
            If IsValidPowQValue(v) Then
                If IsDate(v) Then
                    NormalizeUVRCellByDestCol = CDate(v)
                ElseIf IsNumeric(v) Then
                    If CDbl(v) = 0 Then
                        NormalizeUVRCellByDestCol = ""
                    Else
                        NormalizeUVRCellByDestCol = CDate(CDbl(v))
                    End If
                Else
                    NormalizeUVRCellByDestCol = ""
                End If
            Else
                NormalizeUVRCellByDestCol = ""
            End If

        Case COL_V
            ' V is numeric.
            If IsValidPowQValue(v) And IsNumeric(v) Then
                NormalizeUVRCellByDestCol = CDbl(v)
            Else
                NormalizeUVRCellByDestCol = 0
            End If

        Case COL_W
            ' W is text.
            If IsValidPowQValue(v) Then
                NormalizeUVRCellByDestCol = CStr(v)
            Else
                NormalizeUVRCellByDestCol = ""
            End If

        Case Else
            If IsValidPowQValue(v) Then
                If IsDate(v) Then
                    NormalizeUVRCellByDestCol = v
                ElseIf IsNumeric(v) Then
                    NormalizeUVRCellByDestCol = CDbl(v)
                Else
                    NormalizeUVRCellByDestCol = v
                End If
            Else
                NormalizeUVRCellByDestCol = ""
            End If
    End Select
End Function

' Writes U:X values for yellow-highlighted rows.
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
                wsLiv.Cells(rr, colIdx).Value = ComputeUVRCell(bv, cv, ev, uvrArr, CLng(uvrColMap(colIdx)), colIdx)
            End If
        Next colIdx
    Next rr
End Sub

' Computes Suivi column A key.
Public Function ComputeColA(B As String, C As String, D As String, E As String) As String
    If Trim$(B) = "" Then
        ComputeColA = ""
    Else
        ComputeColA = B & E & D & C
    End If
End Function

' Computes count metric for column F.
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

' Computes blocked count metric for column G.
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

' Computes aggregated effort for column H.
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

' Computes Fin Ref value for column I.
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

' Computes reprise reference date for column M.
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

' Computes latest relevant date for column J.
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

' Computes average metric for column K.
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

' Computes UVR date for column O.
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

' Computes UVR OK date for column T.
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

' Checks whether a sheet exists in current workbook.
Public Function SheetExists(shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
