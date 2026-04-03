Option Explicit

' DEPENDENCY: JsonConverter.bas must be imported into this VBA project.
' Download from https://github.com/VBA-tools/VBA-JSON
' Also requires: Tools > References > Microsoft Scripting Runtime (for Dictionary)

' ---------------------------------------------------------
'  SHEET & DATA CONSTANTS
' ---------------------------------------------------------

Public Const SH_CR         As String = "Suivi_CR"
Public Const SH_LIV        As String = "Suivi_Livrables"
Public Const SH_EXTRACT    As String = "PowQ_Extract"
Public Const SH_UVR        As String = "PowQ_Suivi_UVR"
Public Const SH_TMP        As String = "Suivi_Livrables_Tmp"

Public Const CR_FIRST_ROW  As Long = 3
Public Const LIV_FIRST_ROW As Long = 4
Public Const CR_MAX_ROW    As Long = 9976

' TODO-2: Template range in Suivi_Livrables_Tmp -- rows 4-33, cols B-E.
'         Currently only row 4 is used as the formatting source.
'         Adjust TMP_FIRST_ROW / TMP_LAST_ROW if a block of rows
'         should be copied per STR.
Public Const TMP_FIRST_ROW As Long = 4
Public Const TMP_LAST_ROW  As Long = 33

' ---------------------------------------------------------
'  COLUMN INDEX CONSTANTS  (1-based, matching Excel columns)
' ---------------------------------------------------------

Public Const COL_A As Long = 1
Public Const COL_B As Long = 2
Public Const COL_C As Long = 3
Public Const COL_D As Long = 4
Public Const COL_E As Long = 5
Public Const COL_F As Long = 6
Public Const COL_G As Long = 7
Public Const COL_H As Long = 8
Public Const COL_I As Long = 9
Public Const COL_J As Long = 10
Public Const COL_K As Long = 11
Public Const COL_L As Long = 12
Public Const COL_O As Long = 15
Public Const COL_T As Long = 20
Public Const COL_U As Long = 21
Public Const COL_Y As Long = 25
Public Const COL_Z As Long = 26

' ---------------------------------------------------------
'  CONFIGURATION
' ---------------------------------------------------------

Private m_SharedFolder As String

Public Function SHARED_FOLDER_PATH() As String
    If m_SharedFolder <> "" Then
        SHARED_FOLDER_PATH = m_SharedFolder
        Exit Function
    End If

    Dim dlg As Object
    Set dlg = Application.FileDialog(4)  ' 4 = msoFileDialogFolderPicker
    With dlg
        .Title = "Select the shared folder for Suivi files (LOCK.txt, status.json)"
        .ButtonName = "Select"
        If .Show = -1 Then
            Dim p As String
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

    names = Array(SH_CR, SH_LIV, SH_EXTRACT, SH_TMP)
    missing = ""

    For i = LBound(names) To UBound(names)
        Dim found As Boolean
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
        Dim sheetList As String
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
'  COLUMN F -- COUNTIFS(Suivi_CR!B=$B, C=$D, D=$E, J<>"")
' ---------------------------------------------------------
' NB.SI.ENS / COUNTIFS on Suivi_CR.
' Parameters B/C/D/E are Suivi_Livrable column values.
' Cross-mapping: CR col B = param B, CR col C = param D, CR col D = param E.

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
            total = total + CDblSafe(powqArr(r, COL_Y))
        End If
    Next r

    ComputeColH = total
End Function

' ---------------------------------------------------------
'  COLUMN I -- VLOOKUP on PowQ_Extract col A -> "Fin Ref" column
' ---------------------------------------------------------
' Key: B & "/" & E & "/" & C & "/Sprint " & D

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
            ComputeColI = CStr(powqArr(r, finRefCol) & "")
            Exit Function
        End If
    Next r

    ComputeColI = ""
End Function

' ---------------------------------------------------------
'  COLUMN J -- MAX(PowQ!I) where B=$B, C=$E, U=$D, F=$C, I<>""
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
'  COLUMN K -- Weighted average (IFS-based)
' ---------------------------------------------------------
' IMPORTANT: B/C/D/E are from Suivi_Livrable row N+1 (offset).
' TODO-4: Verify this +1 offset with the sheet owner before deploying.
'
' Cross-mapping for CR scan: CR col B = param B, CR col C = param D,
'                            CR col D = param E.

Public Function ComputeColK(B As String, C As String, D As String, _
                            E As String, crArr As Variant) As String
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
            ComputeColK = CStr(numerator / denominator)
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
            ComputeColK = CStr(numerator / denominator)
        End If

    Else
        ComputeColK = ""
    End If
End Function

' ---------------------------------------------------------
'  COLUMN O -- VLOOKUP on PowQ_Extract col A -> col I (index 9)
' ---------------------------------------------------------
' Key: B & "/" & E & "/UVR " & C & "/Sprint " & D

Public Function ComputeColO(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As String
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & "/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            ComputeColO = CStr(powqArr(r, COL_I) & "")
            Exit Function
        End If
    Next r

    ComputeColO = ""
End Function

' ---------------------------------------------------------
'  COLUMN T -- VLOOKUP on PowQ_Extract col A -> col I (index 9)
' ---------------------------------------------------------
' Key: B & "/" & E & "/UVR " & C & " OK/Sprint " & D
' (same as col O but with " OK" suffix before "/Sprint ")

Public Function ComputeColT(B As String, C As String, D As String, _
                            E As String, powqArr As Variant) As String
    Dim lookupKey As String
    Dim lKey As String
    Dim r As Long

    lookupKey = B & "/" & E & "/UVR " & C & " OK/Sprint " & D
    lKey = LCase(lookupKey)

    For r = 2 To UBound(powqArr, 1)
        If LCase(CStr(powqArr(r, COL_A) & "")) = lKey Then
            ComputeColT = CStr(powqArr(r, COL_I) & "")
            Exit Function
        End If
    Next r

    ComputeColT = ""
End Function
