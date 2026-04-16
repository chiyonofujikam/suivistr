Option Explicit

' Returns a darker color for higher sprint indexes.
Public Function DarkenColor(ByVal baseColor As Long, ByVal sprintIndex As Long, _
                            Optional ByVal stepRatio As Double = SPRINT_DARKEN_STEP) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim factor As Double

    If sprintIndex < 1 Then sprintIndex = 1
    factor = 1# - (stepRatio * (sprintIndex - 1))
    If factor < 0# Then factor = 0#

    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256

    r = CLng(r * factor)
    g = CLng(g * factor)
    b = CLng(b * factor)

    DarkenColor = RGB(r, g, b)
End Function

' Creates block rows for one STR: ADL1 (all sprints/functions) then SWDS.
Public Function InsertGeneratedSTRBlock(wsLiv As Worksheet, ByVal startRow As Long, _
                                        ByVal strKey As String, strSprints As Collection, _
                                        fonctions As Collection, ByVal lastCol As Long, _
                                        ByVal maxSprintKey As String) As Variant
    Dim rowPtr As Long
    Dim sp As Variant
    Dim fn As Variant
    Dim sectionTop As Long
    Dim sectionBottom As Long
    Dim blockTop As Long
    Dim blockBottom As Long
    Dim sprintKey As String
    Dim sprintColorD As Long

    blockTop = startRow
    rowPtr = startRow

    sectionTop = rowPtr
    For Each sp In strSprints
        sprintKey = CStr(sp)
        sprintColorD = DarkenColor(COLOR_BASE_SPRINT, CLng(Val(sprintKey)))
        For Each fn In fonctions
            wsLiv.Cells(rowPtr, COL_B).Value = strKey
            wsLiv.Cells(rowPtr, COL_C).Value = SECTION_ADL1
            wsLiv.Cells(rowPtr, COL_D).Value = sprintKey
            wsLiv.Cells(rowPtr, COL_E).Value = CStr(fn)

            ' Keep backgrounds only where managed by the generation rules.
            wsLiv.Range(wsLiv.Cells(rowPtr, 1), wsLiv.Cells(rowPtr, lastCol)).Interior.ColorIndex = xlNone
            wsLiv.Cells(rowPtr, COL_B).Interior.Color = COLOR_B_BASE_ADL1
            wsLiv.Cells(rowPtr, COL_C).Interior.Color = COLOR_C_ADL1
            wsLiv.Cells(rowPtr, COL_D).Interior.Color = sprintColorD
            If maxSprintKey <> "" And sprintKey = maxSprintKey Then
                wsLiv.Range(wsLiv.Cells(rowPtr, COL_U), wsLiv.Cells(rowPtr, COL_X)).Interior.Color = COLOR_YELLOW_ZONE
            Else
                wsLiv.Range(wsLiv.Cells(rowPtr, COL_U), wsLiv.Cells(rowPtr, COL_X)).Interior.Color = COLOR_UX_DEFAULT
            End If
            rowPtr = rowPtr + 1
        Next fn
    Next sp
    sectionBottom = rowPtr - 1
    If sectionBottom >= sectionTop Then
        ApplyLightOutlineBorder wsLiv, sectionTop, sectionBottom, lastCol
    End If

    sectionTop = rowPtr
    For Each sp In strSprints
        sprintKey = CStr(sp)
        sprintColorD = DarkenColor(COLOR_BASE_SPRINT, CLng(Val(sprintKey)))
        For Each fn In fonctions
            wsLiv.Cells(rowPtr, COL_B).Value = strKey
            wsLiv.Cells(rowPtr, COL_C).Value = SECTION_SWDS
            wsLiv.Cells(rowPtr, COL_D).Value = sprintKey
            wsLiv.Cells(rowPtr, COL_E).Value = CStr(fn)

            ' Keep backgrounds only where managed by the generation rules.
            wsLiv.Range(wsLiv.Cells(rowPtr, 1), wsLiv.Cells(rowPtr, lastCol)).Interior.ColorIndex = xlNone
            wsLiv.Cells(rowPtr, COL_B).Interior.Color = COLOR_B_BASE_SWDS
            wsLiv.Cells(rowPtr, COL_C).Interior.Color = COLOR_C_SWDS
            wsLiv.Cells(rowPtr, COL_D).Interior.Color = sprintColorD
            If maxSprintKey <> "" And sprintKey = maxSprintKey Then
                wsLiv.Range(wsLiv.Cells(rowPtr, COL_U), wsLiv.Cells(rowPtr, COL_X)).Interior.Color = COLOR_YELLOW_ZONE
            Else
                wsLiv.Range(wsLiv.Cells(rowPtr, COL_U), wsLiv.Cells(rowPtr, COL_X)).Interior.Color = COLOR_UX_DEFAULT
            End If
            rowPtr = rowPtr + 1
        Next fn
    Next sp
    sectionBottom = rowPtr - 1
    If sectionBottom >= sectionTop Then
        ApplyLightOutlineBorder wsLiv, sectionTop, sectionBottom, lastCol
    End If

    blockBottom = rowPtr - 1
    If blockBottom >= blockTop Then
        ApplyHardOutlineBorder wsLiv, blockTop, blockBottom, lastCol
    End If

    InsertGeneratedSTRBlock = Array(blockTop, blockBottom)
End Function

' Returns row count for one generated STR block.
Public Function GeneratedBlockRowCount(strSprints As Collection, fonctions As Collection) As Long
    GeneratedBlockRowCount = 2 * CLng(strSprints.Count) * CLng(fonctions.Count)
End Function

' Applies number/date/text formats by Suivi_Livrables column rules.
Public Sub ApplySuiviLivrablesColumnFormats(wsLiv As Worksheet)
    Dim lastRow As Long
    Dim r As Range
    Dim dataRng As Range

    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow < LIV_FIRST_ROW Then Exit Sub

    Set dataRng = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_A), wsLiv.Cells(lastRow, COL_Y))
    dataRng.Font.Name = "Aptos Narrow"
    dataRng.Font.Size = 14

    ' Text columns.
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_B), wsLiv.Cells(lastRow, COL_B)): r.NumberFormat = "@"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_C), wsLiv.Cells(lastRow, COL_C)): r.NumberFormat = "@"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_E), wsLiv.Cells(lastRow, COL_E)): r.NumberFormat = "@"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_W), wsLiv.Cells(lastRow, COL_W)): r.NumberFormat = "@"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_Y), wsLiv.Cells(lastRow, COL_Y)): r.NumberFormat = "@"

    ' Numeric columns.
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_D), wsLiv.Cells(lastRow, COL_D)): r.NumberFormat = "0"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_F), wsLiv.Cells(lastRow, COL_F)): r.NumberFormat = "0"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_G), wsLiv.Cells(lastRow, COL_G)): r.NumberFormat = "0"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_V), wsLiv.Cells(lastRow, COL_V)): r.NumberFormat = "0"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_K), wsLiv.Cells(lastRow, COL_K)): r.NumberFormat = "0%"

    ' Date columns.
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_H), wsLiv.Cells(lastRow, COL_H)): r.NumberFormat = "0"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_I), wsLiv.Cells(lastRow, COL_I)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_J), wsLiv.Cells(lastRow, COL_J)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_L), wsLiv.Cells(lastRow, COL_L)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_M), wsLiv.Cells(lastRow, COL_M)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 14), wsLiv.Cells(lastRow, 14)): r.NumberFormat = "dd/mm/yyyy" ' N
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_O), wsLiv.Cells(lastRow, COL_O)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 16), wsLiv.Cells(lastRow, 16)): r.NumberFormat = "dd/mm/yyyy" ' P
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 17), wsLiv.Cells(lastRow, 17)): r.NumberFormat = "dd/mm/yyyy" ' Q
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 18), wsLiv.Cells(lastRow, 18)): r.NumberFormat = "dd/mm/yyyy" ' R
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, 19), wsLiv.Cells(lastRow, 19)): r.NumberFormat = "dd/mm/yyyy" ' S
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_T), wsLiv.Cells(lastRow, COL_T)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_U), wsLiv.Cells(lastRow, COL_U)): r.NumberFormat = "dd/mm/yyyy"
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_X), wsLiv.Cells(lastRow, COL_X)): r.NumberFormat = "dd/mm/yyyy"

    ' Grey background for selected metric/date columns.
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_E), wsLiv.Cells(lastRow, COL_E)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_I), wsLiv.Cells(lastRow, COL_I)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_J), wsLiv.Cells(lastRow, COL_J)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_K), wsLiv.Cells(lastRow, COL_K)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_M), wsLiv.Cells(lastRow, COL_M)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_O), wsLiv.Cells(lastRow, COL_O)): r.Interior.Color = COLOR_METRIC_BG
    Set r = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_T), wsLiv.Cells(lastRow, COL_T)): r.Interior.Color = COLOR_METRIC_BG

    ApplyUVWXBlockBorders wsLiv, lastRow
    ApplyColumnESeparatorBorder wsLiv, lastRow
End Sub

' Applies hard borders to visually separate U:W:X block.
Private Sub ApplyUVWXBlockBorders(wsLiv As Worksheet, ByVal lastRow As Long)
    Dim rngBlock As Range
    Dim rngLeftSep As Range
    Dim rngRightSep As Range

    If lastRow < LIV_FIRST_ROW Then Exit Sub

    Set rngBlock = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_U), wsLiv.Cells(lastRow, COL_X))
    Set rngLeftSep = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_U), wsLiv.Cells(lastRow, COL_U))
    Set rngRightSep = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_X), wsLiv.Cells(lastRow, COL_X))

    With rngBlock.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rngBlock.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rngLeftSep.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rngRightSep.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
End Sub

' Applies a hard separator on the right side of column E.
Private Sub ApplyColumnESeparatorBorder(wsLiv As Worksheet, ByVal lastRow As Long)
    Dim rngSep As Range

    If lastRow < LIV_FIRST_ROW Then Exit Sub
    Set rngSep = wsLiv.Range(wsLiv.Cells(LIV_FIRST_ROW, COL_E), wsLiv.Cells(lastRow, COL_E))

    With rngSep.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
End Sub

' Checks whether a collection contains a value.
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

' Resolves sprint to highlight in yellow for an STR block.
Public Function GetYellowSprintKeyForSTR(strKey As String, maxSprintMap As Object, _
                                        strSprints As Collection) As String
    Dim candidate As String
    Dim k As String
    Dim i As Long

    GetYellowSprintKeyForSTR = ""
    k = Trim$(CStr(strKey & ""))

    If Not maxSprintMap Is Nothing Then
        If maxSprintMap.Exists(k) Then
            candidate = CStr(maxSprintMap(k))
            If candidate <> "" Then
                If CollectionContains(strSprints, candidate) Then
                    GetYellowSprintKeyForSTR = candidate
                    Exit Function
                End If
            End If
        End If
    End If

    For i = strSprints.Count To 1 Step -1
        candidate = CStr(strSprints(i))
        GetYellowSprintKeyForSTR = candidate
        Exit Function
    Next i
End Function

' Copies U:X yellow formatting from template to target rows.
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

' Draws thin gray border around a row block.
Public Sub ApplyLightOutlineBorder(ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, _
                                  ByVal lastCol As Long)
    Dim rng As Range
    If topRow > bottomRow Then Exit Sub
    Set rng = ws.Range(ws.Cells(topRow, 1), ws.Cells(bottomRow, lastCol))
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = COLOR_BORDER_LIGHT
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = COLOR_BORDER_LIGHT
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = COLOR_BORDER_LIGHT
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = COLOR_BORDER_LIGHT
    End With
End Sub

' Draws medium black border around a row block.
Public Sub ApplyHardOutlineBorder(ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, _
                                  ByVal lastCol As Long)
    Dim rng As Range
    If topRow > bottomRow Then Exit Sub
    Set rng = ws.Range(ws.Cells(topRow, 1), ws.Cells(bottomRow, lastCol))
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_BORDER_HARD
    End With
End Sub

' Rebuilds block borders for all STR sections.
Public Sub RebuildSuiviLivrablesBorders(wsLiv As Worksheet, ByVal lastCol As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim curStr As String
    Dim nextStr As String
    Dim swdsStartRow As Long

    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow < LIV_FIRST_ROW Then Exit Sub

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

        swdsStartRow = blockEnd + 1
        For r = blockStart To blockEnd
            If UCase$(Trim$(CStr(wsLiv.Cells(r, COL_C).Value & ""))) = SECTION_SWDS Then
                swdsStartRow = r
                Exit For
            End If
        Next r

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
