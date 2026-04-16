Option Explicit

Private m_ArchiveRunning As Boolean

' Archives Suivi_Livrables into dated workbook and clears current data.
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
    Dim confirmResp As VbMsgBoxResult
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

    ' Validate sheet and prepare output folder/file names.
    If Not SheetExists(SH_LIV) Then
        MsgBox "La feuille """ & SH_LIV & """ est introuvable.", vbExclamation
        Exit Sub
    End If

    confirmResp = MsgBox("Confirmer l'archivage de """ & SH_LIV & """ ?" & vbCrLf & vbCrLf & _
                         "Cette action va sauvegarder l'etat actuel puis vider les lignes actives de la feuille.", _
                         vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation archivage")
    If confirmResp <> vbYes Then Exit Sub

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

    ' Copy sheet values + formatting into a new workbook.
    Set wsLiv = ThisWorkbook.Sheets(SH_LIV)
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

    ' Reset active livrables rows after archive save.
    lastRow = wsLiv.Cells(wsLiv.Rows.Count, COL_B).End(xlUp).Row
    If lastRow >= LIV_FIRST_ROW Then
        wsLiv.Rows(LIV_FIRST_ROW & ":" & lastRow).Delete Shift:=xlUp
    End If

    resp = MsgBox("Archive enregistree et feuille reinitialisee." & vbCrLf & vbCrLf & _
                  "Ouvrir le fichier archive maintenant ?" & vbCrLf & fullPath, _
                  vbYesNo + vbInformation, "Archive")
    If resp = vbYes Then
        ThisWorkbook.FollowHyperlink fullPath
    End If

Cleanup:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    m_ArchiveRunning = False
    Exit Sub

ErrHandler:
    ' Log archive error and restore application state.
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
    MsgBox "Echec de l'archivage : " & Err.Description, vbCritical
End Sub
