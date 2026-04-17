Option Explicit

Private m_ArchiveBNRunning As Boolean

' Archives BN_Suivi dossier Safety and clears all rows except 1 and 2.
Public Sub ArchiveBNSuivi()
    Dim wsBN As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim srcRng As Range
    Dim dstRng As Range
    Dim folderPath As String
    Dim dayFolder As String
    Dim fileName As String
    Dim fullPath As String
    Dim ts As String
    Dim confirmResp As VbMsgBoxResult
    Dim resp As VbMsgBoxResult
    Dim shp As Shape
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long
    Dim r As Long
    Dim errLine As String
    Dim sharedFolderPath As String

    If m_ArchiveBNRunning Then Exit Sub
    m_ArchiveBNRunning = True

    On Error GoTo ErrHandler

    If Not SheetExists("BN_Suivi dossier Safety") Then
        MsgBox "La feuille ""BN_Suivi dossier Safety"" est introuvable.", vbExclamation
        Exit Sub
    End If

    confirmResp = MsgBox("Confirmer l'archivage de ""BN_Suivi dossier Safety"" ?" & vbCrLf & vbCrLf & _
                         "Cette action va sauvegarder l'etat actuel puis vider toutes les lignes.", _
                         vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation archivage")
    If confirmResp <> vbYes Then GoTo Cleanup

    On Error Resume Next
    sharedFolderPath = SHARED_FOLDER_PATH(False)
    If Err.Number <> 0 Or Trim$(sharedFolderPath) = "" Then
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "La selection du dossier partage n'a pas ete finalisee correctement." & vbCrLf & _
               "L'archivage est annule.", vbExclamation, "Archivage BN_Suivi"
        GoTo Cleanup
    End If
    On Error GoTo ErrHandler
    If Right$(sharedFolderPath, 1) <> "\" Then sharedFolderPath = sharedFolderPath & "\"

    folderPath = sharedFolderPath & "Archived\"
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath

    folderPath = folderPath & "BN_Suivi\"
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath

    dayFolder = folderPath & Format$(Date, "DDMMYYYY") & "\"
    If Dir$(dayFolder, vbDirectory) = "" Then MkDir dayFolder

    ts = Format$(Now, "DDMMYYYY_HHMMSS")
    fileName = "BN_Suivi_dossier_Safety_" & ts & ".xlsx"
    fullPath = dayFolder & fileName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wsBN = ThisWorkbook.Sheets("BN_Suivi dossier Safety")
    If wsBN.AutoFilterMode Then wsBN.AutoFilterMode = False

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = wsBN.Name

    lastRow = wsBN.UsedRange.Row + wsBN.UsedRange.Rows.Count - 1
    lastCol = wsBN.UsedRange.Column + wsBN.UsedRange.Columns.Count - 1
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    Set srcRng = wsBN.Range(wsBN.Cells(1, 1), wsBN.Cells(lastRow, lastCol))
    Set dstRng = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lastRow, lastCol))

    srcRng.Copy
    dstRng.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    dstRng.Value = srcRng.Value

    For c = 1 To lastCol
        wsNew.Columns(c).ColumnWidth = wsBN.Columns(c).ColumnWidth
    Next c
    For r = 1 To lastRow
        wsNew.Rows(r).RowHeight = wsBN.Rows(r).RowHeight
    Next r

    For Each shp In wsNew.Shapes
        shp.Delete
    Next shp

    wbNew.SaveAs fileName:=fullPath, _
                  FileFormat:=xlOpenXMLWorkbook, _
                  CreateBackup:=False
    wbNew.Close SaveChanges:=False

    lastRow = wsBN.Cells(wsBN.Rows.Count, COL_B).End(xlUp).Row
    If lastRow > 2 Then
        wsBN.Rows("3:" & lastRow).Delete Shift:=xlUp
    End If

    resp = MsgBox("Archive BN_Suivi enregistree et feuille reinitialisee." & vbCrLf & vbCrLf & _
                  "Ouvrir le fichier archive maintenant ?" & vbCrLf & fullPath, _
                  vbYesNo + vbInformation, "Archive BN_Suivi")
    If resp = vbYes Then
        ThisWorkbook.FollowHyperlink fullPath
    End If

Cleanup:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    m_ArchiveBNRunning = False
    Exit Sub

ErrHandler:
    errLine = Format$(Now, "YYYY-MM-DD HH:NN:SS") & _
              " | user=" & Environ$("USERNAME") & _
              " | proc=ArchiveBNSuivi" & _
              " | err=" & Err.Number & _
              " | " & Err.Description
    On Error Resume Next
    If Trim$(sharedFolderPath) <> "" Then
        AppendTextFile sharedFolderPath & "error_logs.txt", errLine
    End If
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    m_ArchiveBNRunning = False
    MsgBox "Echec de l'archivage BN_Suivi : " & Err.Description, vbCritical
End Sub
