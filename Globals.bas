Option Explicit

' Sheet names used across update workflows.
Public Const SH_CR         As String = "Suivi_CR"
Public Const SH_LIV        As String = "Suivi_Livrables"
Public Const SH_EXTRACT    As String = "PowQ_Extract"
Public Const SH_UVR        As String = "PowQ_Suivi_UVR"
Public Const SH_VHST       As String = "PowQ_EDU_CE_VHST"
Public Const SH_BN         As String = "BN_Suivi dossier Safety"
Public Const SH_CONFIG     As String = "config"

' Input sheet names used by update flows.
Public Const SH_IN_EXTRACT As String = "Extract"
Public Const SH_IN_UVR     As String = "Global"
Public Const SH_IN_VHST  As String = "Références_CE_VHST"

' Table names used in PowQ sheets.
Public Const TBL_EXTRACT   As String = "Extract_MSP"
Public Const TBL_UVR       As String = "Suivi_UVR"
Public Const TBL_EDU       As String = "EDU_CE_VHST"

' UVR source layout in input sheet.
Public Const UVR_HEADER_ROW As Long = 12
Public Const UVR_FIRST_DATA_ROW As Long = 13

' Generic row/header constants.
Public Const HEADER_ROW_1 As Long = 1
Public Const DATA_ROW_2 As Long = 2
Public Const DATA_ROW_3 As Long = 3

' Shared lock/update constants.
Public Const LOCK_CELL_ADDR As String = "I1"
Public Const LOCK_PREFIX As String = "LOCKED by: "
Public Const LOCK_SEPARATOR As String = " at "
Public Const LOCK_DATE_FORMAT As String = "YYYY-MM-DD HH:NN:SS"
Public Const LOCK_STALE_MINUTES As Double = 30#
Public Const PROTECT_PASSWORD As String = "suivi_update"

' BN archive constants.
Public Const ARCHIVE_ROOT_FOLDER As String = "Archived\"
Public Const ARCHIVE_BN_FOLDER As String = "BN_Suivi\"
Public Const DATE_FOLDER_FORMAT As String = "DDMMYYYY"
Public Const TS_FILE_FORMAT As String = "DDMMYYYY_HHMMSS"
Public Const BN_ARCHIVE_FILE_PREFIX As String = "BN_Suivi_dossier_Safety_"
Public Const BN_FIRST_DATA_ROW As Long = 3
Public Const BN_OBSOLETE_PREVIEW_MAX As Long = 12

' Suivi/Config constants.
Public Const TYPE_LIVRABLE_ADL1 As String = "ADL1"
Public Const TYPE_LIVRABLE_SWDS As String = "SwDS"
Public Const UVR_HEADER_ROW_IN_LIV As Long = 3
Public Const YES_FR As String = "OUI"
Public Const BLOCKED_FR As String = "bloque"
Public Const FILE_DIALOG_FOLDER_PICKER As Long = 4

' Header labels used in source sheets.
Public Const HDR_NOM_STR As String = "Nom_STR"
Public Const HDR_SPRINTS As String = "Sprint"
Public Const HDR_FONCTIONS As String = "Fonctions"
Public Const HDR_TYPE_LIVRABLE As String = "Type de livrable"
Public Const HDR_POLE As String = "Pole"
Public Const HDR_FIN_REF As String = "fin ref"
Public Const HDR_COLLABORATEURS As String = "Collaborateurs"
Public Const HDR_SOCIETES As String = "Sociétés"
Public Const HDR_INFO_COMPLET As String = "Info_Complet"

' First data row indexes per sheet.
Public Const CR_FIRST_ROW  As Long = 3
Public Const LIV_FIRST_ROW As Long = 4

Public Const TMP_FIRST_ROW As Long = 4
Public Const TMP_LAST_ROW  As Long = 33

' Column indexes used by the update logic.
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
Public Const COL_M As Long = 13
Public Const COL_O As Long = 15
Public Const COL_T As Long = 20
Public Const COL_U As Long = 21
Public Const COL_V As Long = 22
Public Const COL_W As Long = 23
Public Const COL_X As Long = 24
Public Const COL_Y As Long = 25

' Placeholder colors (BGR hex in VBA Long). Adjust values as needed.
Public Const COLOR_B_BASE_ADL1 As Long = &H00FBEDCA
Public Const COLOR_B_BASE_SWDS As Long = &H00D0F2DA
Public Const COLOR_C_ADL1 As Long = &H00F3CB61
Public Const COLOR_C_SWDS As Long = &H0073D98E
Public Const COLOR_BASE_SPRINT As Long = &H00D5E2FB
Public Const COLOR_YELLOW_ZONE As Long = &H00AFECFF
Public Const COLOR_UX_DEFAULT As Long = &H00BFBFBF
Public Const COLOR_METRIC_BG As Long = &H00D9D9D9
Public Const COLOR_BORDER_LIGHT As Long = &H00969696
Public Const COLOR_BORDER_HARD As Long = &H00000000

' Darkening step (0-1): each sprint increases darkness.
Public Const SPRINT_DARKEN_STEP As Double = 0.08

' Global header lists for PowQ updates.
Public Function GetPowQExtractHeaders() As Variant
    GetPowQExtractHeaders = Array( _
        "Code tache", "str", "Fonction_1", "Pôle_2", "Nom_3", _
        "Nom simple", "Ressource", "Début_4", "Fin_5", "Travaille", _
        "% reel", "% théorique", "Début Ref", "Fin Ref", "travail ref", _
        "Dec", "Raison", "Début précedent", "Fin précédent", "Actif_6", _
        "Sprint", "Reprise valid", "Valid fonction", "Ressources", "NB CR")
End Function

Public Function GetPowQEduHeaders() As Variant
    GetPowQEduHeaders = Array(HDR_NOM_STR, HDR_SPRINTS, HDR_COLLABORATEURS, HDR_SOCIETES, HDR_INFO_COMPLET)
End Function
