Option Explicit

' Sheet names used across update workflows.
Public Const SH_CR         As String = "Suivi_CR"
Public Const SH_LIV        As String = "Suivi_Livrables"
Public Const SH_EXTRACT    As String = "PowQ_Extract"
Public Const SH_UVR        As String = "PowQ_Suivi_UVR"
Public Const SH_VHST       As String = "PowQ_EDU_CE_VHST"
Public Const SH_TMP        As String = "Suivi_Livrables_Tmp"

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

' Dynamic block generation settings.
Public Const SECTION_ADL1 As String = "ADL1"
Public Const SECTION_SWDS As String = "SWDS"

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
