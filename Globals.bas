Option Explicit

' ---------------------------------------------------------
'  SHEET NAME CONSTANTS
' ---------------------------------------------------------

Public Const SH_CR         As String = "Suivi_CR"
Public Const SH_LIV        As String = "Suivi_Livrables"
Public Const SH_EXTRACT    As String = "PowQ_Extract"
Public Const SH_UVR        As String = "PowQ_Suivi_UVR"
Public Const SH_VHST       As String = "PowQ_EDU_CE_VHST"
Public Const SH_TMP        As String = "Suivi_Livrables_Tmp"

' ---------------------------------------------------------
'  ROW CONSTANTS
' ---------------------------------------------------------

Public Const CR_FIRST_ROW  As Long = 3
Public Const LIV_FIRST_ROW As Long = 4

' Template block in Suivi_Livrables_Tmp (rows TMP_FIRST_ROW:TMP_LAST_ROW).
' Sprints detected from col D; only sprints present in Suivi_CR are copied.
Public Const TMP_FIRST_ROW As Long = 4
Public Const TMP_LAST_ROW  As Long = 33

' ---------------------------------------------------------
'  COLUMN INDEX CONSTANTS (1-based)
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
Public Const COL_X As Long = 24
Public Const COL_Y As Long = 25
Public Const COL_Z As Long = 26
