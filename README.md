You are writing a VBA macro for Excel.

Create a Sub called `UpdateSuiviLivrable` that synchronizes data from the sheet
"Suivi_CR" into the sheet "Suivi_Livrable". The macro must also read lookup data
from "PowQ_Extract" and "PowQ_Suivi_UVR". Below is the full specification.

═══════════════════════════════════════════════════════
FILE ORGANIZATION
═══════════════════════════════════════════════════════

Split the implementation across TWO .bas files:

  1. UpdateSuiviLivrable.bas  — contains ONLY the main Sub and the
     top-level orchestration logic (lock, diff, dispatch to helpers,
     cleanup). No computation logic lives here.

  2. SuiviUtils.bas           — contains ALL helper functions, lookup
     engines, JSON serializer/parser, and formula-equivalent functions.
     Every reusable or testable piece of logic goes here.

Do NOT put helper functions inside the same module as the main Sub.
Each .bas file must begin with the appropriate Attribute VB_Name line.

═══════════════════════════════════════════════════════
ARCHITECTURE OVERVIEW
═══════════════════════════════════════════════════════

The workbook is shared across multiple users. Each user modifies "Suivi_CR"
and then clicks an "Update" button that runs this macro. The macro:

1. Detects whether another update is in progress (LOCK mechanism)
2. On first run: creates a JSON snapshot of the current Suivi_CR state
3. On subsequent runs: computes a diff between the saved snapshot and
   the current state of Suivi_CR
4. If new STR values appear in column B of Suivi_CR → duplicate template rows
   in Suivi_Livrable (preserving fonts, colors, borders)
5. If no new STR values → only update the cells that changed
6. For every affected row in Suivi_Livrable, compute and write the values
   that would result from the formulas listed below (no live formulas are kept
   in Suivi_Livrable — values only)
7. Save the new snapshot to status.json
8. Release the lock

═══════════════════════════════════════════════════════
SHARED FOLDER PATH — CONFIGURATION
═══════════════════════════════════════════════════════

Define SharedFolderPath as a Public Const at the top of SuiviUtils.bas:

  Public Const SHARED_FOLDER_PATH As String = "\\YOUR_SERVER\YOUR_SHARE\"

This is the ONLY place the path is defined. Both LOCK.txt and status.json
are stored at this path. Before deploying, replace the placeholder with
the actual UNC path or mapped drive letter used by all users (e.g.
"\\fileserver\team\suivi\" or "Z:\suivi\").

Leave a prominent TODO comment above the Const reminding the deployer
to update this value before distributing the workbook.

═══════════════════════════════════════════════════════
JSON LIBRARY — VBA-JSON (JsonConverter.bas)
═══════════════════════════════════════════════════════

Use the open-source VBA-JSON library for all JSON serialization and
parsing. This is the standard approach for VBA JSON handling.

Assume JsonConverter.bas is already imported into the project (available
at https://github.com/VBA-tools/VBA-JSON). Reference it as:

  JsonConverter.ParseJson(jsonString)   ' returns a Collection/Dictionary
  JsonConverter.ConvertToJson(obj)      ' returns a JSON string

Do NOT write a custom JSON parser. Do NOT reinvent serialization.

At the top of SuiviUtils.bas, add a comment block:
  ' DEPENDENCY: JsonConverter.bas must be imported into this VBA project.
  ' Download from https://github.com/VBA-tools/VBA-JSON
  ' Also requires: Tools > References > Microsoft Scripting Runtime (for Dictionary)

═══════════════════════════════════════════════════════
CONCURRENCY — LOCK FILE
═══════════════════════════════════════════════════════

Lock file path: SHARED_FOLDER_PATH & "LOCK.txt"

On entry:
  If LOCK.txt exists → MsgBox "An update is already in progress by another
  user. Please wait and try again." and Exit Sub immediately.

If LOCK.txt does NOT exist → create it immediately. Write inside it:
  "LOCKED by: " & Environ("USERNAME") & " at " & Now()

On exit (BOTH success AND any error — use the Cleanup label pattern):
  If FileExists(lockPath) Then Kill lockPath

═══════════════════════════════════════════════════════
STATE FILE — status.json
═══════════════════════════════════════════════════════

Path: SHARED_FOLDER_PATH & "status.json"

Schema — a JSON array, one object per data row of Suivi_CR (starting row 2):
[
  {
    "STR":   "<value of column B>",
    "row":   <integer row number in Suivi_CR>,
    "cells": {
      "B": "<value>", "C": "<value>", "D": "<value>",
      "E": "<value>", ... (all columns in used range)
    }
  },
  ...
]

FIRST RUN (status.json does not exist or is empty / zero-length):
  - Read the entire used data range of Suivi_CR rows 2 onward
  - Serialize to JSON using JsonConverter.ConvertToJson
  - Save to SHARED_FOLDER_PATH & "status.json"
  - MsgBox "Initial snapshot created. The sheet is now tracked.
    Run Update again to perform the first synchronization."
  - Jump to Cleanup (delete lock, restore app settings) and Exit Sub

SUBSEQUENT RUNS:
  - Load status.json → parse with JsonConverter.ParseJson into a
    VBA Collection of Dictionaries keyed by STR value
  - Read current Suivi_CR into a Variant array
  - Compute diff: new STR rows + modified cells in existing rows

═══════════════════════════════════════════════════════
UPDATE LOGIC — NEW STR ROWS
═══════════════════════════════════════════════════════

If column B of Suivi_CR contains STR values NOT present in status.json:

TEMPLATE ROW DETECTION:
  The template row is identified as the last row in Suivi_Livrable that:
    - Has an EMPTY column B value, AND
    - Has visible formatting (non-default border or fill color) applied
  If no such row is found, raise an error:
    Err.Raise vbObjectError + 1001, , "Template row not found in
    Suivi_Livrable. Ensure a formatted empty row exists as the template."
  Leave a TODO comment: verify the template row detection logic matches
  the actual sheet layout before deploying.

FOR EACH new STR value (in the order they appear in Suivi_CR):
  a. Insert a new row in Suivi_Livrable after the last populated data row
     (defined as the last row with a non-empty col B), before any
     footer/total rows. Leave a TODO to confirm whether footer rows exist.
  b. Copy template row formats only:
       templateRow.Copy
       newRow.PasteSpecial Paste:=xlPasteFormats
       Application.CutCopyMode = False
  c. Write values from Suivi_CR into the mapped columns (see COLUMN MAPPING)
  d. Compute and write formula-result values for cols H, I, J, K, O, T

═══════════════════════════════════════════════════════
UPDATE LOGIC — MODIFIED CELLS (no new STR)
═══════════════════════════════════════════════════════

For each row in Suivi_CR whose cell values differ from the snapshot:
  - Identify which columns changed
  - Locate the matching row in Suivi_Livrable by STR value (col B)
  - Write updated values to the corresponding columns
  - Re-compute and overwrite cols H, I, J, K, O, T for that row

═══════════════════════════════════════════════════════
COLUMN MAPPING  Suivi_CR → Suivi_Livrable
═══════════════════════════════════════════════════════

Copy these columns verbatim (value paste only):
  Suivi_CR col B  →  Suivi_Livrable col B   (STR identifier — row key)
  Suivi_CR col C  →  Suivi_Livrable col C
  Suivi_CR col D  →  Suivi_Livrable col D
  Suivi_CR col E  →  Suivi_Livrable col E
  [Add further mappings as needed — leave a TODO comment listing any
   Suivi_CR columns whose destination in Suivi_Livrable is unclear]

Row matching key across both sheets: column B (STR value).

═══════════════════════════════════════════════════════
FORMULA LOGIC  (compute in VBA, write as values)
═══════════════════════════════════════════════════════

For every affected row in Suivi_Livrable, compute values for columns
H, I, J, K, O, T using VBA logic equivalent to the formulas below.

Variables per row:
  B_val = Suivi_Livrable col B value (STR)
  C_val = Suivi_Livrable col C value
  D_val = Suivi_Livrable col D value
  E_val = Suivi_Livrable col E value

All PowQ_Extract lookups operate on a Variant array pre-loaded ONCE
before the main loop (never read the sheet row-by-row inside a loop).

── Column H ──────────────────────────────────────────
Original formula:
  =SUMIFS(PowQ_Extract!Y:Y,
          PowQ_Extract!B:B, $B4,
          PowQ_Extract!C:C, $E4,
          PowQ_Extract!F:F, $C4,
          PowQ_Extract!U:U, $D4)

VBA logic: loop over the PowQ_Extract array. For each row where
  col B = B_val AND col C = E_val AND col F = C_val AND col U = D_val,
  accumulate col Y as a Double. Return the total (0 if no match).

── Column I ──────────────────────────────────────────
Original formula:
  =IFERROR(
     VLOOKUP(
       $B4 & "/" & $E4 & "/" & $C4 & "/Sprint " & $D4,
       PowQ_Extract!A2:ZZ20706,
       MATCH("Fin Ref", PowQ_Extract!A1:W1, 0),
       0),
   "")

VBA logic:
  1. Build lookup key: B_val & "/" & E_val & "/" & C_val & "/Sprint " & D_val
  2. Find "Fin Ref" column index by scanning PowQ_Extract row 1 (header row)
     — cache this column index once before the loop, not per row
  3. Search PowQ_Extract col A for the exact key; on match return the
     value at the cached "Fin Ref" column index
  4. If not found → return ""

── Column J ──────────────────────────────────────────
Original formula (LET-based):
  MAX of PowQ_Extract col I values
  where col B = $B4 AND col C = $E4 AND col U = $D4
        AND col F = $C4 AND col I <> ""

VBA logic: loop PowQ_Extract array, filter by all four conditions plus
non-empty col I. Cast matching values to Double, track maximum.
Return the maximum as a string, or "" if no matches found.

── Column K ──────────────────────────────────────────
Original formula (SI.CONDITIONS / IFS-based weighted average):

  IMPORTANT ROW OFFSET: The original formula references row N+1 of
  Suivi_Livrable to look up data from Suivi_CR. This means when computing
  col K for Suivi_Livrable row N, the lookup variables are:
    B_val = Suivi_Livrable col B at row N+1
    C_val = Suivi_Livrable col C at row N+1
    D_val = Suivi_Livrable col D at row N+1
    E_val = Suivi_Livrable col E at row N+1
  Pass (rowNum + 1) as the source row when calling ComputeColK.
  Leave a TODO comment here to verify this offset with the sheet owner
  before deploying.

Logic:
  If B_val = "" → return ""

  If C_val = "ADL1":
    Scan Suivi_CR rows 2–9976 for rows where:
      col B = B_val AND col C = D_val AND col D = E_val
      AND col Z <> "" AND col J <> ""
    Numerator   = SUM of (col J × col Z) for matching rows
    Denominator = SUM of col Z for matching rows
    If Denominator = 0 → return ""
    Else return Numerator / Denominator

  If C_val = "SwDS" OR C_val = "Reprise suite valid":
    Same as above but use col L instead of col J:
    Numerator   = SUM of (col L × col Z) for matching rows
    Denominator = SUM of col Z for matching rows
    If Denominator = 0 → return ""
    Else return Numerator / Denominator

  Else → return ""

Suivi_CR data must also be pre-loaded into a Variant array once before
the main loop.

── Column O ──────────────────────────────────────────
Original formula:
  =IFERROR(
     VLOOKUP(
       $B4 & "/" & $E4 & "/UVR " & $C4 & "/Sprint " & $D4,
       PowQ_Extract!A2:I20706, 9, 0),
   "")

VBA logic:
  Lookup key: B_val & "/" & E_val & "/UVR " & C_val & "/Sprint " & D_val
  Search PowQ_Extract col A for exact match; return col I (index 9).
  If not found → return ""

── Column T ──────────────────────────────────────────
Original formula:
  =IFERROR(
     VLOOKUP(
       $B4 & "/" & $E4 & "/UVR " & $C4 & " OK/Sprint " & $D4,
       PowQ_Extract!A2:I20706, 9, 0),
   "")

VBA logic:
  Lookup key: B_val & "/" & E_val & "/UVR " & C_val & " OK/Sprint " & D_val
  (note the " OK" suffix before "/Sprint " — the only difference from col O)
  Search PowQ_Extract col A for exact match; return col I (index 9).
  If not found → return ""

═══════════════════════════════════════════════════════
PERFORMANCE REQUIREMENTS
═══════════════════════════════════════════════════════

- Load PowQ_Extract into a Variant array ONCE at entry:
    Dim powqArr As Variant
    powqArr = Sheets("PowQ_Extract").UsedRange.Value
- Load Suivi_CR into a Variant array ONCE at entry:
    Dim crArr As Variant
    crArr = Sheets("Suivi_CR").UsedRange.Value
- Never access sheet cells inside any loop.
- Write results to Suivi_Livrable in bulk where possible (build an output
  array, then assign the entire array to the target range in one step).
- Disable at entry, re-enable at Cleanup:
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

═══════════════════════════════════════════════════════
ERROR HANDLING
═══════════════════════════════════════════════════════

Use this pattern in UpdateSuiviLivrable.bas:

  On Error GoTo ErrHandler
  ' ... main logic ...
  GoTo Cleanup

ErrHandler:
  MsgBox "Update failed: " & Err.Description & _
         " (Error " & Err.Number & ")", vbCritical, "Suivi Update"

Cleanup:
  If FileExists(SHARED_FOLDER_PATH & "LOCK.txt") Then
      Kill SHARED_FOLDER_PATH & "LOCK.txt"
  End If
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

═══════════════════════════════════════════════════════
HELPER FUNCTIONS — ALL GO IN SuiviUtils.bas
═══════════════════════════════════════════════════════

Write every function below in SuiviUtils.bas, not in the main module.

1.  Function FileExists(path As String) As Boolean

2.  Function ReadTextFile(path As String) As String

3.  Sub WriteTextFile(path As String, content As String)

4.  Function SerializeSnapshotToJson(crArr As Variant) As String
      -- Serialize the Suivi_CR Variant array (rows 2 onward) to a JSON
         string matching the status.json schema defined above.
      -- Use JsonConverter.ConvertToJson internally.

5.  Function ParseSnapshotFromJson(jsonStr As String) As Object
      -- Parse the JSON string into a Scripting.Dictionary keyed by STR
         value (col B), where each value is itself a Dictionary of
         { colLetter: cellValue }.
      -- Use JsonConverter.ParseJson internally.

6.  Function FindFinRefColumn(powqArr As Variant) As Long
      -- Scan row 1 of powqArr for the header "Fin Ref".
      -- Return its 1-based column index, or 0 if not found.

7.  Function ComputeColH(B As String, C As String, D As String, _
                         E As String, powqArr As Variant) As Double

8.  Function ComputeColI(B As String, C As String, D As String, _
                         E As String, powqArr As Variant, _
                         finRefCol As Long) As String

9.  Function ComputeColJ(B As String, C As String, D As String, _
                         E As String, powqArr As Variant) As String

10. Function ComputeColK(B As String, C As String, D As String, _
                         E As String, crArr As Variant) As String
      -- B/C/D/E are from Suivi_Livrable row N+1 (see offset note above)

11. Function ComputeColO(B As String, C As String, D As String, _
                         E As String, powqArr As Variant) As String

12. Function ComputeColT(B As String, C As String, D As String, _
                         E As String, powqArr As Variant) As String

13. Function GetLastDataRow(ws As Worksheet, keyCol As Long) As Long
      -- Return the last row index in ws where column keyCol is non-empty.

14. Function FindRowBySTR(livArr As Variant, strVal As String) As Long
      -- Search col B of the Suivi_Livrable Variant array for strVal.
      -- Return the 1-based row index, or 0 if not found.

═══════════════════════════════════════════════════════
SHEET & DATA CONSTANTS
═══════════════════════════════════════════════════════

Define these as Public Consts at the top of SuiviUtils.bas:

  Public Const SH_CR         As String = "Suivi_CR"
  Public Const SH_LIV        As String = "Suivi_Livrable"
  Public Const SH_EXTRACT    As String = "PowQ_Extract"
  Public Const SH_UVR        As String = "PowQ_Suivi_UVR"
  Public Const CR_FIRST_ROW  As Long   = 2     ' Suivi_CR data starts row 2
  Public Const LIV_FIRST_ROW As Long   = 4     ' Suivi_Livrable data starts row 4
  Public Const CR_MAX_ROW    As Long   = 9976  ' used in col K scan range

═══════════════════════════════════════════════════════
TODO CHECKLIST — verify with sheet owner before deploying
═══════════════════════════════════════════════════════

Leave these as TODO comments in the code at the relevant locations:

  TODO-1: Set SHARED_FOLDER_PATH to the correct UNC or drive path.
  TODO-2: Confirm template row detection logic (empty col B + formatting).
  TODO-3: Confirm whether Suivi_Livrable has footer/total rows that must
          be skipped when inserting new rows.
  TODO-4: Confirm the col K row offset (+1) with the sheet owner — verify
          that col K for row N reads lookup keys from row N+1.
  TODO-5: Complete the column mapping table (which Suivi_CR columns beyond
          B, C, D, E are copied to Suivi_Livrable, and to which columns).
  TODO-6: Confirm LIV_FIRST_ROW = 4 (rows 1-3 are headers/template).
  TODO-7: Import JsonConverter.bas and enable Microsoft Scripting Runtime
          before distributing the workbook.