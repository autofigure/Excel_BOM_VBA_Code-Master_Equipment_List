Option Explicit

'===========================================================
' WORKSHEET CHANGE EVENT - MASTER EQUIPMENT LIST
'===========================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim wasProtected As Boolean

    Set ws = Me

    ' try to get the table
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_MASTER)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    ' only care if the change touched the table body
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    ' remember if sheet/workbook was protected
    wasProtected = ws.ProtectContents

    ' turn events off for the whole routine
    Application.EnableEvents = False

    ' unprotect if needed
    If wasProtected Then
        Global_Unprotect
    End If

    ' normalize just the row that was touched
    ManualEntryDef_MasterSingle lo, Target

    ' re-protect if it was protected
    If wasProtected Then
        Global_Protect
    End If

    ' turn events back on
    Application.EnableEvents = True
End Sub


'===========================================================
' MANUAL ENTRY DEF FOR SINGLE ROW
'===========================================================
Private Sub ManualEntryDef_MasterSingle(lo As ListObject, rng As Range)
    Dim sourceCol As Long
    Dim Target As Range
    Dim rowRange As Range

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    sourceCol = GetTableColIndex(lo, "Source")
    If sourceCol = 0 Then Exit Sub

    Set Target = Intersect(rng.EntireRow, lo.DataBodyRange)
    If Target Is Nothing Then Exit Sub

    For Each rowRange In Target.Rows
        If Trim$(CStr(rowRange.Cells(1, sourceCol).Value)) = "" Then
            rowRange.Cells(1, sourceCol).Value = "MAN"
            
            ' Set defaults for manual entries
            SetIfColExistsSingle lo, rowRange, "P&ID Tags", ""
            SetIfColExistsSingle lo, rowRange, "Include in I/O List?", "N"
            SetIfColExistsSingle lo, rowRange, "Include in Utility Load Table?", "N"
            SetIfColExistsSingle lo, rowRange, "Include in Heat Load & Noise Table?", "N"
            SetIfColExistsSingle lo, rowRange, "Removed from BOM", "N"
            SetIfColExistsSingle lo, rowRange, "Notes", ""
            
            ' Clear inherited shading
            rowRange.Interior.Pattern = xlNone
            
            ' Manual rows should be fully editable
            rowRange.Locked = False
        End If
    Next rowRange
End Sub


'===========================================================
' HELPER
'===========================================================
Private Sub SetIfColExistsSingle(lo As ListObject, rowRange As Range, colTitle As String, ByVal val As Variant)
    Dim idx As Long
    idx = GetTableColIndex(lo, colTitle)
    If idx > 0 Then
        ' Only set if currently blank
        If Trim$(CStr(rowRange.Cells(1, idx).Value)) = "" Then
            rowRange.Cells(1, idx).Value = val
        End If
    End If
End Sub

