Option Explicit

'===========================================================
' POST-SYNC FORMATTING FOR ALL TABLES
'===========================================================
Public Sub Post_Sync_Format()
    ' NOTE: Caller must handle Global_Unprotect/Global_Protect
    ' This function only handles data normalization and visual formatting
    ' NO protection logic here - that's in Global_Protect
    
    Format_Master_Equipment_List
End Sub


'===========================================================
' FORMAT MASTER EQUIPMENT LIST
'===========================================================
Private Sub Format_Master_Equipment_List()
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set lo = ws.ListObjects(TABLE_MASTER)
    
    If Not TableHasRows(lo) Then Exit Sub
    
    ' 1) Normalize blank-Source rows to MAN & auto-fill Include columns
    ManualEntryDef_Master lo
    
    ' 2) Shade locked rows (visual indicator only)
    MarkLockedItems_Master lo
End Sub


'===========================================================
' MANUAL ENTRY DEFAULTS - MASTER
'===========================================================
Private Sub ManualEntryDef_Master(lo As ListObject, Optional rng As Range)
    Dim sourceCol As Long
    Dim Target As Range
    Dim rowRange As Range

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    sourceCol = GetTableColIndex(lo, "Source")
    If sourceCol = 0 Then Exit Sub

    ' Decide whether we're fixing the whole table or just one row
    If rng Is Nothing Then
        Set Target = lo.DataBodyRange
    Else
        Set Target = Intersect(rng.EntireRow, lo.DataBodyRange)
        If Target Is Nothing Then Exit Sub
    End If

    For Each rowRange In Target.Rows
        ' Set blank Source to MAN
        If Trim$(CStr(rowRange.Cells(1, sourceCol).Value)) = "" Then
            rowRange.Cells(1, sourceCol).Value = "MAN"
            
            ' Set defaults for new manual entries
            SetIfColExists lo, rowRange, "P&ID Tags", ""
            SetIfColExists lo, rowRange, "Removed from BOM", "N"
            SetIfColExists lo, rowRange, "Notes", ""
            
            ' Clear inherited shading
            rowRange.Interior.Pattern = xlNone
            
            ' Manual rows should be fully editable
            rowRange.Locked = False
        End If
        
        ' Auto-fill Include columns with "N" if empty (ALL rows)
        SetDefaultIfBlank lo, rowRange, "Include in I/O List?", "N"
        SetDefaultIfBlank lo, rowRange, "Include in Utility Load Table?", "N"
        SetDefaultIfBlank lo, rowRange, "Include in Heat Load & Noise Table?", "N"
    Next rowRange
End Sub


'===========================================================
' MARK LOCKED ITEMS - MASTER (VISUAL SHADING ONLY)
'===========================================================
Private Sub MarkLockedItems_Master(lo As ListObject)
    Dim sourceCol As Long
    Dim removedCol As Long
    Dim itemCol As Long
    Dim r As Long
    Dim sourceCell As Range, remCell As Range, itemCell As Range

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    sourceCol = GetTableColIndex(lo, "Source")
    removedCol = GetTableColIndex(lo, "Removed from BOM")
    itemCol = GetTableColIndex(lo, "Master Equipment List Item")

    If sourceCol = 0 Or removedCol = 0 Or itemCol = 0 Then Exit Sub

    Application.ScreenUpdating = False

    For r = 1 To lo.DataBodyRange.Rows.Count
        Set sourceCell = lo.DataBodyRange.Cells(r, sourceCol)
        Set remCell = lo.DataBodyRange.Cells(r, removedCol)
        Set itemCell = lo.DataBodyRange.Cells(r, itemCol)

        ' Shade locked BOM rows (visual indicator)
        If sourceCell.Locked And remCell.Locked Then
            sourceCell.Interior.Color = RGB(230, 230, 230)
            remCell.Interior.Color = RGB(230, 230, 230)
            itemCell.Interior.Color = RGB(230, 230, 230)
        Else
            sourceCell.Interior.Pattern = xlNone
            remCell.Interior.Pattern = xlNone
            itemCell.Interior.Pattern = xlNone
        End If
    Next r

    Application.ScreenUpdating = True
End Sub


'===========================================================
' HELPER FUNCTIONS
'===========================================================
Private Sub SetIfColExists(lo As ListObject, rowRange As Range, colTitle As String, ByVal val As Variant)
    Dim idx As Long
    idx = GetTableColIndex(lo, colTitle)
    If idx > 0 Then
        rowRange.Cells(1, idx).Value = val
    End If
End Sub


Private Sub SetDefaultIfBlank(lo As ListObject, rowRange As Range, colTitle As String, defaultVal As String)
    Dim idx As Long
    Dim cellVal As String
    
    idx = GetTableColIndex(lo, colTitle)
    If idx > 0 Then
        cellVal = Trim$(CStr(rowRange.Cells(1, idx).Value))
        If cellVal = "" Then
            rowRange.Cells(1, idx).Value = defaultVal
        End If
    End If
End Sub
