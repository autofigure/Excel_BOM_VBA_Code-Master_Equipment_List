Option Explicit

' Module-level variable to track row count
Private m_LastRowCount As Long

'===========================================================
' TABLE UPDATE EVENT - DETECTS ROW ADDITIONS
'===========================================================
Private Sub Worksheet_TableUpdate(ByVal Target As TableObject)
    Dim lo As ListObject
    Dim currentRowCount As Long
    Dim sourceCol As Long
    Dim lastRow As Range
    Dim numNewRows As Long
    Dim i As Long
    Dim checkRow As Range

    ' Only process if it's the Master Equipment table
    On Error Resume Next
    Set lo = Me.ListObjects(TABLE_MASTER)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    ' Make sure it's our table
    If Target.Name <> TABLE_MASTER Then Exit Sub
    
    ' Get current row count
    If lo.DataBodyRange Is Nothing Then
        currentRowCount = 0
    Else
        currentRowCount = lo.DataBodyRange.Rows.Count
    End If
    
    ' Initialize tracking variable on first run
    If m_LastRowCount = 0 Then
        m_LastRowCount = currentRowCount
        Exit Sub
    End If
    
    ' Check if rows were added
    If currentRowCount <= m_LastRowCount Then
        ' No new rows, just update count and exit
        m_LastRowCount = currentRowCount
        Exit Sub
    End If
    
    ' Calculate how many new rows were added
    numNewRows = currentRowCount - m_LastRowCount
    m_LastRowCount = currentRowCount
    
    sourceCol = GetTableColIndex(lo, "Source")
    If sourceCol = 0 Then Exit Sub
    
    ' Apply defaults to all new rows (starting from the end)
    Application.EnableEvents = False
    
    For i = currentRowCount - numNewRows + 1 To currentRowCount
        Set checkRow = lo.DataBodyRange.Rows(i)
        
        ' Only apply defaults if Source is blank
        If Trim$(CStr(checkRow.Cells(1, sourceCol).Value)) = "" Then
            ' Set Source to N/A (user-added, not from BOM)
            checkRow.Cells(1, sourceCol).Value = "N/A"
            
            ' Assign next sequential Item number
            SetIfColExistsSingle lo, checkRow, "Master Equipment List Item", GetNextItemNumber(lo)
            
            ' Set defaults for user entries
            SetIfColExistsSingle lo, checkRow, "P&ID Tags", ""
            SetIfColExistsSingle lo, checkRow, "Include in I/O List?", "N"
            SetIfColExistsSingle lo, checkRow, "Include in Utility Load Table?", "N"
            SetIfColExistsSingle lo, checkRow, "Include in Heat Load & Noise Table?", "N"
            SetIfColExistsSingle lo, checkRow, "Removed from BOM", "N"  ' N = active part, not marked for deletion
            SetIfColExistsSingle lo, checkRow, "Notes", ""
        End If
    Next i
    
    Application.EnableEvents = True
End Sub


'===========================================================
' WORKSHEET CHANGE EVENT - HANDLE EDITS TO EXISTING ROWS
'===========================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' This catches when user manually types in a cell
    ' Useful if they paste data or type in Source column
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sourceCol As Long
    Dim changedRow As Range
    Dim rowRange As Range

    Set ws = Me
    
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_MASTER)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    ' Only care if the change touched the table body
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    
    sourceCol = GetTableColIndex(lo, "Source")
    If sourceCol = 0 Then Exit Sub
    
    ' Check if user edited the Source column specifically
    Set changedRow = Intersect(Target, lo.ListColumns(sourceCol).DataBodyRange)
    If changedRow Is Nothing Then Exit Sub
    
    ' Apply defaults to rows where Source was just set to blank or deleted
    Application.EnableEvents = False
    
    For Each rowRange In Intersect(Target.EntireRow, lo.DataBodyRange).Rows
        If Trim$(CStr(rowRange.Cells(1, sourceCol).Value)) = "" Then
            ' Set Source to N/A
            rowRange.Cells(1, sourceCol).Value = "N/A"
            
            ' Assign Item number if blank
            SetIfColExistsSingle lo, rowRange, "Master Equipment List Item", GetNextItemNumber(lo)
            
            ' Set defaults
            SetIfColExistsSingle lo, rowRange, "P&ID Tags", ""
            SetIfColExistsSingle lo, rowRange, "Include in I/O List?", "N"
            SetIfColExistsSingle lo, rowRange, "Include in Utility Load Table?", "N"
            SetIfColExistsSingle lo, rowRange, "Include in Heat Load & Noise Table?", "N"
            SetIfColExistsSingle lo, rowRange, "Removed from BOM", "N"
            SetIfColExistsSingle lo, rowRange, "Notes", ""
        End If
    Next rowRange
    
    Application.EnableEvents = True
End Sub


'===========================================================
' WORKSHEET ACTIVATE EVENT - REFRESH ROW COUNT
'===========================================================
Private Sub Worksheet_Activate()
    ' Refresh row count when sheet is activated
    ' This handles cases where sync added rows while sheet wasn't active
    Dim lo As ListObject
    
    On Error Resume Next
    Set lo = Me.ListObjects(TABLE_MASTER)
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then
            m_LastRowCount = 0
        Else
            m_LastRowCount = lo.DataBodyRange.Rows.Count
        End If
    End If
End Sub


'===========================================================
' GET NEXT ITEM NUMBER
'===========================================================
Private Function GetNextItemNumber(lo As ListObject) As Long
    ' Find the highest existing Item number and return next sequential number
    Dim itemCol As Long
    Dim r As Long
    Dim maxItem As Long
    Dim itemVal As Variant
    
    itemCol = GetTableColIndex(lo, "Master Equipment List Item")
    If itemCol = 0 Then
        GetNextItemNumber = 1
        Exit Function
    End If
    
    maxItem = 0
    
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            itemVal = lo.DataBodyRange.Cells(r, itemCol).Value
            If IsNumeric(itemVal) Then
                If CLng(itemVal) > maxItem Then
                    maxItem = CLng(itemVal)
                End If
            End If
        Next r
    End If
    
    ' Return next number in sequence
    GetNextItemNumber = maxItem + 1
End Function


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

