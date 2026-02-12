Option Explicit

'===========================================================
' SYNC DESTINATION SHEETS ONLY USING ONE BUTTON
'===========================================================
Public Sub Apply_Changes()
    ' Sync destination tables based on current Include settings
    ' Does NOT re-sync from BOMs
    
    Global_Unprotect
    
    ' Sync to destination tables only
    Sync_All_Destinations
    
    ' Apply formatting and locking
    Post_Sync_Format
    
    Global_Protect
    
    MsgBox "Destination tables updated successfully!", vbInformation
End Sub


'===========================================================
' SYNC ALL DESTINATIONS (CALLED BY THE MASTER EQUIPMENT SYNC)
'===========================================================
Public Sub Sync_All_Destinations()
    ' Sync from Master Equipment List to all destination tables
    Sync_IO_List
    Sync_PID_Tag_List
    Sync_Utility_Load
    Sync_Heat_Noise
End Sub


'===========================================================
' SYNC TO IO LIST
'===========================================================
Private Sub Sync_IO_List()
    Dim loMaster As ListObject
    Dim loIO As ListObject
    Dim wsMaster As Worksheet, wsIO As Worksheet
    
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set wsIO = ThisWorkbook.Worksheets(SHEET_IO_LIST)
    Set loMaster = wsMaster.ListObjects(TABLE_MASTER)
    Set loIO = wsIO.ListObjects(TABLE_IO_LIST)
    
    ' Get column indices
    Dim mPartCol As Long, mItemCol As Long, mIncludeCol As Long, mElecTagCol As Long
    mPartCol = GetTableColIndex(loMaster, "Part Number")
    mItemCol = GetTableColIndex(loMaster, "Master Equipment List Item")
    mIncludeCol = GetTableColIndex(loMaster, "Include in I/O List?")
    mElecTagCol = GetTableColIndex(loMaster, "ELEC Tags")
    
    If mPartCol = 0 Or mItemCol = 0 Or mIncludeCol = 0 Or mElecTagCol = 0 Then Exit Sub
    
    ' Build dictionary of what should be in IO List
    ' Key: Item# & "|" & ELEC Tag
    Dim dictShouldExist As Object
    Set dictShouldExist = CreateObject("Scripting.Dictionary")
    dictShouldExist.CompareMode = vbTextCompare
    
    Dim masterRow As Range
    Dim includeVal As String
    Dim itemNum As String
    Dim partNum As String
    Dim elecTags As String
    Dim tagArray() As String
    Dim tag As Variant
    Dim i As Long
    Dim key As Variant
    
    ' Process Master Equipment List
    If TableHasRows(loMaster) Then
        For Each masterRow In loMaster.DataBodyRange.Rows
            includeVal = UCase$(Trim$(CStr(masterRow.Cells(1, mIncludeCol).Value)))
            
            If includeVal = "Y" Then
                itemNum = CStr(masterRow.Cells(1, mItemCol).Value)
                partNum = Trim$(CStr(masterRow.Cells(1, mPartCol).Value))
                elecTags = Trim$(CStr(masterRow.Cells(1, mElecTagCol).Value))
                
                If elecTags <> "" Then
                    ' Split tags and create entry for each
                    tagArray = SplitTags(elecTags)
                    For i = LBound(tagArray) To UBound(tagArray)
                        tag = Trim$(tagArray(i))
                        If tag <> "" Then
                            key = itemNum & "|" & tag
                            If Not dictShouldExist.Exists(key) Then
                                dictShouldExist.Add key, masterRow
                            End If
                        End If
                    Next i
                End If
            End If
        Next masterRow
    End If
    
    ' Get IO List column indices
    Dim ioItemCol As Long, ioTagCol As Long
    ioItemCol = GetTableColIndex(loIO, "Master Equipment List Item")
    ioTagCol = GetTableColIndex(loIO, "ELEC Tag")
    
    If ioItemCol = 0 Or ioTagCol = 0 Then Exit Sub
    
    ' PASS 1: Remove rows that shouldn't exist (Include = N)
    Dim r As Long
    Dim ioRow As Range
    Dim ioKey As String
    
    If TableHasRows(loIO) Then
        For r = loIO.DataBodyRange.Rows.Count To 1 Step -1
            Set ioRow = loIO.DataBodyRange.Rows(r)
            itemNum = CStr(ioRow.Cells(1, ioItemCol).Value)
            tag = Trim$(CStr(ioRow.Cells(1, ioTagCol).Value))
            
            ' Skip manual entries (blank item number)
            If itemNum <> "" Then
                ioKey = itemNum & "|" & tag
                
                If dictShouldExist.Exists(ioKey) Then
                    ' Update existing row
                    UpdateIORow loIO, ioRow, dictShouldExist(ioKey), CStr(tag)
                    dictShouldExist.Remove ioKey
                Else
                    ' Delete row (Include changed to N or tag removed)
                    ioRow.Delete
                End If
            End If
        Next r
    End If
    
    ' PASS 2: Add new rows
    Dim newRow As ListRow
    For Each key In dictShouldExist.Keys
        Set masterRow = dictShouldExist(key)
        
        ' Extract tag from key
        Dim keyParts() As String
        keyParts = Split(key, "|")
        If UBound(keyParts) >= 1 Then
            tag = keyParts(1)
            Set newRow = loIO.ListRows.Add
            UpdateIORow loIO, newRow.Range, masterRow, CStr(tag)
        End If
    Next key
    
    SortByItem loIO
    
End Sub


'===========================================================
' UPDATE IO LIST ROW
'===========================================================
Private Sub UpdateIORow(loIO As ListObject, ioRow As Range, masterRow As Range, elecTag As String)
    Dim loMaster As ListObject
    Set loMaster = ThisWorkbook.Worksheets(SHEET_MASTER).ListObjects(TABLE_MASTER)
    
    ' Copy from master (will overwrite user changes to Description/Notes per requirements)
    SetCellValue loIO, ioRow, "Master Equipment List Item", GetCellValue(masterRow, "Master Equipment List Item", loMaster)
    SetCellValue loIO, ioRow, "Manufacturer", GetCellValue(masterRow, "Manufacturer", loMaster)
    SetCellValue loIO, ioRow, "Part Number", GetCellValue(masterRow, "Part Number", loMaster)
    SetCellValue loIO, ioRow, "ELEC Tag", elecTag
    
    ' Only set Description and Notes if they're currently blank (preserve user edits)
    If Trim$(GetCellValueFromTable(loIO, ioRow, "Description")) = "" Then
        SetCellValue loIO, ioRow, "Description", GetCellValue(masterRow, "Functional Description", loMaster)
    End If
    
    If Trim$(GetCellValueFromTable(loIO, ioRow, "Notes")) = "" Then
        SetCellValue loIO, ioRow, "Notes", GetCellValue(masterRow, "Notes", loMaster)
    End If
    
    ' Leave I/O Controller and I/O Point blank (user entry only)
End Sub


'===========================================================
' SYNC TO P&ID TAG LIST
'===========================================================
Private Sub Sync_PID_Tag_List()
    Dim loMaster As ListObject
    Dim loPID As ListObject
    Dim wsMaster As Worksheet, wsPID As Worksheet
    
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set wsPID = ThisWorkbook.Worksheets(SHEET_PID_TAG_LIST)
    Set loMaster = wsMaster.ListObjects(TABLE_MASTER)
    Set loPID = wsPID.ListObjects(TABLE_PID_TAG_LIST)
    
    ' Get column indices
    Dim mPartCol As Long, mItemCol As Long, mPIDTagCol As Long
    mPartCol = GetTableColIndex(loMaster, "Part Number")
    mItemCol = GetTableColIndex(loMaster, "Master Equipment List Item")
    mPIDTagCol = GetTableColIndex(loMaster, "P&ID Tags")
    
    If mPartCol = 0 Or mItemCol = 0 Or mPIDTagCol = 0 Then Exit Sub
    
    ' Build dictionary of what should exist
    ' Key: Item# & "|" & P&ID Tag
    Dim dictShouldExist As Object
    Set dictShouldExist = CreateObject("Scripting.Dictionary")
    dictShouldExist.CompareMode = vbTextCompare
    
    Dim masterRow As Range
    Dim itemNum As String
    Dim pidTags As String
    Dim tagArray() As String
    Dim tag As Variant
    Dim i As Long
    Dim key As Variant
    
    ' Process Master Equipment List
    If TableHasRows(loMaster) Then
        For Each masterRow In loMaster.DataBodyRange.Rows
            pidTags = Trim$(CStr(masterRow.Cells(1, mPIDTagCol).Value))
            
            If pidTags <> "" Then
                itemNum = CStr(masterRow.Cells(1, mItemCol).Value)
                
                ' Split tags and create entry for each
                tagArray = SplitTags(pidTags)
                For i = LBound(tagArray) To UBound(tagArray)
                    tag = Trim$(tagArray(i))
                    If tag <> "" Then
                        key = itemNum & "|" & tag
                        If Not dictShouldExist.Exists(key) Then
                            dictShouldExist.Add key, masterRow
                        End If
                    End If
                Next i
            End If
        Next masterRow
    End If
    
    ' Get P&ID List column indices
    Dim pidItemCol As Long, pidTagCol As Long
    pidItemCol = GetTableColIndex(loPID, "Master Equipment List Item")
    pidTagCol = GetTableColIndex(loPID, "P&ID Tag")
    
    If pidItemCol = 0 Or pidTagCol = 0 Then Exit Sub
    
    ' PASS 1: Update existing or remove obsolete
    Dim r As Long
    Dim pidRow As Range
    Dim pidKey As String
    
    If TableHasRows(loPID) Then
        For r = loPID.DataBodyRange.Rows.Count To 1 Step -1
            Set pidRow = loPID.DataBodyRange.Rows(r)
            itemNum = CStr(pidRow.Cells(1, pidItemCol).Value)
            tag = Trim$(CStr(pidRow.Cells(1, pidTagCol).Value))
            
            ' Skip manual entries (blank item number)
            If itemNum <> "" Then
                pidKey = itemNum & "|" & tag
                
                If dictShouldExist.Exists(pidKey) Then
                    ' Update existing row
                    UpdatePIDRow loPID, pidRow, dictShouldExist(pidKey), CStr(tag)
                    dictShouldExist.Remove pidKey
                Else
                    ' Delete row (tag removed from master)
                    pidRow.Delete
                End If
            End If
        Next r
    End If
    
    ' PASS 2: Add new rows
    Dim newRow As ListRow
    For Each key In dictShouldExist.Keys
        Set masterRow = dictShouldExist(key)
        
        ' Extract tag from key
        Dim keyParts() As String
        keyParts = Split(key, "|")
        If UBound(keyParts) >= 1 Then
            tag = keyParts(1)
            Set newRow = loPID.ListRows.Add
            UpdatePIDRow loPID, newRow.Range, masterRow, CStr(tag)
        End If
    Next key
    
    ' Sort by Instrument/Equipment, then Loop/Equipment Number
    SortPIDTagList loPID
End Sub


'===========================================================
' UPDATE P&ID TAG LIST ROW
'===========================================================
Private Sub UpdatePIDRow(loPID As ListObject, pidRow As Range, masterRow As Range, pidTag As String)
    Dim loMaster As ListObject
    Set loMaster = ThisWorkbook.Worksheets(SHEET_MASTER).ListObjects(TABLE_MASTER)
    
    ' Copy from master
    SetCellValue loPID, pidRow, "Master Equipment List Item", GetCellValue(masterRow, "Master Equipment List Item", loMaster)
    SetCellValue loPID, pidRow, "Manufacturer", GetCellValue(masterRow, "Manufacturer", loMaster)
    SetCellValue loPID, pidRow, "Part Number", GetCellValue(masterRow, "Part Number", loMaster)
    SetCellValue loPID, pidRow, "P&ID Tag", pidTag
    
    ' Only set Description and Notes if they're currently blank (preserve user edits)
    If Trim$(GetCellValueFromTable(loPID, pidRow, "Description")) = "" Then
        SetCellValue loPID, pidRow, "Description", GetCellValue(masterRow, "Functional Description", loMaster)
    End If
    
    If Trim$(GetCellValueFromTable(loPID, pidRow, "Notes")) = "" Then
        SetCellValue loPID, pidRow, "Notes", GetCellValue(masterRow, "Notes", loMaster)
    End If
    
    ' Leave Loop/Equipment Number and Instrument/Equipment blank (user entry)
End Sub


'===========================================================
' SORT P&ID TAG LIST
'===========================================================
Private Sub SortPIDTagList(loPID As ListObject)
    If Not TableHasRows(loPID) Then Exit Sub
    If loPID.Range.Parent.FilterMode Then Exit Sub
    
    With loPID.Sort
        .SortFields.Clear
        
        ' Sort by Master Equipment List Item first
        On Error Resume Next
        If GetTableColIndex(loPID, "Master Equipment List Item") > 0 Then
            .SortFields.Add key:=loPID.ListColumns("Master Equipment List Item").Range, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, _
                            DataOption:=xlSortNormal
        End If
        
        ' Then by Instrument/Equipment
        If GetTableColIndex(loPID, "Instrument / Equipment?") > 0 Then
            .SortFields.Add key:=loPID.ListColumns("Instrument / Equipment?").Range, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, _
                            DataOption:=xlSortTextAsNumbers
        End If
        
        ' Then by Loop/Equipment Number
        If GetTableColIndex(loPID, "Loop / Equipment Number") > 0 Then
            .SortFields.Add key:=loPID.ListColumns("Loop / Equipment Number").Range, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, _
                            DataOption:=xlSortTextAsNumbers
        End If
        
        .Header = xlYes
        .Apply
    End With
End Sub


'===========================================================
' SYNC TO UTILITY LOAD TABLE
'===========================================================
Private Sub Sync_Utility_Load()
    Dim loMaster As ListObject
    Dim loUtil As ListObject
    Dim wsMaster As Worksheet, wsUtil As Worksheet
    
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set wsUtil = ThisWorkbook.Worksheets(SHEET_UTILITY_LOAD)
    Set loMaster = wsMaster.ListObjects(TABLE_MASTER)
    Set loUtil = wsUtil.ListObjects(TABLE_UTILITY_LOAD)
    
    ' Get column indices
    Dim mItemCol As Long, mIncludeCol As Long
    mItemCol = GetTableColIndex(loMaster, "Master Equipment List Item")
    mIncludeCol = GetTableColIndex(loMaster, "Include in Utility Load Table?")
    
    If mItemCol = 0 Or mIncludeCol = 0 Then Exit Sub
    
    ' Build dictionary of what should exist (keyed by Item#)
    Dim dictShouldExist As Object
    Set dictShouldExist = CreateObject("Scripting.Dictionary")
    dictShouldExist.CompareMode = vbTextCompare
    
    Dim masterRow As Range
    Dim includeVal As String
    Dim itemNum As String
    
    If TableHasRows(loMaster) Then
        For Each masterRow In loMaster.DataBodyRange.Rows
            includeVal = UCase$(Trim$(CStr(masterRow.Cells(1, mIncludeCol).Value)))
            
            If includeVal = "Y" Then
                itemNum = CStr(masterRow.Cells(1, mItemCol).Value)
                If Not dictShouldExist.Exists(itemNum) Then
                    dictShouldExist.Add itemNum, masterRow
                End If
            End If
        Next masterRow
    End If
    
    ' Get Utility Load column indices
    Dim utilItemCol As Long
    utilItemCol = GetTableColIndex(loUtil, "Master Equipment List Item")
    
    If utilItemCol = 0 Then Exit Sub
    
    ' PASS 1: Update existing or remove obsolete
    Dim r As Long
    Dim utilRow As Range
    
    If TableHasRows(loUtil) Then
        For r = loUtil.DataBodyRange.Rows.Count To 1 Step -1
            Set utilRow = loUtil.DataBodyRange.Rows(r)
            itemNum = CStr(utilRow.Cells(1, utilItemCol).Value)
            
            If dictShouldExist.Exists(itemNum) Then
                ' Update existing row
                UpdateUtilityLoadRow loUtil, utilRow, dictShouldExist(itemNum)
                dictShouldExist.Remove itemNum
            Else
                ' Delete row (Include changed to N)
                utilRow.Delete
            End If
        Next r
    End If
    
    ' PASS 2: Add new rows
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In dictShouldExist.Keys
        Set masterRow = dictShouldExist(key)
        Set newRow = loUtil.ListRows.Add
        UpdateUtilityLoadRow loUtil, newRow.Range, masterRow
    Next key
    
    SortByItem loUtil
    
End Sub


'===========================================================
' UPDATE UTILITY LOAD ROW
'===========================================================
Private Sub UpdateUtilityLoadRow(loUtil As ListObject, utilRow As Range, masterRow As Range)
    Dim loMaster As ListObject
    Set loMaster = ThisWorkbook.Worksheets(SHEET_MASTER).ListObjects(TABLE_MASTER)
    
    ' Copy from master (all tags stay as comma-separated lists)
    SetCellValue loUtil, utilRow, "Master Equipment List Item", GetCellValue(masterRow, "Master Equipment List Item", loMaster)
    SetCellValue loUtil, utilRow, "QTY", GetCellValue(masterRow, "QTY", loMaster)
    SetCellValue loUtil, utilRow, "Manufacturer", GetCellValue(masterRow, "Manufacturer", loMaster)
    SetCellValue loUtil, utilRow, "Part Number", GetCellValue(masterRow, "Part Number", loMaster)
    SetCellValue loUtil, utilRow, "P&ID Tags", GetCellValue(masterRow, "P&ID Tags", loMaster)
    SetCellValue loUtil, utilRow, "ELEC Tags", GetCellValue(masterRow, "ELEC Tags", loMaster)
    SetCellValue loUtil, utilRow, "HYD Tags", GetCellValue(masterRow, "HYD Tags", loMaster)
    SetCellValue loUtil, utilRow, "PNU Tags", GetCellValue(masterRow, "PNU Tags", loMaster)
    
    ' Leave all load data columns blank (user entry)
End Sub


'===========================================================
' SYNC TO HEAT LOAD & NOISE TABLE
'===========================================================
Private Sub Sync_Heat_Noise()
    Dim loMaster As ListObject
    Dim loHeat As ListObject
    Dim wsMaster As Worksheet, wsHeat As Worksheet
    
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set wsHeat = ThisWorkbook.Worksheets(SHEET_HEAT_NOISE)
    Set loMaster = wsMaster.ListObjects(TABLE_MASTER)
    Set loHeat = wsHeat.ListObjects(TABLE_HEAT_NOISE)
    
    ' Get column indices
    Dim mItemCol As Long, mIncludeCol As Long
    mItemCol = GetTableColIndex(loMaster, "Master Equipment List Item")
    mIncludeCol = GetTableColIndex(loMaster, "Include in Heat Load & Noise Table?")
    
    If mItemCol = 0 Or mIncludeCol = 0 Then Exit Sub
    
    ' Build dictionary of what should exist (keyed by Item#)
    Dim dictShouldExist As Object
    Set dictShouldExist = CreateObject("Scripting.Dictionary")
    dictShouldExist.CompareMode = vbTextCompare
    
    Dim masterRow As Range
    Dim includeVal As String
    Dim itemNum As String
    
    If TableHasRows(loMaster) Then
        For Each masterRow In loMaster.DataBodyRange.Rows
            includeVal = UCase$(Trim$(CStr(masterRow.Cells(1, mIncludeCol).Value)))
            
            If includeVal = "Y" Then
                itemNum = CStr(masterRow.Cells(1, mItemCol).Value)
                If Not dictShouldExist.Exists(itemNum) Then
                    dictShouldExist.Add itemNum, masterRow
                End If
            End If
        Next masterRow
    End If
    
    ' Get Heat & Noise column indices
    Dim heatItemCol As Long
    heatItemCol = GetTableColIndex(loHeat, "Master Equipment List Item")
    
    If heatItemCol = 0 Then Exit Sub
    
    ' PASS 1: Update existing or remove obsolete
    Dim r As Long
    Dim heatRow As Range
    
    If TableHasRows(loHeat) Then
        For r = loHeat.DataBodyRange.Rows.Count To 1 Step -1
            Set heatRow = loHeat.DataBodyRange.Rows(r)
            itemNum = CStr(heatRow.Cells(1, heatItemCol).Value)
            
            If dictShouldExist.Exists(itemNum) Then
                ' Update existing row
                UpdateHeatNoiseRow loHeat, heatRow, dictShouldExist(itemNum)
                dictShouldExist.Remove itemNum
            Else
                ' Delete row (Include changed to N)
                heatRow.Delete
            End If
        Next r
    End If
    
    ' PASS 2: Add new rows
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In dictShouldExist.Keys
        Set masterRow = dictShouldExist(key)
        Set newRow = loHeat.ListRows.Add
        UpdateHeatNoiseRow loHeat, newRow.Range, masterRow
    Next key
    
    SortByItem loHeat
    
End Sub


'===========================================================
' UPDATE HEAT & NOISE ROW
'===========================================================
Private Sub UpdateHeatNoiseRow(loHeat As ListObject, heatRow As Range, masterRow As Range)
    Dim loMaster As ListObject
    Set loMaster = ThisWorkbook.Worksheets(SHEET_MASTER).ListObjects(TABLE_MASTER)
    
    ' Copy from master (all tags stay as comma-separated lists, include QTY)
    SetCellValue loHeat, heatRow, "Master Equipment List Item", GetCellValue(masterRow, "Master Equipment List Item", loMaster)
    SetCellValue loHeat, heatRow, "QTY", GetCellValue(masterRow, "QTY", loMaster)
    SetCellValue loHeat, heatRow, "Manufacturer", GetCellValue(masterRow, "Manufacturer", loMaster)
    SetCellValue loHeat, heatRow, "Part Number", GetCellValue(masterRow, "Part Number", loMaster)
    SetCellValue loHeat, heatRow, "P&ID Tags", GetCellValue(masterRow, "P&ID Tags", loMaster)
    SetCellValue loHeat, heatRow, "ELEC Tags", GetCellValue(masterRow, "ELEC Tags", loMaster)
    SetCellValue loHeat, heatRow, "HYD Tags", GetCellValue(masterRow, "HYD Tags", loMaster)
    SetCellValue loHeat, heatRow, "PNU Tags", GetCellValue(masterRow, "PNU Tags", loMaster)
    
    ' Leave heat/noise data columns blank (user entry)
End Sub


'===========================================================
' HELPER FUNCTIONS
'===========================================================
Private Function SplitTags(tagString As String) As String()
    ' Split comma-separated tags and return array
    Dim tags() As String
    
    If InStr(tagString, ",") > 0 Then
        tags = Split(tagString, ",")
    Else
        ReDim tags(0 To 0)
        tags(0) = tagString
    End If
    
    SplitTags = tags
End Function


Private Function GetCellValueFromTable(lo As ListObject, rowRange As Range, colName As String) As String
    Dim idx As Long
    idx = GetTableColIndex(lo, colName)
    If idx > 0 Then
        GetCellValueFromTable = CStr(rowRange.Cells(1, idx).Value)
    Else
        GetCellValueFromTable = ""
    End If
End Function


Private Sub SortByItem(lo As ListObject)
    ' Sort table by Master Equipment List Item
    If Not TableHasRows(lo) Then Exit Sub
    If lo.Range.Parent.FilterMode Then Exit Sub
    
    On Error Resume Next
    With lo.Sort
        .SortFields.Clear
        If GetTableColIndex(lo, "Master Equipment List Item") > 0 Then
            .SortFields.Add key:=lo.ListColumns("Master Equipment List Item").Range, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, _
                            DataOption:=xlSortNormal
        End If
        .Header = xlYes
        .Apply
    End With
    On Error GoTo 0
End Sub

