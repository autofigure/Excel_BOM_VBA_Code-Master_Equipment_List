Option Explicit

'===========================================================
' MAIN ENTRY POINT
'===========================================================
Public Sub Run_Master_Equipment_Sync()
    Dim conn As WorkbookConnection
    Dim c As WorkbookConnection
    Dim queryExists As Boolean

    ' pull Job Number (used in missing-BOM message)
    On Error Resume Next
    Job_Number = Trim$(CStr(ThisWorkbook.Names("Job_Number").RefersToRange.Value))
    On Error GoTo 0

    ' 1) fully unlock workbook before we start
    Global_Unprotect

    ' 2) Check if Power Query exists - if not, create it automatically
    queryExists = False
    On Error Resume Next
    queryExists = (ThisWorkbook.Queries("All_BOM_Parts").Name <> "")
    On Error GoTo 0
    
    If Not queryExists Then
        ' Query doesn't exist - create it automatically on first run
        MsgBox "Power Query not found. Creating All_BOM_Parts query automatically..." & vbCrLf & vbCrLf & _
               "This only happens once when setting up a new project.", vbInformation, "First-Time Setup"
        Create_BOM_Query showSuccessMessage:=False  ' Silent mode - we'll show our own message
        
        ' Verify it was created successfully
        queryExists = False
        On Error Resume Next
        queryExists = (ThisWorkbook.Queries("All_BOM_Parts").Name <> "")
        On Error GoTo 0
        
        If Not queryExists Then
            MsgBox "Failed to create Power Query. Please run 'Create_BOM_Query' manually and check:" & vbCrLf & _
                   "1. Proj_Number and Proj_Folder cells are filled in" & vbCrLf & _
                   "2. BOM files exist at the expected location", vbCritical
            Global_Protect
            Exit Sub
        Else
            MsgBox "Power Query created successfully! Continuing with sync...", vbInformation, "Setup Complete"
        End If
    End If

    ' 3) refresh the Power Query that feeds All_BOM_Parts
    Set conn = Nothing
    For Each c In ThisWorkbook.Connections
        If InStr(1, c.Name, "All_BOM_Parts", vbTextCompare) > 0 Then
            Set conn = c
            Exit For
        End If
    Next c

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If Not conn Is Nothing Then
        conn.Refresh
        Application.CalculateUntilAsyncQueriesDone
    Else
        ' fallback if name wasn't found
        ThisWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' 4) run the actual sync
    Sync_Master_Equipment_List

    ' 5) sync to destination tables
    Sync_All_Destinations

    ' 6) normalize / lock / shade tables
    Post_Sync_Format

    ' 7) lock workbook again
    Global_Protect

    MsgBox "Master Equipment List sync completed successfully!", vbInformation
End Sub


'===========================================================
' SYNC BOM ? MASTER EQUIPMENT LIST
'===========================================================
Private Sub Sync_Master_Equipment_List()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim loSrc As ListObject, loDst As ListObject
    Dim msg As String
    
    Set wb = ThisWorkbook
    Set wsSrc = wsAllBOM
    Set wsDst = wsMaster
    Set loSrc = wsSrc.ListObjects(TABLE_ALL_BOM)
    Set loDst = wsDst.ListObjects(TABLE_MASTER)
    
    msg = ""
    
    ' --- Check for missing BOMs ---
    Dim missingBOMs As String
    missingBOMs = DetectMissingBOMs(loSrc, Array("ELEC", "HYD", "PNU", "MECH"))
    If Len(missingBOMs) > 0 Then
        msg = msg & PadRight("WARNING:", 10) & "The following BOMs were empty or NOT found:" & vbCrLf & missingBOMs & vbCrLf
    End If
    
    ' --- Source must have rows ---
    If Not TableHasRows(loSrc) Then
        MsgBox msg & vbCrLf & PadRight("WARNING:", 20) & "The BOM Query returned 0 rows." & vbCrLf & vbCrLf & "No changes made.", vbExclamation
        WriteSyncLog "WARNING", "All_BOM_Parts returned 0 rows. No changes made."
        Exit Sub
    End If
    
    ' --- Get column indices ---
    Dim srcPartCol As Long, srcBOMCol As Long, srcMfgCol As Long
    srcPartCol = GetTableColIndex(loSrc, BOM_COL_PART_NUMBER)
    srcBOMCol = GetTableColIndex(loSrc, BOM_COL_SOURCE)
    srcMfgCol = GetTableColIndex(loSrc, BOM_COL_MANUFACTURER)
    
    If srcPartCol = 0 Or srcBOMCol = 0 Then
        MsgBox "SYNC ABORTED: Required columns missing in All_BOM_Parts.", vbCritical
        WriteSyncLog "ABORTED", "Required columns missing in All_BOM_Parts."
        Exit Sub
    End If
    
    ' --- Build source dictionary keyed by Part Number ---
    ' Structure: dictSrc(partNumber) = Collection of BOM occurrences
    Dim dictSrc As Object
    Set dictSrc = CreateObject("Scripting.Dictionary")
    dictSrc.CompareMode = vbTextCompare
    
    Dim srcRow As ListRow
    Dim partKey As String
    Dim bomSource As String
    Dim bomData As Object
    Dim i As Long
    
    ' Collect all BOM data per part number
    For Each srcRow In loSrc.ListRows
        partKey = Trim$(CStr(srcRow.Range.Cells(1, srcPartCol).Value))
        If partKey <> "" Then
            bomSource = Trim$(CStr(srcRow.Range.Cells(1, srcBOMCol).Value))
            
            If Not dictSrc.Exists(partKey) Then
                Set dictSrc(partKey) = CreateObject("Scripting.Dictionary")
                dictSrc(partKey).CompareMode = vbTextCompare
            End If
            
            ' Store this BOM's data for this part
            Set bomData = CreateObject("Scripting.Dictionary")
            bomData("BOM") = bomSource
            Set bomData("Row") = srcRow
            
            ' Use BOM source as key to handle same part in multiple BOMs
            If Not dictSrc(partKey).Exists(bomSource) Then
                dictSrc(partKey).Add bomSource, bomData
            End If
        End If
    Next srcRow
    
    msg = msg & PadRight("Unique parts found:", 28) & dictSrc.Count & vbCrLf
    
    ' --- Get Master Equipment List column indices ---
    Dim dstItemCol As Long, dstPartCol As Long, dstSourceCol As Long
    Dim dstMfgCol As Long, dstRemovedCol As Long
    dstItemCol = GetTableColIndex(loDst, "Master Equipment List Item")
    dstPartCol = GetTableColIndex(loDst, "Part Number")
    dstSourceCol = GetTableColIndex(loDst, "Source")
    dstMfgCol = GetTableColIndex(loDst, "Manufacturer")
    dstRemovedCol = GetTableColIndex(loDst, "Removed from BOM")
    
    If dstItemCol = 0 Or dstPartCol = 0 Or dstSourceCol = 0 Or dstRemovedCol = 0 Then
        MsgBox "SYNC ABORTED: Required columns missing in Master Equipment List.", vbCritical
        WriteSyncLog "ABORTED", "Required columns missing in Master Equipment List."
        Exit Sub
    End If
    
    ' --- Find next available Item number ---
    Dim nextItemNum As Long
    nextItemNum = GetNextItemNumber(loDst, dstItemCol)
    
    ' --- PASS 1: Update existing master rows ---
    Dim dstRow As Range
    Dim rowCount As Long
    Dim updatedExisting As Long, markedRemoved As Long, addedNew As Long
    updatedExisting = 0: markedRemoved = 0: addedNew = 0
    
    If TableHasRows(loDst) Then
        rowCount = loDst.DataBodyRange.Rows.Count
        Dim r As Long
        For r = 1 To rowCount
            Set dstRow = loDst.DataBodyRange.Rows(r)
            partKey = Trim$(CStr(dstRow.Cells(1, dstPartCol).Value))
            
            Dim sourceVal As String
            sourceVal = UCase$(Trim$(CStr(dstRow.Cells(1, dstSourceCol).Value)))
            
            If sourceVal = "MAN" Then
                ' Manual entry - leave as-is
            ElseIf partKey <> "" Then
                If dictSrc.Exists(partKey) Then
                    ' Part still exists in BOM - update it
                    UpdateMasterRow loDst, dstRow, dictSrc(partKey), loSrc
                    dstRow.Cells(1, dstRemovedCol).Value = "N"
                    dictSrc.Remove partKey
                    updatedExisting = updatedExisting + 1
                Else
                    ' Part not in source anymore
                    dstRow.Cells(1, dstRemovedCol).Value = "Y"
                    markedRemoved = markedRemoved + 1
                End If
            End If
        Next r
    End If
    
    ' --- PASS 2: Add new parts ---
    Dim key As Variant
    Dim newRow As ListRow
    
    For Each key In dictSrc.Keys
        Set newRow = loDst.ListRows.Add
        
        ' Assign Item number
        newRow.Range.Cells(1, dstItemCol).Value = nextItemNum
        nextItemNum = nextItemNum + 1
        
        ' Update with BOM data
        UpdateMasterRow loDst, newRow.Range, dictSrc(key), loSrc
        
        ' Set defaults for new rows
        SetDefaultsForNewMasterRow loDst, newRow.Range
        
        addedNew = addedNew + 1
    Next key
    
    ' --- Optional sort ---
    If loDst.Range.Parent.FilterMode = False Then
        With loDst.Sort
            .SortFields.Clear
            If dstItemCol > 0 Then
                .SortFields.Add key:=loDst.ListColumns("Master Equipment List Item").Range, _
                                SortOn:=xlSortOnValues, Order:=xlAscending, _
                                DataOption:=xlSortTextAsNumbers
            End If
            .Header = xlYes
            .Apply
        End With
    End If
    
    ' --- Update LAST_SYNC_DATE ---
    On Error Resume Next
    With ThisWorkbook.Names("LAST_SYNC_DATE").RefersToRange
        .Value = Now
        .NumberFormat = "mm/dd/yyyy hh:nn"
    End With
    On Error GoTo 0
    
    ' --- Build summary message ---
    msg = msg & PadRight("Parts updated:", 32) & updatedExisting & vbCrLf
    msg = msg & PadRight("Parts added:", 34) & addedNew & vbCrLf
    msg = msg & PadRight("Identified as removed:", 27) & markedRemoved & vbCrLf
    msg = msg & PadRight("Sync Completed:", 30) & Format(Now, "mm/dd/yyyy hh:nn")
    
    WriteSyncLog "SUCCESS", msg
End Sub


'===========================================================
' UPDATE MASTER ROW WITH BOM DATA
'===========================================================
Private Sub UpdateMasterRow(loDst As ListObject, dstRow As Range, bomDict As Object, loSrc As ListObject)
    ' bomDict contains all BOM occurrences for this part
    ' Structure: bomDict(bomSource) = {BOM, Row}
    ' Ownership hierarchy: HYD > PNU > ELEC > MECH
    
    Dim bomKey As Variant
    Dim bomData As Object
    Dim srcRow As ListRow
    Dim sources As String
    Dim totalAssyQty As Long, totalQty As Long, totalNeedQty As Long
    Dim elecTags As String, hydTags As String, pnuTags As String
    Dim bestDesc As String
    Dim descPriority As Long ' 1=MECH, 2=ELEC, 3=PNU, 4=HYD
    Dim mfg As String, partNum As String
    Dim ownerBOM As String
    Dim ownerPriority As Long
    
    sources = ""
    totalAssyQty = 0
    totalQty = 0
    totalNeedQty = 0
    elecTags = ""
    hydTags = ""
    pnuTags = ""
    descPriority = 0
    ownerPriority = 0
    ownerBOM = ""
    
    ' First pass: Determine owner based on hierarchy
    For Each bomKey In bomDict.Keys
        Set bomData = bomDict(bomKey)
        Dim bomSource As String
        bomSource = UCase$(bomData("BOM"))
        
        Dim currentPriority As Long
        currentPriority = 0
        If bomSource = "MECH" Then currentPriority = 1
        If bomSource = "PNU" Then currentPriority = 2
        If bomSource = "HYD" Then currentPriority = 3
        If bomSource = "ELEC" Then currentPriority = 4
        
        If currentPriority > ownerPriority Then
            ownerPriority = currentPriority
            ownerBOM = bomSource
        End If
    Next bomKey
    
    ' Second pass: Process each BOM occurrence
    For Each bomKey In bomDict.Keys
        Set bomData = bomDict(bomKey)
        Set srcRow = bomData("Row")
        bomSource = UCase$(bomData("BOM"))
        
        ' Build source list
        If sources <> "" Then sources = sources & ", "
        sources = sources & bomData("BOM")
        
        ' Get quantities from BOM
        Dim assyQty As Long, qty As Long, needQty As Long
        assyQty = CLng(val(GetCellValue(srcRow.Range, "Assy QTY", loSrc)))
        qty = CLng(val(GetCellValue(srcRow.Range, "QTY", loSrc)))
        needQty = CLng(val(GetCellValue(srcRow.Range, "NEED", loSrc)))
        
        ' If this is the owner BOM, use its quantities
        ' Otherwise, only sum the Need QTY (for non-owners)
        If bomSource = ownerBOM Then
            totalAssyQty = totalAssyQty + assyQty
            totalQty = totalQty + qty
            totalNeedQty = totalNeedQty + needQty
            
            ' Owner BOM provides manufacturer and part number
            mfg = GetCellValue(srcRow.Range, "Manufacturer", loSrc.Parent.ListObjects(TABLE_ALL_BOM))
            partNum = GetCellValue(srcRow.Range, "Part Number", loSrc.Parent.ListObjects(TABLE_ALL_BOM))
        Else
            ' Non-owner: only contribute Need QTY
            totalNeedQty = totalNeedQty + needQty
        End If
        
        ' Collect LOC tags by BOM source (all BOMs contribute their tags)
        Dim locTags As String
        locTags = GetCellValue(srcRow.Range, "LOC", loSrc.Parent.ListObjects(TABLE_ALL_BOM))
        
        If bomSource = "ELEC" Then
            elecTags = AppendTags(elecTags, locTags)
        ElseIf bomSource = "HYD" Then
            hydTags = AppendTags(hydTags, locTags)
        ElseIf bomSource = "PNU" Then
            pnuTags = AppendTags(pnuTags, locTags)
        End If
        
        ' Get description with same priority as ownership: HYD > PNU > ELEC > MECH
        Dim currentDescPriority As Long
        currentDescPriority = 0
        If bomSource = "MECH" Then currentDescPriority = 1
        If bomSource = "ELEC" Then currentDescPriority = 2
        If bomSource = "PNU" Then currentDescPriority = 3
        If bomSource = "HYD" Then currentDescPriority = 4
        
        If currentDescPriority > descPriority Then
            Dim loc As String, locDesc As String
            loc = Trim$(GetCellValue(srcRow.Range, "LOC", loSrc.Parent.ListObjects(TABLE_ALL_BOM)))
            locDesc = Trim$(GetCellValue(srcRow.Range, "LOC Description", loSrc.Parent.ListObjects(TABLE_ALL_BOM)))
            
            ' Only use LOC Description
			If locDesc <> "" Then
				bestDesc = locDesc
			End If
            descPriority = currentDescPriority
        End If
    Next bomKey
    
    ' Write to master row
    SetCellValue loDst, dstRow, "Source", sources
    SetCellValue loDst, dstRow, "Manufacturer", mfg
    SetCellValue loDst, dstRow, "Part Number", partNum
    SetCellValue loDst, dstRow, "Assy QTY", totalAssyQty
    SetCellValue loDst, dstRow, "QTY", totalQty
    SetCellValue loDst, dstRow, "Need QTY", totalNeedQty
    SetCellValue loDst, dstRow, "ELEC Tags", elecTags
    SetCellValue loDst, dstRow, "HYD Tags", hydTags
    SetCellValue loDst, dstRow, "PNU Tags", pnuTags
    SetCellValue loDst, dstRow, "Functional Description", bestDesc
End Sub


'===========================================================
' SET DEFAULTS FOR NEW MASTER ROW
'===========================================================
Private Sub SetDefaultsForNewMasterRow(loDst As ListObject, dstRow As Range)
    SetCellValue loDst, dstRow, "P&ID Tags", ""
    SetCellValue loDst, dstRow, "Include in I/O List?", "N"
    SetCellValue loDst, dstRow, "Include in Utility Load Table?", "N"
    SetCellValue loDst, dstRow, "Include in Heat Load & Noise Table?", "N"
    SetCellValue loDst, dstRow, "Removed from BOM", "N"
    SetCellValue loDst, dstRow, "Notes", ""
    ' Assy QTY, QTY, and Need QTY are set by UpdateMasterRow
End Sub


'===========================================================
' HELPER FUNCTIONS
'===========================================================
Private Function GetNextItemNumber(lo As ListObject, itemCol As Long) As Long
    Dim maxItem As Long
    Dim r As Long
    Dim itemVal As Variant
    
    maxItem = 0
    
    If TableHasRows(lo) Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            itemVal = lo.DataBodyRange.Cells(r, itemCol).Value
            If IsNumeric(itemVal) Then
                If CLng(itemVal) > maxItem Then
                    maxItem = CLng(itemVal)
                End If
            End If
        Next r
    End If
    
    GetNextItemNumber = maxItem + 1
End Function


Private Function AppendTags(existingTags As String, newTags As String) As String
    existingTags = Trim$(existingTags)
    newTags = Trim$(newTags)
    
    If existingTags = "" Then
        AppendTags = newTags
    ElseIf newTags = "" Then
        AppendTags = existingTags
    Else
        AppendTags = existingTags & ", " & newTags
    End If
End Function



