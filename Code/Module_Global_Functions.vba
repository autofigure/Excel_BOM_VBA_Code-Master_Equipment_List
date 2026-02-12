Option Explicit

Public Job_Number As String


'==========================================================
' GLOBAL UNPROTECT FUNCTION
'==========================================================
Public Sub Global_Unprotect()
    Dim wsActive As Worksheet
    On Error Resume Next
    Set wsActive = ActiveSheet

    With ThisWorkbook
        .Worksheets(SHEET_MASTER).Unprotect ""
        .Worksheets(SHEET_IO_LIST).Unprotect ""
        .Worksheets(SHEET_PID_TAG_LIST).Unprotect ""
        .Worksheets(SHEET_PID_LOOPS).Unprotect ""
        .Worksheets(SHEET_UTILITY_LOAD).Unprotect ""
        .Worksheets(SHEET_HEAT_NOISE).Unprotect ""
        .Worksheets(SHEET_ALL_BOM).Unprotect ""
        .Worksheets(SHEET_CONFIG).Unprotect ""
        '.Worksheets(SHEET_SYNC_LOG).Unprotect ""  ' Commented - sheet doesn't exist yet
    End With

    ' go back to where user was
    If Not wsActive Is Nothing Then wsActive.Activate
End Sub


'==========================================================
' GLOBAL PROTECT FUNCTION
'==========================================================
Public Sub Global_Protect()
    Dim wsActive As Worksheet
    Dim ws As Worksheet
    Dim rng As Range

    On Error Resume Next
    Set wsActive = ActiveSheet
    On Error GoTo 0

    ' -------------------------------
    ' 1) Master Equipment List
    ' -------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MASTER)
    If Not ws Is Nothing Then
        ' Unlock project info cells (merged cells)
        ws.Range("C2:C7").MergeArea.Locked = False
        ws.Range("E6").MergeArea.Locked = False
        
        ' Apply row-level locking to Master Equipment List table
        Dim loMaster As ListObject
        Set loMaster = ws.ListObjects(TABLE_MASTER)
        
        If Not loMaster Is Nothing And Not loMaster.DataBodyRange Is Nothing Then
            Dim sourceCol As Long, removedCol As Long
            Dim r As Long, c As Long
            Dim rowRange As Range
            Dim srcVal As String, removedVal As String
            Dim colName As String
            Dim isLockedSource As Boolean
            
            ' Columns that stay editable even for BOM rows
            Dim editableCols As Variant
            editableCols = Array("P&ID Tags", "Include in I/O List?", _
                                "Include in Utility Load Table?", _
                                "Include in Heat Load & Noise Table?", _
                                "User Entries", "Notes")
            
            sourceCol = GetTableColIndex(loMaster, "Source")
            removedCol = GetTableColIndex(loMaster, "Removed from BOM")
            
            If sourceCol > 0 And removedCol > 0 Then
                For r = 1 To loMaster.DataBodyRange.Rows.Count
                    Set rowRange = loMaster.DataBodyRange.Rows(r)
                    
                    srcVal = UCase$(Trim$(CStr(rowRange.Cells(1, sourceCol).Value)))
                    removedVal = UCase$(Trim$(CStr(rowRange.Cells(1, removedCol).Value)))
                    
                    ' 1) Fully unlock removed rows
                    If removedVal = "Y" Then
                        rowRange.Locked = False
                    
                    ' 2) Manual rows stay editable
                    ElseIf srcVal = "MAN" Then
                        rowRange.Locked = False
                    
                    Else
                        ' 3) Real BOM rows - lock most columns except editable ones
                        isLockedSource = (InStr(srcVal, "ELEC") > 0 Or InStr(srcVal, "HYD") > 0 Or _
                                          InStr(srcVal, "PNU") > 0 Or InStr(srcVal, "MECH") > 0)
                        
                        For c = 1 To loMaster.ListColumns.Count
                            colName = loMaster.ListColumns(c).Name
                            
                            If isLockedSource Then
                                If IsInArray(colName, editableCols) Then
                                    rowRange.Cells(1, c).Locked = False
                                Else
                                    rowRange.Cells(1, c).Locked = True
                                End If
                            Else
                                rowRange.Cells(1, c).Locked = False
                            End If
                        Next c
                    End If
                Next r
            End If
        End If
        
        ' Now protect the sheet
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 2) IO List
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_IO_LIST)
    If Not ws Is Nothing Then
        Dim loIO As ListObject
        Set loIO = ws.ListObjects(TABLE_IO_LIST)
        
        If Not loIO Is Nothing And Not loIO.DataBodyRange Is Nothing Then
            Dim lockedColsIO As Variant
            lockedColsIO = Array("Master Equipment List Item", "Manufacturer", "Part Number", "ELEC Tag")
            
            Dim rIO As Long, cIO As Long
            Dim rowRangeIO As Range
            Dim colNameIO As String
            
            For rIO = 1 To loIO.DataBodyRange.Rows.Count
                Set rowRangeIO = loIO.DataBodyRange.Rows(rIO)
                
                For cIO = 1 To loIO.ListColumns.Count
                    colNameIO = loIO.ListColumns(cIO).Name
                    
                    If IsInArray(colNameIO, lockedColsIO) Then
                        rowRangeIO.Cells(1, cIO).Locked = True
                    Else
                        rowRangeIO.Cells(1, cIO).Locked = False
                    End If
                Next cIO
            Next rIO
        End If
        
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 3) P&ID Tag List
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_PID_TAG_LIST)
    If Not ws Is Nothing Then
        Dim loPID As ListObject
        Set loPID = ws.ListObjects(TABLE_PID_TAG_LIST)
        
        If Not loPID Is Nothing And Not loPID.DataBodyRange Is Nothing Then
            Dim lockedColsPID As Variant
            lockedColsPID = Array("Master Equipment List Item", "Manufacturer", "Part Number", "P&ID Tag")
            
            Dim itemColPID As Long
            Dim rPID As Long, cPID As Long
            Dim rowRangePID As Range
            Dim colNamePID As String
            Dim itemValPID As String
            
            itemColPID = GetTableColIndex(loPID, "Master Equipment List Item")
            
            If itemColPID > 0 Then
                For rPID = 1 To loPID.DataBodyRange.Rows.Count
                    Set rowRangePID = loPID.DataBodyRange.Rows(rPID)
                    itemValPID = Trim$(CStr(rowRangePID.Cells(1, itemColPID).Value))
                    
                    ' Manual entries (blank item#) are fully editable
                    If itemValPID = "" Then
                        rowRangePID.Locked = False
                    Else
                        ' Lock columns from master
                        For cPID = 1 To loPID.ListColumns.Count
                            colNamePID = loPID.ListColumns(cPID).Name
                            
                            If IsInArray(colNamePID, lockedColsPID) Then
                                rowRangePID.Cells(1, cPID).Locked = True
                            Else
                                rowRangePID.Cells(1, cPID).Locked = False
                            End If
                        Next cPID
                    End If
                Next rPID
            End If
        End If
        
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 4) P&ID Loops
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_PID_LOOPS)
    If Not ws Is Nothing Then
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 5) Utility Load Table
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_UTILITY_LOAD)
    If Not ws Is Nothing Then
        Dim loUtil As ListObject
        Set loUtil = ws.ListObjects(TABLE_UTILITY_LOAD)
        
        If Not loUtil Is Nothing And Not loUtil.DataBodyRange Is Nothing Then
            Dim lockedColsUtil As Variant
            lockedColsUtil = Array("Master Equipment List Item", "QTY", "Manufacturer", "Part Number", _
                                   "P&ID Tags", "ELEC Tags", "HYD Tags", "PNU Tags")
            
            Dim rUtil As Long, cUtil As Long
            Dim rowRangeUtil As Range
            Dim colNameUtil As String
            
            For rUtil = 1 To loUtil.DataBodyRange.Rows.Count
                Set rowRangeUtil = loUtil.DataBodyRange.Rows(rUtil)
                
                For cUtil = 1 To loUtil.ListColumns.Count
                    colNameUtil = loUtil.ListColumns(cUtil).Name
                    
                    If IsInArray(colNameUtil, lockedColsUtil) Then
                        rowRangeUtil.Cells(1, cUtil).Locked = True
                    Else
                        rowRangeUtil.Cells(1, cUtil).Locked = False
                    End If
                Next cUtil
            Next rUtil
        End If
        
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 6) Heat Load & Noise
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_HEAT_NOISE)
    If Not ws Is Nothing Then
        Dim loHeat As ListObject
        Set loHeat = ws.ListObjects(TABLE_HEAT_NOISE)
        
        If Not loHeat Is Nothing And Not loHeat.DataBodyRange Is Nothing Then
            Dim lockedColsHeat As Variant
            lockedColsHeat = Array("Master Equipment List Item", "QTY", "Manufacturer", "Part Number", _
                                   "P&ID Tags", "ELEC Tags", "HYD Tags", "PNU Tags")
            
            Dim rHeat As Long, cHeat As Long
            Dim rowRangeHeat As Range
            Dim colNameHeat As String
            
            For rHeat = 1 To loHeat.DataBodyRange.Rows.Count
                Set rowRangeHeat = loHeat.DataBodyRange.Rows(rHeat)
                
                For cHeat = 1 To loHeat.ListColumns.Count
                    colNameHeat = loHeat.ListColumns(cHeat).Name
                    
                    If IsInArray(colNameHeat, lockedColsHeat) Then
                        rowRangeHeat.Cells(1, cHeat).Locked = True
                    Else
                        rowRangeHeat.Cells(1, cHeat).Locked = False
                    End If
                Next cHeat
            Next rHeat
        End If
        
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowSorting:=True, _
                       AllowInsertingRows:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 7) All_BOM_Parts
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_ALL_BOM)
    If Not ws Is Nothing Then
        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If
    
    ' -------------------------------
    ' 8) Config
    ' -------------------------------
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    If Not ws Is Nothing Then
        ' Unlock named cells first
        Set rng = ThisWorkbook.Names("ThisWorkbookFolder").RefersToRange
        If Not rng Is Nothing Then rng.Locked = False

        Set rng = ThisWorkbook.Names("DropboxStatus").RefersToRange
        If Not rng Is Nothing Then rng.Locked = False

        If Not ws.ProtectContents Then
            ws.Protect Password:="", _
                       DrawingObjects:=False, _
                       Contents:=True, _
                       Scenarios:=False, _
                       AllowFiltering:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
        End If
    End If

    ' -------------------------------
    ' 9) SYNC_Log (TEMPORARILY DISABLED)
    ' -------------------------------
    ' TODO: Create SYNC_Log sheet with Sync_Log table containing:
    '       Timestamp, Status, Details, User columns
    '
    'Set ws = ThisWorkbook.Worksheets(SHEET_SYNC_LOG)
    'If Not ws Is Nothing Then
    '    If Not ws.ProtectContents Then
    '        ws.Protect Password:="", _
    '                   DrawingObjects:=True, _
    '                   Contents:=True, _
    '                   Scenarios:=True, _
    '                   AllowFiltering:=False, _
    '                   AllowFormattingColumns:=True, _
    '                   AllowFormattingRows:=True
    '    End If
    'End If

    ' put the user back (safely handle protected sheets)
    On Error Resume Next
    If Not wsActive Is Nothing Then wsActive.Activate
    On Error GoTo 0
End Sub


'===========================================================
' TABLE HELPER FUNCTIONS
'===========================================================
Public Function GetTableColIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            GetTableColIndex = lc.Index
            Exit Function
        End If
    Next lc
    GetTableColIndex = 0
End Function


Public Function TableHasRows(lo As ListObject) As Boolean
    If lo Is Nothing Then
        TableHasRows = False
    ElseIf lo.DataBodyRange Is Nothing Then
        TableHasRows = False
    ElseIf lo.DataBodyRange.Rows.Count = 0 Then
        TableHasRows = False
    Else
        TableHasRows = True
    End If
End Function


Public Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(val, arr(i), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


'===========================================================
' BOM DETECTION
'===========================================================
Public Function DetectMissingBOMs(loSrc As ListObject, expected As Variant) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim srcBOMCol As Long
    srcBOMCol = GetTableColIndex(loSrc, BOM_COL_SOURCE)
    If srcBOMCol = 0 Then
        DetectMissingBOMs = ""
        Exit Function
    End If

    Dim lr As ListRow
    For Each lr In loSrc.ListRows
        seen(Trim$(CStr(lr.Range.Cells(1, srcBOMCol).Value))) = True
    Next lr

    Dim missing As String
    Dim i As Long, label As String
    For i = LBound(expected) To UBound(expected)
        label = expected(i)
        If Not seen.Exists(label) Then
            missing = missing & Job_Number & "-" & label & " BOM (INTERNAL)" & vbCrLf
        End If
    Next i

    DetectMissingBOMs = missing
End Function


'===========================================================
' LOGGING
'===========================================================
Public Sub WriteSyncLog(ByVal statusText As String, ByVal detailText As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim newRow As ListRow

    Set ws = ThisWorkbook.Worksheets(SHEET_SYNC_LOG)
    Set lo = ws.ListObjects(TABLE_SYNC_LOG)

    Set newRow = lo.ListRows.Add(1)
    With newRow.Range
        .Cells(1, 1).Value = Now
        .Cells(1, 2).Value = statusText
        .Cells(1, 3).Value = detailText
        .Cells(1, 4).Value = Environ$("Username")
    End With
    lo.ListColumns(1).Range.NumberFormat = "mm/dd/yyyy hh:mm:ss"
End Sub


'===========================================================
' String formatting helpers
'===========================================================
Public Function PadRight(ByVal txt As String, ByVal width As Long) As String
    If Len(txt) >= width Then
        PadRight = Left$(txt, width)
    Else
        PadRight = txt & Space(width - Len(txt))
    End If
End Function

Public Function PadLeft(ByVal txt As String, ByVal width As Long) As String
    If Len(txt) >= width Then
        PadLeft = Right$(txt, width)
    Else
        PadLeft = Space(width - Len(txt)) & txt
    End If
End Function

Public Function PadBoth(ByVal txt As String, ByVal width As Long) As String
    Dim padTotal As Long, padLeftLen As Long
    padTotal = width - Len(txt)
    If padTotal <= 0 Then
        PadBoth = txt
    Else
        padLeftLen = padTotal \ 2
        PadBoth = Space(padLeftLen) & txt & Space(padTotal - padLeftLen)
    End If
End Function


'===========================================================
' TABLE DATA HELPERS
'===========================================================
Public Function GetCellValue(rowRange As Range, colName As String, lo As ListObject) As String
    Dim idx As Long
    idx = GetTableColIndex(lo, colName)
    If idx > 0 Then
        GetCellValue = CStr(rowRange.Cells(1, idx).Value)
    Else
        GetCellValue = ""
    End If
End Function


Public Sub SetCellValue(lo As ListObject, rowRange As Range, colName As String, val As Variant)
    Dim idx As Long
    idx = GetTableColIndex(lo, colName)
    If idx > 0 Then
        rowRange.Cells(1, idx).Value = val
    End If
End Sub


'===========================================================
' POWER QUERY CREATION
'===========================================================
Public Sub Create_BOM_Query(Optional showSuccessMessage As Boolean = True)
    ' Dynamically creates the Power Query for All_BOM_Parts based on
    ' Proj_Folder, Proj_Number, and DropboxStatus cells
    ' Uses All_BOM_Parts table column names to drive source BOM lookups
    
    Dim projPath As String
    Dim projNumber As String
    Dim mCode As String
    Dim conn As WorkbookConnection
    Dim loAllBOM As ListObject
    Dim columnList As String
    Dim col As ListColumn
    
    ' Get project info
    projPath = GetProjectPath()  ' Uses Dropbox logic from Module_Config
    On Error Resume Next
    projNumber = Trim$(CStr(ThisWorkbook.Names("Proj_Number").RefersToRange.Value))
    On Error GoTo 0
    
    If projNumber = "" Then
        MsgBox "Proj_Number named cell is empty. Please fill it in before creating the query.", vbExclamation
        Exit Sub
    End If
    
    If projPath = "" Then
        MsgBox "Proj_Folder named cell is empty. Please fill it in before creating the query.", vbExclamation
        Exit Sub
    End If
    
    ' Get All_BOM_Parts table to read its column names
    On Error Resume Next
    Set loAllBOM = wsAllBOM.ListObjects(TABLE_ALL_BOM)
    On Error GoTo 0
    
    If loAllBOM Is Nothing Then
        MsgBox "All_BOM_Parts table not found. Please create the table first.", vbCritical
        Exit Sub
    End If
    
    ' Build column list from All_BOM_Parts table (excluding BOM Source which we add)
    columnList = ""
    For Each col In loAllBOM.ListColumns
        If col.Name <> "BOM Source" Then
            If columnList <> "" Then columnList = columnList & ", "
            columnList = columnList & """" & col.Name & """"
        End If
    Next col
    
    ' Build M code for Power Query - RESILIENT VERSION (handles missing BOMs)
    ' Split into parts to avoid VBA's 24 line continuation limit
    Dim mCode1 As String, mCode2 As String, mCode3 As String
    
    mCode1 = "let" & vbCrLf & _
            "    // ELEC BOM" & vbCrLf & _
            "    Source_ELEC = try Excel.Workbook(File.Contents(""" & projPath & "\Job Documentation\Electrical\" & projNumber & "-ELEC BOM (INTERNAL).xlsx""), null, true) otherwise null," & vbCrLf & _
            "    BOM_ELEC = if Source_ELEC = null then #table({}, {}) else Source_ELEC{[Item=""BOM"",Kind=""Table""]}[Data]," & vbCrLf & _
            "    Add_Source_ELEC = if Table.RowCount(BOM_ELEC) = 0 then #table({}, {}) else Table.AddColumn(BOM_ELEC, ""BOM Source"", each ""ELEC"")," & vbCrLf & _
            vbCrLf & _
            "    // HYD BOM" & vbCrLf & _
            "    Source_HYD = try Excel.Workbook(File.Contents(""" & projPath & "\Job Documentation\Hydraulic\" & projNumber & "-HYD BOM (INTERNAL).xlsx""), null, true) otherwise null," & vbCrLf & _
            "    BOM_HYD = if Source_HYD = null then #table({}, {}) else Source_HYD{[Item=""BOM"",Kind=""Table""]}[Data]," & vbCrLf & _
            "    Add_Source_HYD = if Table.RowCount(BOM_HYD) = 0 then #table({}, {}) else Table.AddColumn(BOM_HYD, ""BOM Source"", each ""HYD"")," & vbCrLf
    
    mCode2 = "    // PNU BOM" & vbCrLf & _
            "    Source_PNU = try Excel.Workbook(File.Contents(""" & projPath & "\Job Documentation\Pneumatic\" & projNumber & "-PNU BOM (INTERNAL).xlsx""), null, true) otherwise null," & vbCrLf & _
            "    BOM_PNU = if Source_PNU = null then #table({}, {}) else Source_PNU{[Item=""BOM"",Kind=""Table""]}[Data]," & vbCrLf & _
            "    Add_Source_PNU = if Table.RowCount(BOM_PNU) = 0 then #table({}, {}) else Table.AddColumn(BOM_PNU, ""BOM Source"", each ""PNU"")," & vbCrLf & _
            vbCrLf & _
            "    // MECH BOM" & vbCrLf & _
            "    Source_MECH = try Excel.Workbook(File.Contents(""" & projPath & "\Job Documentation\Mechanical\" & projNumber & "-MECH BOM (INTERNAL).xlsx""), null, true) otherwise null," & vbCrLf & _
            "    BOM_MECH = if Source_MECH = null then #table({}, {}) else Source_MECH{[Item=""BOM"",Kind=""Table""]}[Data]," & vbCrLf & _
            "    Add_Source_MECH = if Table.RowCount(BOM_MECH) = 0 then #table({}, {}) else Table.AddColumn(BOM_MECH, ""BOM Source"", each ""MECH"")," & vbCrLf
    
    mCode3 = "    // Combine only non-empty tables" & vbCrLf & _
            "    AllTables = {Add_Source_ELEC, Add_Source_HYD, Add_Source_PNU, Add_Source_MECH}," & vbCrLf & _
            "    NonEmptyTables = List.Select(AllTables, each Table.RowCount(_) > 0)," & vbCrLf & _
            "    Combined = if List.Count(NonEmptyTables) = 0 then #table({}, {}) else Table.Combine(NonEmptyTables)," & vbCrLf & _
            "    SelectColumns = if Table.RowCount(Combined) = 0 then Combined else Table.SelectColumns(Combined,{""BOM Source"", " & columnList & "})" & vbCrLf & _
            "in" & vbCrLf & _
            "    SelectColumns"
    
    ' Concatenate all parts
    mCode = mCode1 & vbCrLf & mCode2 & vbCrLf & mCode3
    
    ' Delete existing query if it exists
    On Error Resume Next
    For Each conn In ThisWorkbook.Connections
        If conn.Name = "All_BOM_Parts" Then
            conn.Delete
        End If
    Next conn
    ThisWorkbook.Queries("All_BOM_Parts").Delete
    On Error GoTo 0
    
    ' Create new query
    On Error GoTo QueryError
    With ThisWorkbook.Queries
        .Add Name:="All_BOM_Parts", Formula:=mCode
    End With
    
    If showSuccessMessage Then
        MsgBox "Power Query created successfully!" & vbCrLf & vbCrLf & _
               "Query will load data from:" & vbCrLf & projPath & vbCrLf & vbCrLf & _
               "Columns being pulled: " & Replace(columnList, """", "") & vbCrLf & vbCrLf & _
               "Now run 'Run_Master_Equipment_Sync' to populate your tables.", vbInformation
    End If
    Exit Sub
    
QueryError:
    MsgBox "Error creating Power Query:" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Check that:" & vbCrLf & _
           "1. BOM files exist at: " & projPath & "\Job Documentation\..." & vbCrLf & _
           "2. Each BOM file has a table named 'BOM'" & vbCrLf & _
           "3. BOM tables have columns matching All_BOM_Parts table" & vbCrLf & _
           "4. File names match pattern: " & projNumber & "-[BOM] BOM (INTERNAL).xlsx" & vbCrLf & _
           "5. Folders named: Electrical, Hydraulic, Pneumatic, Mechanical", vbCritical
End Sub

