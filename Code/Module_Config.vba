Option Explicit

'===========================================================
' SHEET NAME CONSTANTS
'===========================================================
Public Const SHEET_MASTER As String = "Master Equipment List"
Public Const SHEET_IO_LIST As String = "IO List"
Public Const SHEET_PID_TAG_LIST As String = "P&ID Tag List"
Public Const SHEET_PID_LOOPS As String = "P&ID Loops"
Public Const SHEET_UTILITY_LOAD As String = "Utility Load Table"
Public Const SHEET_HEAT_NOISE As String = "Heat Load & Noise"
Public Const SHEET_ALL_BOM As String = "All_BOM_Parts"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_SYNC_LOG As String = "SYNC_Log"

'===========================================================
' TABLE NAME CONSTANTS
'===========================================================
Public Const TABLE_MASTER As String = "Master_Equipment"
Public Const TABLE_IO_LIST As String = "IO_List"
Public Const TABLE_PID_TAG_LIST As String = "PID_Tag_List"
Public Const TABLE_PID_LOOPS As String = "PID_Loops"
Public Const TABLE_UTILITY_LOAD As String = "Utility_Load"
Public Const TABLE_HEAT_NOISE As String = "Heat_Noise"
Public Const TABLE_ALL_BOM As String = "All_BOM_Parts"
Public Const TABLE_SYNC_LOG As String = "Sync_Log"

'===========================================================
' ALL_BOM_PARTS COLUMN NAMES
' These drive the column lookups in source BOM files
' (Except "BOM Source" which is added by the query)
'===========================================================
Public Const BOM_COL_SOURCE As String = "BOM Source"
Public Const BOM_COL_ITEM As String = "Item"
Public Const BOM_COL_PHASE As String = "Phase"
Public Const BOM_COL_MANUFACTURER As String = "Manufacturer"
Public Const BOM_COL_PART_NUMBER As String = "Part Number"
Public Const BOM_COL_ASSY_QTY As String = "Assy QTY"
Public Const BOM_COL_QTY As String = "QTY"
Public Const BOM_COL_NEED_QTY As String = "Need QTY"
Public Const BOM_COL_LOC As String = "LOC"
Public Const BOM_COL_LOC_DESC As String = "LOC Description"

'===========================================================
' PATH HELPER FUNCTION
'===========================================================
Public Function GetProjectPath() As String
    Dim dropboxStatus As String
    Dim projFolder As String
    Dim userPath As String
    
    dropboxStatus = UCase$(Trim$(CStr(ThisWorkbook.Names("DropboxStatus").RefersToRange.Value)))
    projFolder = Trim$(CStr(ThisWorkbook.Names("Proj_Folder").RefersToRange.Value))
    
    If dropboxStatus = "Y" Then
        ' Build Dropbox path: C:\Users\USERNAME\[user's path]
        userPath = "C:\Users\" & Environ$("Username") & "\"
        GetProjectPath = userPath & projFolder
    Else
        ' Use direct path (full path entered by user)
        GetProjectPath = projFolder
    End If
End Function

'===========================================================
' WORKSHEET GET FUNCTIONS
'===========================================================
Public Function wsMaster() As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)
End Function

Public Function wsIOList() As Worksheet
    Set wsIOList = ThisWorkbook.Worksheets(SHEET_IO_LIST)
End Function

Public Function wsPIDTagList() As Worksheet
    Set wsPIDTagList = ThisWorkbook.Worksheets(SHEET_PID_TAG_LIST)
End Function

Public Function wsPIDLoops() As Worksheet
    Set wsPIDLoops = ThisWorkbook.Worksheets(SHEET_PID_LOOPS)
End Function

Public Function wsUtilityLoad() As Worksheet
    Set wsUtilityLoad = ThisWorkbook.Worksheets(SHEET_UTILITY_LOAD)
End Function

Public Function wsHeatNoise() As Worksheet
    Set wsHeatNoise = ThisWorkbook.Worksheets(SHEET_HEAT_NOISE)
End Function

Public Function wsAllBOM() As Worksheet
    Set wsAllBOM = ThisWorkbook.Worksheets(SHEET_ALL_BOM)
End Function

Public Function wsConfig() As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
End Function

Public Function wsSyncLog() As Worksheet
    Set wsSyncLog = ThisWorkbook.Worksheets(SHEET_SYNC_LOG)
End Function

