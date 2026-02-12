Option Explicit

'===========================================================
' DELETE REMOVED PARTS FROM MASTER EQUIPMENT LIST
'===========================================================
Public Sub Delete_Removed_Parts()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim removedCol As Long
    Dim partCol As Long
    Dim descCol As Long
    Dim mfgCol As Long
    Dim sourceCol As Long
    Dim r As Long
    Dim resp As VbMsgBoxResult
    Dim deletedCount As Long
    Dim msg As String
    Dim showOneByOne As Boolean

    ' 1) unprotect everything first
    Global_Unprotect

    ' grab the table
    Set ws = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set lo = ws.ListObjects(TABLE_MASTER)

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "No rows found in Master Equipment List.", vbInformation
        Global_Protect
        Exit Sub
    End If

    ' find columns we care about
    removedCol = GetTableColIndex(lo, "Removed from BOM")
    sourceCol = GetTableColIndex(lo, "Source")
    partCol = GetTableColIndex(lo, "Part Number")
    descCol = GetTableColIndex(lo, "Description")
    mfgCol = GetTableColIndex(lo, "Manufacturer")

    If removedCol = 0 Then
        MsgBox "'Removed from BOM' column not found.", vbCritical
        Global_Protect
        Exit Sub
    End If

    ' 2) ask the user which mode they want
    resp = MsgBox( _
        "Delete ALL rows where 'Removed from BOM' = 'Y' ?" & vbCrLf & vbCrLf & _
        PadRight("Yes:", 14) & "Delete ALL removed rows" & vbCrLf & _
        PadRight("No:", 13) & "Review each removed row individually" & vbCrLf & _
        PadRight("Cancel:", 11) & "Exit", _
        vbYesNoCancel + vbQuestion, _
        "Delete Removed Parts")

    If resp = vbCancel Then
        Global_Protect
        Exit Sub
    ElseIf resp = vbYes Then
        showOneByOne = False
    Else
        showOneByOne = True
    End If

    deletedCount = 0

    ' 3) go bottom-up to delete safely
    For r = lo.DataBodyRange.Rows.Count To 1 Step -1
        Dim rowRange As Range
        Dim removedVal As String

        Set rowRange = lo.DataBodyRange.Rows(r)
        removedVal = UCase$(Trim$(CStr(rowRange.Cells(1, removedCol).Value)))

        If removedVal = "Y" Then
            If showOneByOne Then
                ' build a little info string to help user decide
                msg = "Remove this row?" & vbCrLf & vbCrLf

                If sourceCol > 0 Then msg = msg & PadRight("Source:", 20) & CStr(rowRange.Cells(1, sourceCol).Value) & vbCrLf
                If mfgCol > 0 Then msg = msg & PadRight("Manufacturer:", 20) & CStr(rowRange.Cells(1, mfgCol).Value) & vbCrLf
                If partCol > 0 Then msg = msg & PadRight("Part Number:", 21) & CStr(rowRange.Cells(1, partCol).Value) & vbCrLf
                If descCol > 0 Then msg = msg & PadRight("Description:", 23) & CStr(rowRange.Cells(1, descCol).Value) & vbCrLf

                msg = msg & vbCrLf & "(Yes = Delete,  No = Keep,  Cancel = Exit)"

                resp = MsgBox(msg, vbYesNoCancel + vbQuestion, "Delete this removed row?")

                If resp = vbYes Then
                    rowRange.Delete
                    deletedCount = deletedCount + 1
                ElseIf resp = vbNo Then
                    ' keep, do nothing
                ElseIf resp = vbCancel Then
                    Exit For    ' stop reviewing more rows
                End If
            Else
                ' bulk delete mode
                rowRange.Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next r

    ' 4) re-apply your table-level rules (shading, locking, etc.)
    Post_Sync_Format

    ' 5) re-protect workbook
    Global_Protect

    ' 6) tell user what happened
    MsgBox "Delete removed parts completed." & vbCrLf & _
           PadRight("Rows deleted:", 20) & deletedCount, vbInformation
End Sub


