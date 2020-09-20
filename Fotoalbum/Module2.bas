Attribute VB_Name = "Module2"
Option Explicit

Public Sub HandleDTGMultiSelect(dtg As DataGrid, y As Single)
'---------------------------------------------------------------------------------------
' Procedure : HandleDTGMultiSelect
' DateTime : 30/03/2005 16:23
' Author : michelle
' Purpose : The Microsoft Datagrid control handles the use of the CTRL key to make
' multiple selections, however, the user cannot use SHIFT in combination with CTRL
' to select batches of records in the expected method. This function enables the
' use of SHIFT. All records between the CURRENT bookmark (i.e. most recently selected)
' and the previous bookmark (i.e. the previous bookmark made) are selected.
'MUST BE CALLED FROM MOUSEUP EVENT
'---------------------------------------------------------------------------------------
    
    'On Error GoTo HandleDTGMultiSelect_Error
    Dim i As Integer
    Dim j As Integer
    Dim dir As Integer
    Dim LastSelectedRow As Long
    Dim FirstSelectedRow As Long
    Dim SecondLastSelRow As Long
    
    With dtg
        'if the user has made a previous selection, get the LAST ADDED bookmark
        'in the selected bookmarks collection
        If dtg.SelBookmarks.Count > 0 Then
            
            'get the row the user last clicked on
            LastSelectedRow = .RowContaining(y) + GetRowFromBookmark(dtg, .FirstRow)
            
            
            'get the ROW that was selected before the last one (which is represened
            'by LastSelectedRow. We want to select the records between the LastSelectedRow (the one
            'they last clicked on, and the one that was selected before that.
            SecondLastSelRow = GetPreviouslySelectedRow(dtg)
            
            'work out the first row then clicked on. This is the "current" row, i.e. the GetBookmark
            'method will return bookmarks relative to this position.
            FirstSelectedRow = GetFirstSelectedRow(dtg)
            
            'if the last selected bookmark ROW is BEFORE the LastSelectedRow the user clicked,
            'the loop needs to go FORWARD. Otherwise, the loop needs to go BACKWARDS
            If SecondLastSelRow < LastSelectedRow Then
                dir = 1
                'work out the position of the first row to be selected (between
                'our two targets) relative to teh current row (dtg.row)
                j = SecondLastSelRow - FirstSelectedRow
            Else
                dir = -1
                j = SecondLastSelRow - FirstSelectedRow
            End If
            
            For i = SecondLastSelRow To LastSelectedRow Step dir
                dtg.SelBookmarks.Add dtg.GetBookmark(j)
                j = j + dir
            Next
            
        End If
    End With
    
    On Error GoTo 0
    
exitproc:
    
    Exit Sub
    
HandleDTGMultiSelect_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleDTGMultiSelect of Module mdlDatagridFunctions", vbExclamation
    GoTo exitproc
End Sub

Public Function GetRowFromBookmark(dtg As DataGrid, Bookmark As Variant)
'---------------------------------------------------------------------------------------
' Procedure : GetRowFromBookmark
' DateTime : 13/10/2005 14:52
' Author : mod.au.smith
' Purpose : The rows in the datagrid start at 0 to dtg.approxcount. The bookmarks
' run consequequitevly, but don't start at 0. So, get the last bookmark in the grid.
' this is equivalent to approxcount. The bookmark that represents line 0 is last bookmark
'- approxcount
'---------------------------------------------------------------------------------------
    Dim i As Integer
    Dim MaxBookMark As Long
    Dim MinBookMark As Long
    Dim MaxRecords As Long
    
    MaxRecords = dtg.ApproxCount - 1
    
    On Error Resume Next 'we're expecting an error, because the getbookmark method is
    'relative to the current row position (which we don't know). We're just trying to
    'get the last bookmark in the grid.
    For i = 0 To dtg.ApproxCount
        MaxBookMark = dtg.GetBookmark(i)
    Next
    On Error GoTo 0
    
    MinBookMark = MaxBookMark - MaxRecords
    
    'now convert our known bookmark to a row number
    GetRowFromBookmark = Bookmark - MinBookMark
End Function

Public Function GetFirstSelectedRow(dtg As DataGrid) As Long
'---------------------------------------------------------------------------------------
' Procedure : GetFirstSelectedRow
' DateTime : 13/10/2005 13:40
' Author : mod.au.smith
' Purpose : You would think that dtg.row returns the first selected row, but you would
' be wrong. If the user has SCROLLED teh datagrid, the datagrid will return -1. The
' safest way to work out which row the user first clicked, is to loop through all the
' rows in the datagrid, until you find one with a bookmark that matches the FIRST bookmark
' in the selbookmarks collection.
'---------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rownum As Integer
    Dim SelRowIndex
    
    Select Case dtg.SelBookmarks.Count
        Case Is > 0
            SelRowIndex = 0
    End Select
    
    GetFirstSelectedRow = GetRowFromBookmark(dtg, dtg.SelBookmarks.Item(SelRowIndex))
End Function

Private Function GetPreviouslySelectedRow(dtg As DataGrid) As Long
'---------------------------------------------------------------------------------------
' Procedure : GetPreviouslySelectedRow
' DateTime : 13/10/2005 13:40
' Author : mod.au.smith
' Purpose : The second-to-last selected row is the one that is in the second-to-last selbookmarks index.
' Cannot use dtg.row because the current row is NOT always the one that was last selected
' (the row is the one with the triangle next to it in the recordselectors). This function
' loops through each ROW in the datagrid until it finds the row with the bookmark that
' matches the second-to-last one in the selbookmarks collection. This is the row number that is
' returned.
'---------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rownum As Integer
    Dim SelRowIndex
    
    Select Case dtg.SelBookmarks.Count
        Case Is < 2
            SelRowIndex = 0
        Case Is >= 2
            SelRowIndex = dtg.SelBookmarks.Count - 2
    End Select
    
    GetPreviouslySelectedRow = GetRowFromBookmark(dtg, dtg.SelBookmarks.Item(SelRowIndex))
End Function

