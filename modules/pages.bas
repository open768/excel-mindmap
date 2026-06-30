Attribute VB_Name = "pages"

Option Explicit

Public Const COL_PAGE_ID = 1
Const COL_PAGE_NAME = 2
Const COL_PAGE_ROOT = 3
Public Const CURRENT_PAGE_RANGE = "current_page"
Public Const PAGES_LIST = "lst_pages"

Public Function get_page_row(psName$) As Long

    Dim s_name_col$, vResult
    
    s_name_col = ColumnLetter(COL_PAGE_NAME)
    vResult = Application.match(psName$, moPagesSheet.Columns(s_name_col), EXACT_MATCH)
    get_page_row = CLng(vResult)
    
End Function


'######################################################################################
'# events
'######################################################################################


Sub on_click_save()
    Call init_sheet_vars
    Call save_nodes
    Call save_page_edges
    On Error Resume Next
        Call ActiveWorkbook.save
    On Error GoTo 0
End Sub

Public Sub onPageRenameClick()
    Dim sCurrent$, sNew$, sName$
    Call init_sheet_vars
    
    sCurrent = current_page()
    sName = InputBox("rename page", "new name of page", APP_TITLE & sCurrent)
    If sName = "" Then Exit Sub
    If sName = sCurrent Then Exit Sub
    
    'get the row to update
    Dim iRow&
    iRow = get_page_row(sCurrent)
    moPagesSheet.Cells(iRow, COL_PAGE_NAME).Value = sName
    
    'refresh the list
    set_current_page sName
    populate_pages_listbox
End Sub

Public Sub btn_new_page_Click()
    Dim sName$
    sName = InputBox("Name", APP_TITLE & "create a mindmap page")
    If sName = "" Then Exit Sub
    sName = Trim$(sName)
    Call init_sheet_vars
    
    '----------check if duplicate
    Dim iCount&
    iCount = Application.WorksheetFunction.CountIf(moPagesSheet.Columns(COL_PAGE_NAME), sName)
    If iCount > 0 Then
        MsgBox "duplicate name", vbOKOnly + vbCritical, APP_TITLE
        Exit Sub
    End If
    
    '----------create
    Call init_sheet_vars
    Call save_nodes
    Call save_page_edges
    
    Call delete_all_shapes
    Call create_page(sName)
End Sub


Public Sub create_page(psName$)
    '-----------create page
    Dim iRow&, iPageID&, iLast&
    
    
    iRow = get_last_row(moPagesSheet, ColumnLetter(COL_PAGE_ID))
    If iRow < 1 Then iRow = 1
    
    On Error Resume Next
    iLast = moPagesSheet.Cells(iRow, COL_PAGE_ID).Value
    If Err.Number <> 0 Then iLast = 0
    On Error GoTo 0
    
    iPageID = iLast + 1
    iRow = iRow + 1
    moPagesSheet.Cells(iRow, COL_PAGE_ID).Value = iPageID
    moPagesSheet.Cells(iRow, COL_PAGE_NAME).Value = psName
    Range(CURRENT_PAGE_RANGE).Value = psName
    
    
    '-----------create node
    Dim oCentre As tCoordinate, oCell As Range
    Set oCell = moMapSheet.Cells(15, 7)
    oCentre.x = oCell.Left + oCell.Width / 2
    oCentre.y = oCell.Top + oCell.Height / 2
    
    Call create_node(oCentre)
    
    Call populate_pages_listbox
    Call select_page_in_list(psName)
End Sub

Public Sub click_pages_list()
    Dim oList As Shape
    Dim sPage$
    Call init_sheet_vars
    
    Set oList = moMapSheet.Shapes(PAGES_LIST)
    sPage = get_list_selection(moMapSheet, oList)
    Call draw_mindmap(sPage)
End Sub

'######################################################################################
'# current page
'######################################################################################
Public Sub set_current_page(psName$, Optional pbCheck As Boolean = True)
    Dim sCurrent$
    
    sCurrent = current_page()
    If pbCheck Then
        If psName$ = sCurrent Then Exit Sub
        Range(CURRENT_PAGE_RANGE).Value = psName$
    End If
End Sub

Public Function current_page() As String
    current_page = Range(CURRENT_PAGE_RANGE).Value
End Function

Public Function get_current_page_id() As Long
    Dim sPage$
    sPage = current_page()
    
    'lookup page name in worksheet "pages" column "B" and get the value from column "A"
    Dim s_ID_col$, s_name_col$
    s_ID_col = ColumnLetter(COL_PAGE_ID)
    s_name_col = ColumnLetter(COL_PAGE_NAME)
    
    On Error Resume Next
        With moPagesSheet
            get_current_page_id = WorksheetFunction.XLookup(sPage, .Columns(s_name_col), .Columns(s_ID_col))
        End With
        If Err.Number <> 0 Then get_current_page_id = 0
            
    On Error GoTo 0
    
End Function


'######################################################################################
'# page list
'######################################################################################

Public Function get_page_name(piID&) As String
    Dim iLastRow&, iRow&, iRowID&
    
    iLastRow = get_last_row(moPagesSheet, "A")
    
    For iRow = 2 To iLastRow Step 1
        iRowID = moPagesSheet.Cells(iRow, COL_PAGE_ID).Value
        If iRowID = piID Then
            get_page_name = moPagesSheet.Cells(iRow, COL_PAGE_NAME).Value
            Exit Function
        End If
    Next iRow

End Function

Public Sub populate_pages_listbox()
    'get the last row
    Dim iLastRow&, iRow&
    iLastRow = get_last_row(moPagesSheet, "A")
    If iLastRow < 2 Then
        MsgBox "not enough data in pages sheet", vbOKOnly + vbCritical, APP_TITLE
        Exit Sub
    End If
        
    'populate form control
    Dim oList As Shape, sName$
    Set oList = moMapSheet.Shapes(PAGES_LIST)
    With oList.ControlFormat
        .RemoveAllItems

        For iRow = 2 To iLastRow
            sName = moPagesSheet.Range(ColumnLetter(COL_PAGE_NAME) & iRow).Value
            .AddItem sName
        Next iRow
    End With
End Sub

Public Sub select_page_in_list(psName$, Optional pbForce As Boolean = False)
    Dim sCurrent$, sPage$, oPageList As Shape, i&
    Dim bFound As Boolean, bContinue As Boolean
    
    Set oPageList = moMapSheet.Shapes(PAGES_LIST)
    
    bContinue = True
    If Not pbForce Then
        sCurrent = current_page()
        sPage = get_list_selection(moMapSheet, oPageList)
        If (sPage = psName Or sCurrent = psName) Then bContinue = False
    End If
    
    If bContinue Then
        bFound = False
        With oPageList.ControlFormat
            For i = 1 To .ListCount
                If LCase(.List(i)) = LCase(psName) Then
                    .ListIndex = i   ' or use .Selected(i) for multi-select
                    DoEvents
                    click_pages_list
                    bFound = True
                    Exit For
                End If
            Next i
        End With
        
        If Not bFound Then MsgBox "page not found: " & sPage, vbOKOnly + vbExclamation, APP_TITLE
    End If
    
    'draw_mindmap sCurrent
End Sub

