Attribute VB_Name = "import"
Option Explicit

Const APP_NAME = "SV mindmap"

Dim mbConfirmed As Boolean


Public Sub on_click_import()
    Dim fd As FileDialog, sPath$, vConfirm As VbMsgBoxResult
    
    Call init_sheet_vars
    
    If Not mbConfirmed Then
        vConfirm = MsgBox("this will replace all data ", vbYesNo + vbCritical, APP_TITLE & "import")
        If vConfirm <> vbYes Then Exit Sub
        mbConfirmed = True
    End If
    
       

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Excel file to import from"
        .Filters.Clear
        .Filters.Add "mindmap files", "*.xlsm"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"

        If .Show <> -1 Then
            MsgBox "nothing selected to import", vbInformation + vbOKOnly, APP_TITLE
            Exit Sub 'cancelled
        End If
        sPath = .SelectedItems(1)
    End With
        
    '--- prevent selecting current workbook
    If StrComp(sPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "You cannot select the currently open workbook.", vbExclamation + vbOKOnly, APP_TITLE
        Call on_click_import
        Exit Sub
    End If
    
    Dim sFname$
    sFname = Replace(sPath, "/", "\")
    sFname = Mid$(sFname, InStrRev(sFname, "\") + 1)
    If StrComp(sFname, ThisWorkbook.name, vbTextCompare) = 0 Then
        MsgBox "You cannot select a filename with the same name as the currently open workbook.", vbExclamation + vbOKOnly, APP_TITLE
        Call on_click_import
        Exit Sub
    End If

    
    'open the file
    Dim oWbImport As Workbook, bOK As Boolean
    On Error GoTo cleanup1
        Application.ScreenUpdating = False
        Set oWbImport = Workbooks.Open(Filename:=sPath, ReadOnly:=True)
        If oWbImport Is Nothing Then
            MsgBox "unable to open file", vbExclamation + vbOKOnly, APP_TITLE
            Exit Sub
        End If
cleanup1:
        Application.ScreenUpdating = True
    On Error GoTo 0
    
    'validate and perform import
    bOK = True
    On Error GoTo cleanup2
        oWbImport.Windows(1).Visible = False
        
        If Not is_supported_mindmap(oWbImport) Then
            MsgBox "spreadsheet contents not recognised", vbOKOnly + vbExclamation, APP_TITLE
        Else
            Application.ScreenUpdating = False
            Call perform_import(oWbImport)
        End If

cleanup2:
        If Err.Number <> 0 Then MsgBox "import failed: " & Err.Description, vbOKOnly + vbExclamation, APP_TITLE
        oWbImport.Close False
        Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub


Private Function is_supported_mindmap(poImportWB As Workbook) As Boolean
    is_supported_mindmap = False
    
    Dim sVersion$, sApp$
    On Error Resume Next
        'check the name and version
        sVersion = poImportWB.Names(NAME_VERSION).RefersToRange.Value
        If Err.Number <> 0 Then Exit Function
        If sVersion <> 1 Then Exit Function
        
        sApp = poImportWB.Names(NAME_APP).RefersToRange.Value
        If Err.Number <> 0 Then Exit Function
        If sApp <> APP_NAME Then Exit Function
        
        'check that the expected sheets are there
        Dim oSheet As Worksheet
        Set oSheet = poImportWB.Worksheets(NODES_SHEET)
        If Err.Number <> 0 Then Exit Function
        Set oSheet = poImportWB.Worksheets(EDGES_SHEET)
        If Err.Number <> 0 Then Exit Function
        Set oSheet = poImportWB.Worksheets(PAGES_SHEET)
        If Err.Number <> 0 Then Exit Function
        Set oSheet = poImportWB.Worksheets(EXTRAS_SHEET)
        If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    'check that sheets exist
    is_supported_mindmap = True
End Function


Private Sub perform_import(poWB As Workbook)
    Call import_from(poWB, NODES_SHEET)
    Call import_from(poWB, EDGES_SHEET)
    Call import_from(poWB, PAGES_SHEET)
    Call import_from(poWB, EXTRAS_SHEET)
    
    Call delete_all_shapes
    
    'refesh the page list
    Call populate_pages_listbox
    
    'get the currently selected page
    Dim sPage$
    sPage = poWB.Names(CURRENT_PAGE_RANGE).RefersToRange.Value
    If sPage <> "" Then Call select_page_in_list(sPage, True)
    
End Sub

Private Sub import_from(poWB As Workbook, psSheet$)
    Dim oThis As Worksheet, oThat As Worksheet
    
    Set oThis = Worksheets(psSheet)
    Set oThat = poWB.Worksheets(psSheet)
    
    'erase everything
    oThis.Cells.Clear
       

    'Copy everything (values, formats, column widths, etc.)
    With oThat
        .UsedRange.Copy
        With oThis.Range("A1")
            .PasteSpecial xlPasteAll
            .PasteSpecial xlPasteColumnWidths
        End With
    End With
    
    Application.CutCopyMode = False


End Sub
