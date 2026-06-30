VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Workbook_Open()
    Dim sPage$
    
    If Not check_integrity() Then
        MsgBox "failed integrity check !", vbOKOnly + vbCritical, APP_TITLE
        Exit Sub
    End If
    
    Call init_sheet_vars
    Call populate_pages_listbox
    Worksheets(MINDMAP_SHEET).Activate
    
    sPage = current_page()
    Call select_page_in_list(sPage, True)
End Sub
