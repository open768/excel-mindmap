Attribute VB_Name = "sheets"
Option Explicit

Public moMapSheet As Worksheet
Public moNodeSheet As Worksheet
Public moPagesSheet As Worksheet
Public moEdgesSheet As Worksheet
Public moExtrasSheet As Worksheet

Public Const MINDMAP_SHEET = "mindmap"
Public Const PAGES_SHEET = "pages"
Public Const NODES_SHEET = "nodes"
Public Const EDGES_SHEET = "edges"
Public Const EXTRAS_SHEET = "extra shapes"

Public Const NAME_BTN_ADD_NODE = "btn_add_node"
Public Const NAME_BTN_OPTIONS = "btn_options"
Public Const NAME_BTN_SEARCH = "btn_search"


Public moReservedControls As Scripting.Dictionary

Public Sub init_sheet_vars()
    If moMapSheet Is Nothing Then Set moMapSheet = Worksheets(MINDMAP_SHEET)
    If moNodeSheet Is Nothing Then Set moNodeSheet = Worksheets(NODES_SHEET)
    If moPagesSheet Is Nothing Then Set moPagesSheet = Worksheets(PAGES_SHEET)
    If moEdgesSheet Is Nothing Then Set moEdgesSheet = Worksheets(EDGES_SHEET)
    If moExtrasSheet Is Nothing Then Set moExtrasSheet = Worksheets(EXTRAS_SHEET)
    
    Call add_reserved_controls
    
End Sub

Private Sub add_reserved_controls()
    Set moReservedControls = New Scripting.Dictionary
    moReservedControls(PAGES_LIST) = 1
    moReservedControls(NAME_BTN_OPTIONS) = 1
    moReservedControls(NAME_BTN_SEARCH) = 1
End Sub



Public Function check_integrity() As Boolean
    
    check_integrity = False
    
    'check sheets exist
    If Not sheet_must_exist(MINDMAP_SHEET) Then Exit Function
    If Not sheet_must_exist(NODES_SHEET) Then Exit Function
    If Not sheet_must_exist(PAGES_SHEET) Then Exit Function
    If Not sheet_must_exist(EDGES_SHEET) Then Exit Function
    If Not sheet_must_exist(EXTRAS_SHEET) Then Exit Function
    
    'check form exists
    Dim fTest As UserForm
    On Error Resume Next
        Set fTest = frm_options
        If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
        
    
    'check controls exist
    Dim oMapSheet As Worksheet
    Set oMapSheet = Worksheets(MINDMAP_SHEET)
    If Not shape_exists(NAME_BTN_OPTIONS, oMapSheet) Then Exit Function
    If Not shape_exists(PAGES_LIST, oMapSheet) Then Exit Function
    If Not shape_exists(NAME_BTN_SEARCH, oMapSheet) Then Exit Function
    
    check_integrity = True
End Function

