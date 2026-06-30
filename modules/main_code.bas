Attribute VB_Name = "main_code"
Option Explicit

Const COMMENT_TYPE = "comment"
Const LINK_TYPE = "link"

Public Const CONFIRM_PASS = "diqwy4vo,yqkyer"
Public Const SHAPE_OVAL = msoShapeOval
Public Const SHAPE_RECT = msoShapeRoundedRectangle

Const EXCLUDED_ZONE = "A1:E20"

'####################################################################################################
'# click cell
'####################################################################################################
Sub click_cell(poTarget As Range)
    Call init_sheet_vars
    
    Dim oSheet As Worksheet
    Set oSheet = poTarget.Worksheet
    If Not oSheet.name = MINDMAP_SHEET Then Exit Sub
        
    'remove extra button
    remove_all_buttons
    
    'dont let range between a1 and D20
    Dim oIntersect As Range
    Set oIntersect = Application.Intersect(poTarget, oSheet.Range(EXCLUDED_ZONE))
    If Not oIntersect Is Nothing Then Exit Sub
        
    add_add_node_button poTarget
End Sub



'####################################################################################################
'# node button
'####################################################################################################


Public Sub remove_add_node_button()
    moMapSheet.Shapes(NAME_BTN_ADD_NODE).Delete
End Sub

Sub add_add_node_button(poTarget As Range)
    'add a button to create a node
    Dim oCentre As tCoordinate
    With poTarget
        oCentre.x = .Left + .Width / 2
        oCentre.y = .Top + .Height / 2
    End With

    Dim oButton As Shape
    With oCentre
        Set oButton = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x - ADD_NODE_SIZE / 2, .y - ADD_NODE_SIZE / 2, ADD_NODE_SIZE, ADD_NODE_SIZE)
    End With
    With oButton
        .name = NAME_BTN_ADD_NODE
        .OnAction = "onClickAddNode"
        .AlternativeText = "create a node"
    End With
    With oButton.TextFrame.Characters
        .Text = "+"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 0, 255)
    End With
End Sub


Sub onclickaddnode()
    Dim oCentre As tCoordinate
    Dim oBtn As Shape
    
    Set oBtn = moMapSheet.Shapes(NAME_BTN_ADD_NODE)
    With oBtn
        oCentre.x = .Left + ADD_NODE_SIZE / 2
        oCentre.y = .Top - ADD_NODE_SIZE / 2
    End With
    
    'remove add node button
    On Error Resume Next
        moMapSheet.Shapes(NAME_BTN_ADD_NODE).Delete
    On Error GoTo 0
    
    'create the actual node
    Call create_node(oCentre)
End Sub



Public Sub on_click_reset()
    Dim sResponse$
    sResponse = MsgBox("Are you sure, this will delete everything!", vbYesNo + vbExclamation, APP_TITLE)
    If sResponse = vbYes Then
        sResponse = MsgBox("Checking again Are you really sure?", vbYesNo + vbCritical, APP_TITLE)
        If sResponse = vbYes Then
            Call init_sheet_vars
            
            'delete nodes, edges, pages
            Call delete_all_rows(CONFIRM_PASS, moNodeSheet, COL_NODE_ID)
            Call delete_all_rows(CONFIRM_PASS, moEdgesSheet, COL_EDGE_ID)
            Call delete_all_rows(CONFIRM_PASS, moPagesSheet, COL_PAGE_ID)
            
            
            'refresh list
            Call create_page("Start")
            
        End If
    End If
End Sub


'####################################################################################################
'# draw mindmap
'####################################################################################################

Sub draw_mindmap(psPage$)
    init_sheet_vars
    
    Dim sCurrent$
    sCurrent = current_page()
    If psPage <> sCurrent Then
        Call save_nodes
        Call save_page_edges
    End If
    
    On Error Resume Next
        Call ActiveWorkbook.save
    On Error GoTo 0
    Dim iPageID&
    iPageID = get_current_page_id
    If iPageID = 0 Then Exit Sub
    
    Call set_current_page(psPage)
    DoEvents
    Call delete_all_shapes
    DoEvents
    Call restore_nodes
    DoEvents
    Call restore_extra_shapes
    DoEvents
    Call restore_edges
    DoEvents
    Call collapse_nodes
End Sub








