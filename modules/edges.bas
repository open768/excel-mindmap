Attribute VB_Name = "edges"
Option Explicit

Public Const COL_EDGE_ID = 1
Public Const COL_EDGE_PAGE = 2
Public Const COL_EDGE_SOURCE = 3
Public Const COL_EDGE_TARGET = 4
Public Const COL_EDGE_COLOUR = 5
Public Const COL_EDGE_THICKNESS = 6

Public Const EDGE_BUTTON_NAME = "btn_create_node"
Public Const EDGE_BUTTON_size = 20
Const EDGE_PREFIX = "mindmapedge-"

'##########################################################################################################
'# ACTION BUTTON
'##########################################################################################################
Public Sub remove_edge_button()
    On Error Resume Next
    moMapSheet.Shapes(EDGE_BUTTON_NAME).Delete
    On Error GoTo 0
End Sub

Public Sub check_selected_node_edges()
    Dim iPage&
    Dim oSrcCol As Range, oTgtCol As Range, oPageCol
    Dim oNodes As tSelectedNodes
    
    oNodes = get_selected_node_details()
    iPage = get_current_page_id()
    Set oSrcCol = moEdgesSheet.Columns(COL_EDGE_SOURCE)
    Set oTgtCol = moEdgesSheet.Columns(COL_EDGE_TARGET)
    Set oPageCol = moEdgesSheet.Columns(COL_EDGE_PAGE)
    
    'lookup edge ID in worksheet "pages" column "B" and get the value from column "A"
    Dim iCount
    iCount = Application.WorksheetFunction.CountIfs(oPageCol, iPage, oSrcCol, oNodes.source.id, oTgtCol, oNodes.target.id)
    If iCount = 0 Then iCount = Application.WorksheetFunction.CountIfs(oPageCol, iPage, oTgtCol, oNodes.source.id, oSrcCol, oNodes.target.id)
    If iCount > 0 Then
        Call deselect_shapes
        Exit Sub
    End If
    
    'there were no edges,
    Call create_edge_button
End Sub


Public Sub create_edge_button()
    ' remove button
    Call remove_all_buttons
    
    'get the selected shapes
    Dim oNodes As tSelectedNodes
    oNodes = get_selected_node_details()
    
    'figure out the middle of the two shapes
    Dim oCentre As tCoordinate, oSrcMiddle As tCoordinate, oTgtMiddle As tCoordinate
    With oNodes
        oSrcMiddle = shape_middle(.source.shp)
        oTgtMiddle = shape_middle(.target.shp)
    End With
    
    With oCentre
        .x = (oSrcMiddle.x + oTgtMiddle.x) / 2
        .y = (oSrcMiddle.y + oTgtMiddle.y) / 2
    
        'create the button
        Dim oButton As Shape
        Set oButton = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x - EDGE_BUTTON_size / 2, .y - EDGE_BUTTON_size / 2, EDGE_BUTTON_size, EDGE_BUTTON_size)
        With oButton
            .name = EDGE_BUTTON_NAME
            .OnAction = "onClickCreateEdge"
        End With
        With oButton.TextFrame.Characters
            .Text = "*"
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
            .Font.Size = 20
        End With
        
    End With
End Sub

'##########################################################################################################
'# CREATE EDGE
'##########################################################################################################
Public Sub onClickCreateEdge()
    init_sheet_vars
    
    Dim oNodes As tSelectedNodes, iID&, iPage&
    oNodes = get_selected_node_details()
    iPage = get_current_page_id()
    iID = get_next_edge_id()
        
    'create the connector
    Dim oEdge As Shape
    Set oEdge = make_edge(iID, oNodes)
    
    'store the edge details
    Dim iRow&
    iRow = get_last_row(moEdgesSheet, "A") + 1
    With moEdgesSheet
        .Cells(iRow, COL_EDGE_ID).Value = iID
        .Cells(iRow, COL_EDGE_PAGE).Value = iPage
        .Cells(iRow, COL_EDGE_SOURCE).Value = oNodes.source.id
        .Cells(iRow, COL_EDGE_TARGET).Value = oNodes.target.id
        .Cells(iRow, COL_EDGE_COLOUR).Value = oEdge.Line.ForeColor.RGB
        .Cells(iRow, COL_EDGE_THICKNESS).Value = oEdge.Line.Weight
    End With
    
    remove_all_buttons
    Call deselect_shapes

End Sub

Public Function make_edge(piID&, poNodes As tSelectedNodes) As Shape
    Dim oEdge As Shape
    
    Set oEdge = moMapSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
    With oEdge
        .ConnectorFormat.BeginConnect poNodes.source.shp, 1
        .ConnectorFormat.EndConnect poNodes.target.shp, 1
        .Line.EndArrowheadStyle = msoArrowheadTriangle
        .name = EDGE_PREFIX & piID
        Call .RerouteConnections
    End With
    
    Set make_edge = oEdge

End Function

Function get_next_edge_id() As Long
    Dim iRow&, iID&
    iRow = get_last_row(moEdgesSheet, "A")
    If iRow < 2 Then
        iID = 1
    Else
        iID = moEdgesSheet.Cells(iRow, COL_EDGE_ID).Value
        iID = iID + 1
    End If
    get_next_edge_id = iID
End Function


Public Function get_edge_row(psName$) As Long
    Dim iID&
    
    'check if the name starts with a prefix
    If InStr(1, psName, EDGE_PREFIX) = 0 Then
        get_edge_row = 0
        Exit Function
    End If
    
    iID = CLng(Right(psName, Len(psName) - Len(EDGE_PREFIX)))
    get_edge_row = CLng(Application.match(iID, moEdgesSheet.Columns(COL_EDGE_ID), EXACT_MATCH))
End Function

'#############################################################################################
'# save and restore
'#############################################################################################
Public Sub restore_edges()
    Dim iPage&
    iPage = get_current_page_id
    
    Dim iLastRow&
    iLastRow = get_last_row(moEdgesSheet, "A")
    
    Dim iRow&, iRowPageID&
    For iRow = 2 To iLastRow
        iRowPageID = moEdgesSheet.Cells(iRow, COL_EDGE_PAGE).Value
        If iRowPageID = iPage Then Call restore_edge_at_row(iRow)
    Next iRow
End Sub

Private Sub restore_edge_at_row(piRow&)
    Dim oNodes As tSelectedNodes, iID&
    Dim oEdge As Shape
    
    
    iID = moEdgesSheet.Cells(piRow, COL_EDGE_ID).Value
    oNodes.source = get_node_info(moEdgesSheet.Cells(piRow, COL_EDGE_SOURCE).Value)
    oNodes.target = get_node_info(moEdgesSheet.Cells(piRow, COL_EDGE_TARGET).Value)
    
    If oNodes.source.shp Is Nothing Then
        Debug.Print "cant find edge restore source"
        Exit Sub
    End If
    If oNodes.target.shp Is Nothing Then
        Debug.Print "cant find edge restore target"
        Exit Sub
    End If
    
    Set oEdge = make_edge(iID, oNodes)
    With oEdge.Line
        .ForeColor.RGB = moEdgesSheet.Cells(piRow, COL_EDGE_COLOUR).Value
        .Weight = moEdgesSheet.Cells(piRow, COL_EDGE_THICKNESS).Value
    End With


End Sub

Public Sub save_page_edges()
    Dim oShape As Shape
    
    Call init_sheet_vars
    
    For Each oShape In moMapSheet.Shapes
        If oShape.Connector Then Call update_edge(oShape)
    Next oShape
    Call remove_deleted_edges
End Sub


Public Sub update_edge(poShape As Shape)
    Dim iPage&
    iPage = get_current_page_id
    
    Dim iRow&
    iRow = get_edge_row(poShape.name)
    If iRow = 0 Then Exit Sub
    
    With moEdgesSheet
        .Cells(iRow, COL_EDGE_COLOUR).Value = poShape.Line.ForeColor.RGB
        .Cells(iRow, COL_EDGE_THICKNESS).Value = poShape.Line.Weight
    End With
    
End Sub

Public Sub remove_deleted_edges()

    Dim iLastRow&, iRow&, iID&, iPage&, sName$, bExists As Boolean
    
    iPage = get_current_page_id
    
    iLastRow = get_last_row(moEdgesSheet, "A")
    With moEdgesSheet
        For iRow = iLastRow To 2 Step -1
            If .Cells(iRow, COL_EDGE_PAGE) = iPage Then
                iID = .Cells(iRow, COL_EDGE_ID).Value
                sName = EDGE_PREFIX & iID
                bExists = False
                On Error Resume Next
                    bExists = Not moMapSheet.Shapes(sName) Is Nothing
                On Error GoTo 0
                If Not bExists Then .Rows(iRow).Delete
            End If
        Next iRow
    End With
End Sub

'#############################################################################################
'collapse nodes - following edges
Public Sub collapse_outgoing_edges(piNodeID&, pbCollapsing As Boolean)
    Dim iPage&, iLastRow&, iRow&, iEdgeID&, sEdgeName$, iTargetNode&
    
    iPage = get_current_page_id()
    iLastRow = get_last_row(moEdgesSheet, "A")
    
    With moEdgesSheet
        For iRow = 2 To iLastRow
            If .Cells(iRow, COL_EDGE_PAGE).Value = iPage And .Cells(iRow, COL_EDGE_SOURCE).Value = piNodeID Then
                iEdgeID = .Cells(iRow, COL_EDGE_ID).Value
                sEdgeName = EDGE_PREFIX & iEdgeID
                On Error Resume Next
                    moMapSheet.Shapes(sEdgeName).Visible = Not pbCollapsing
                On Error GoTo 0
                
                iTargetNode = .Cells(iRow, COL_EDGE_TARGET).Value
                Call collapse_node(iTargetNode, pbCollapsing)
            End If
        Next iRow
    End With
    DoEvents
End Sub
