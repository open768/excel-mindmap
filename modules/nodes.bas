Attribute VB_Name = "nodes"
Option Explicit

Const SHAPE_WIDTH = 200
Const SHAPE_HEIGHT = 100

Const NODENAME_NAME = "node_name"
Const NODENAME2_NAME = "node_name_2"
Const NODE_TYPE = "node"

Public Const COL_NODE_ID = 1
Const COL_NODE_PAGE = 2
Const COL_NODE_TYPE = 3
Const COL_NODE_TEXT = 4
Const COL_NODE_COLOUR = 5
Const COL_NODE_X = 6
Const COL_NODE_Y = 7
Const COL_NODE_WIDTH = 8
Const COL_NODE_HEIGHT = 9
Const COL_NODE_TXT_COLOUR = 10
Const COL_NODE_border_size = 11
Const COL_NODE_border_colour = 12
Const COL_NODE_SHAPETYPE = 13
Const COL_NODE_ZORDER = 14
Const COL_NODE_TXT_VALIGN = 15
Const COL_NODE_TXT_WRAP = 16
Const COL_NODE_TXT_FONT = 17
Const COL_NODE_TXT_FONTSIZE = 18
Const COL_NODE_ROTATION = 19
Const COL_NODE_COLLAPSED = 20
Const COL_NODE_LINK = 21
Const COL_NODE_TXT_HALIGN = 22

Const NAME_BTN_COLLAPSE = "btn_collapse"
Const NAME_BTN_TMP = "btn_tmp"
Const NAME_BTN_LINK = "btn_link"
Const NAME_BTN_GO = "btn_link_go"
'Const NAME_BTN_CHANGE = "btn_change_shape"

Const BTN_SIZE = 25
Const BTN_MARGIN = 10

Const NODE_PREFIX = "mindmap_"

Public Const ADD_NODE_SIZE = 15

Public Type tNodeInfo
    shp As Shape
    name As String
    id As Long
End Type

Public Type tSelectedNodes
    source As tNodeInfo
    target As tNodeInfo
End Type


Private moCollapsed As Scripting.Dictionary


'####################################################################################################
'# search all nodes
'####################################################################################################
Public Function search_all_nodes(psSearch$) As Collection
    Dim iLastRow&, iRow&, sNodeText$, sSearch$
    
    init_sheet_vars
    
    sSearch = LCase(Trim(psSearch))
    If sSearch = "" Then
        Set search_all_nodes = Nothing
        Exit Function
    End If
        
    Dim oResults As New Collection
    Dim oResult As cSearchResult
    
    iLastRow = get_last_row(moNodeSheet, ColumnLetter(COL_NODE_ID))
    For iRow = 2 To iLastRow Step 1
        sNodeText = LCase(moNodeSheet.Cells(iRow, COL_NODE_TEXT))
        If InStr(sNodeText, sSearch) > 0 Then
            Set oResult = New cSearchResult
            With oResult
                .match = sNodeText
                .page_id = moNodeSheet.Cells(iRow, COL_NODE_PAGE)
                .page_name = get_page_name(.page_id)
            End With
            oResults.Add oResult
        End If
    Next iRow
    
    Set search_all_nodes = oResults
End Function

'####################################################################################################
'# create nodes
'####################################################################################################
Public Function create_node(poCentre As tCoordinate) As Shape
    Dim oShapeLeft As tCoordinate
    oShapeLeft.x = poCentre.x - SHAPE_WIDTH / 2
    oShapeLeft.y = poCentre.y - SHAPE_HEIGHT / 2
    
    Dim oShape As Shape
    Dim iShapetype As MsoAutoShapeType
    iShapetype = get_default_shape
    
    Set oShape = make_a_shape(moMapSheet, iShapetype, oShapeLeft.x, oShapeLeft.y, SHAPE_WIDTH, SHAPE_HEIGHT)
    With oShape
        .Fill.ForeColor.RGB = get_default_colour()
        Dim iWidth%
        
        iWidth = get_default_border_size()
        If iWidth > 0 Then
            .Line.ForeColor.RGB = get_default_border_colour()
            .Line.Weight = iWidth
        End If
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = get_default_font_colour()
        .TextFrame2.TextRange.Font.Size = get_default_font_size()
    End With
    
    Call save_new_shape(NODE_TYPE, oShape)
    oShape.OnAction = "onClickNode"
    
    Set create_node = oShape
End Function


Public Function get_selected_node_details() As tSelectedNodes
    Dim oNodes As tSelectedNodes
    
    With oNodes
        .source.name = selected_node_name()
        .source.id = get_node_id(.source.name)
        .target.name = Range(NODENAME2_NAME).Value
        .target.id = get_node_id(.target.name)
        Set .source.shp = moMapSheet.Shapes(.source.name)
        Set .target.shp = moMapSheet.Shapes(.target.name)
    End With
    
    get_selected_node_details = oNodes
End Function

Public Function get_node_info(piNodeID&) As tNodeInfo
    Dim oInfo As tNodeInfo
    
    oInfo.id = piNodeID
    oInfo.name = NODE_PREFIX & piNodeID
    On Error Resume Next
        Set oInfo.shp = moMapSheet.Shapes(oInfo.name)
    On Error GoTo 0
    
    get_node_info = oInfo
End Function

'####################################################################################################
'# node interactions
'####################################################################################################
Sub onClickNode()
    Dim sName$, oShape As Shape
   
    Call init_sheet_vars
    Call remove_all_buttons
    
    sName$ = Application.Caller
    Set oShape = moMapSheet.Shapes(sName$)
    
    If GetKeyState(VK_SHIFT) < 0 Then
        'shift key is pressed
        Call shift_select_node(oShape)
    Else
        Call store_selected_node(sName$)
        Range(NODENAME2_NAME).Value = ""
        Call oShape.Select
        Call add_collapse_button
        Call add_link_button
        'Call add_change_shape_button
    End If
End Sub


Sub shift_select_node(poShape As Shape)

    'dont select more than 2 shapes
    Dim iSelected&
    iSelected = CountSelectedShapes()
    If Not iSelected = 1 Then Exit Sub

    'if the shape is the current one selected, deselect it and exit
    Dim sCurrent$
    sCurrent$ = selected_node_name()
    If sCurrent = poShape.name Then
        Call store_selected_node("")
        Call deselect_shapes
        Exit Sub
    End If
    
    'select the shape
    poShape.Select (False)
    
    ' find out the 2nd shape name
    Dim oRange As ShapeRange, sShape1$, sShape2$
    sShape1$ = selected_node_name()
    
    Set oRange = Selection.ShapeRange

    If oRange(1).name = sShape1 Then
        sShape2 = oRange(2).name
    Else
        sShape2 = oRange(1).name
    End If
    Range(NODENAME2_NAME).Value = sShape2
    
    'check selected shape edge connections
    Call check_selected_node_edges
    
End Sub

Sub on_click_change_shape()
  'not sure this is even possible
End Sub

'################################################################################
'# Node IDs
'################################################################################
Public Function get_node_row(psName$) As Long
    Dim iID&
    
    'check if the name starts with a prefix
    If InStr(1, psName, NODE_PREFIX) = 0 Then
        get_node_row = 0
        Exit Function
    End If
    
    iID = CLng(Right(psName, Len(psName) - Len(NODE_PREFIX)))
    get_node_row = CLng(Application.match(iID, moNodeSheet.Columns("A"), EXACT_MATCH))

End Function


Public Function get_next_node_name(piPage&) As String
    
    Dim iRow&, iNum&
    iRow = get_last_row(moNodeSheet, ColumnLetter(COL_PAGE_ID))
    If iRow < 2 Then
        iNum = 1
    Else
        iNum = Int(moNodeSheet.Cells(iRow, COL_NODE_ID).Value) + 1
    End If
    
    Dim sName$
    get_next_node_name = NODE_PREFIX & iNum
End Function

Function get_node_id(psName$) As Long
    Dim iPos&, sRight$
    iPos = InStr(psName, NODE_PREFIX)
    If iPos = 0 Then Err.Raise vbObjectError, "shape name not valid"
    
    sRight$ = Right(psName, Len(psName) - Len(NODE_PREFIX))
    get_node_id = CLng(sRight)
End Function

Function get_current_node() As Shape

End Function

'################################################################################
'# SAVE and restore
'################################################################################
Public Sub save_nodes()
    Call init_sheet_vars
    
    ' iterate through and update shape data
    Dim oShape As Shape
    Dim iZOrder&
    
    iZOrder = 1
    For Each oShape In moMapSheet.Shapes
        If oShape.Type = msoAutoShape Then
            Call update_node(oShape, iZOrder)
            iZOrder = iZOrder + 1
        End If
    Next oShape
    
    Call remove_deleted_nodes
End Sub

Public Sub restore_nodes()
    Dim iPageID&
    iPageID = get_current_page_id()
    
    Dim iLastRow&
    iLastRow = get_last_row(moNodeSheet, ColumnLetter(COL_NODE_ID))
    
    Dim iRow&, iRowPageID&
    For iRow = 2 To iLastRow
        iRowPageID = moNodeSheet.Cells(iRow, COL_NODE_PAGE).Value
        If iRowPageID = iPageID Then Call restore_node_at_row(iRow)
    Next iRow
    
    DoEvents
    Call correct_node_zorder
End Sub

Public Function save_new_shape(psType$, poShape As Shape)
    Dim iPage&, sName$
    
    iPage = get_current_page_id()
    sName = get_next_node_name(iPage)
    poShape.name = sName
    
    Dim iRow&
    iRow = get_last_row(moNodeSheet, ColumnLetter(COL_NODE_ID)) + 1
    
    With moNodeSheet
        .Cells(iRow, COL_NODE_ID).Value = .Cells(iRow - 1, COL_NODE_ID).Value + 1
        .Cells(iRow, COL_NODE_PAGE).Value = iPage
        .Cells(iRow, COL_NODE_TYPE).Value = psType
    End With
    Call update_node_row(poShape, iRow)

End Function

Public Sub update_node(poShape As Shape, piZ&)
    'get the node row
    Dim iRow&
    iRow = get_node_row(poShape.name)
    If iRow = 0 Then Exit Sub
    
     moNodeSheet.Cells(iRow, COL_NODE_ZORDER).Value = piZ
     Call update_node_row(poShape, iRow)
     
End Sub

Private Sub update_node_row(poShape As Shape, piRow&)
    If piRow = 0 Then Exit Sub
    
     With moNodeSheet
        .Cells(piRow, COL_NODE_TEXT).Value = poShape.TextFrame2.TextRange.Text
        .Cells(piRow, COL_NODE_X).Value = poShape.Left
        .Cells(piRow, COL_NODE_Y).Value = poShape.Top
        .Cells(piRow, COL_NODE_WIDTH).Value = poShape.Width
        .Cells(piRow, COL_NODE_HEIGHT).Value = poShape.Height
        .Cells(piRow, COL_NODE_COLOUR).Value = poShape.Fill.ForeColor.RGB
        .Cells(piRow, COL_NODE_border_size).Value = poShape.Line.Weight
        .Cells(piRow, COL_NODE_border_colour).Value = poShape.Line.ForeColor.RGB
        .Cells(piRow, COL_NODE_SHAPETYPE).Value = poShape.AutoShapeType
        .Cells(piRow, COL_NODE_TXT_WRAP).Value = poShape.TextFrame2.WordWrap
        .Cells(piRow, COL_NODE_TXT_COLOUR).Value = poShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        .Cells(piRow, COL_NODE_TXT_VALIGN).Value = poShape.TextFrame2.VerticalAnchor
        .Cells(piRow, COL_NODE_TXT_HALIGN).Value = poShape.TextFrame2.TextRange.ParagraphFormat.Alignment
        .Cells(piRow, COL_NODE_TXT_FONT).Value = poShape.TextFrame2.TextRange.Font.name
        .Cells(piRow, COL_NODE_ROTATION).Value = poShape.Rotation
        .Cells(piRow, COL_NODE_TXT_FONTSIZE).Value = poShape.TextFrame2.TextRange.Font.Size
    End With
    
End Sub

Public Sub restore_node_at_row(piRow&)
    Dim oShape As Shape, sFont$, iFontSize&
    
    ' draw a shape
    With moNodeSheet
        Set oShape = make_a_shape( _
            moMapSheet, _
            msoShapeOval, _
            .Cells(piRow, COL_NODE_X).Value, _
            .Cells(piRow, COL_NODE_Y).Value, _
            .Cells(piRow, COL_NODE_WIDTH).Value, _
            .Cells(piRow, COL_NODE_HEIGHT).Value _
        )
        
        oShape.name = NODE_PREFIX & .Cells(piRow, COL_NODE_ID).Text
        oShape.AutoShapeType = .Cells(piRow, COL_NODE_SHAPETYPE).Value
    End With
        
    On Error Resume Next
        With oShape.TextFrame2
            .TextRange.Text = moNodeSheet.Cells(piRow, COL_NODE_TEXT).Text
            .HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Font.Fill.ForeColor.RGB = moNodeSheet.Cells(piRow, COL_NODE_TXT_COLOUR).Value
            .WordWrap = moNodeSheet.Cells(piRow, COL_NODE_TXT_VALIGN).Value
            .VerticalAnchor = moNodeSheet.Cells(piRow, COL_NODE_TXT_VALIGN).Value
            .TextRange.ParagraphFormat.Alignment = moNodeSheet.Cells(piRow, COL_NODE_TXT_HALIGN).Value
            sFont = moNodeSheet.Cells(piRow, COL_NODE_TXT_FONT).Value
            If Not sFont = "" Then .TextRange.Font.name = sFont
            
            iFontSize = moNodeSheet.Cells(piRow, COL_NODE_TXT_FONTSIZE).Value
            If Not iFontSize = 0 Then .TextRange.Font.Size = iFontSize
        End With
    
        'set the colour
        Dim iColour&
        iColour = moNodeSheet.Cells(piRow, COL_NODE_COLOUR).Value
        oShape.Fill.ForeColor.RGB = iColour
        
        
        oShape.Rotation = moNodeSheet.Cells(piRow, COL_NODE_ROTATION).Value
    
        'set the border
        With oShape.Line
            .Weight = moNodeSheet.Cells(piRow, COL_NODE_border_size).Value
            .ForeColor.RGB = moNodeSheet.Cells(piRow, COL_NODE_border_colour).Value
        End With
    On Error GoTo 0
        
    'click handler
    oShape.OnAction = "onClickNode"
End Sub
Public Sub remove_deleted_nodes()
    Dim iPage&
    iPage = get_current_page_id()
    
    ' iterate through the data and remove shapes that have been deleted
    Dim iLastRow&, iRow&, iRowPage, iRowID&, sRowName$, bExists As Boolean
    iLastRow = get_last_row(moNodeSheet, "A")
    For iRow = iLastRow To 2 Step -1
        iRowPage = moNodeSheet.Cells(iRow, COL_NODE_PAGE)
        If iRowPage = iPage Then
            sRowName = NODE_PREFIX & moNodeSheet.Cells(iRow, COL_NODE_ID).Value
            If Not shape_exists(sRowName) Then moNodeSheet.Rows(iRow).Delete
        End If
    Next iRow
End Sub

Public Sub correct_node_zorder()

    Dim iLastRow&, iPage&
    Dim sColPage$, sColZ$
    Dim sDataRange$, sPageRange$
    
    iLastRow = get_last_row(moNodeSheet, "A")
    iPage = get_current_page_id()
    sColPage = ColumnLetter(COL_NODE_PAGE)
    sColZ = ColumnLetter(COL_NODE_ZORDER)
    sDataRange = "A2:" & sColZ & iLastRow
    sPageRange = sColPage & "2:" & sColPage & iLastRow
        
    'Expressions
    Dim sFilterFn$, sSortFn$, vResult As Variant
    
    sFilterFn = "FILTER(" & sDataRange & "," & sPageRange & "=" & iPage & ")"
    sSortFn = "SORT(" & sFilterFn & "," & COL_NODE_ZORDER & ",1)"
    vResult = moNodeSheet.Evaluate(sSortFn)
    If IsError(vResult) Then Exit Sub
    
    'if there is only one shape the vresult is different
    Dim itest
    On Error Resume Next
        itest = UBound(vResult, 2)
        If Not Err.Number = 0 Then Exit Sub
    On Error GoTo 0
    
    'bring shapes to front
    Dim i&
    For i = LBound(vResult) To UBound(vResult)
        Dim sName$, oShape As Shape
        
        sName = NODE_PREFIX & vResult(i, COL_NODE_ID)
        Set oShape = moMapSheet.Shapes(sName)
        oShape.ZOrder msoBringToFront
    Next i
End Sub

'################################################################################
Function selected_node_name() As String
    selected_node_name = Range(NODENAME_NAME).Value
End Function

Function selected_node() As Shape
    Dim sName$
    sName = selected_node_name()
    Set selected_node = moMapSheet.Shapes(sName)
End Function

Sub store_selected_node(psName$)
    Range(NODENAME_NAME).Value = psName
End Sub

'################################################################################
'# change shape
'
'Sub remove_change_shape_button()
'    If shape_exists(NAME_BTN_CHANGE) Then moMapSheet.Shapes(NAME_BTN_CHANGE).Delete
'End Sub
'Sub add_change_shape_button()
'    Call remove_change_shape_button
'
'    Dim oNode As Shape, oBtnCoord As tCoordinate, oBtn As Shape
'    Set oNode = selected_node()
'    oBtnCoord = getShapeLeft(oNode)
'    With oBtnCoord
'        .x = .x - BTN_MARGIN - 2 * BTN_SIZE
'        .y = .y + BTN_MARGIN
'        Set oBtn = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x + BTN_SIZE / 2, .y - BTN_SIZE, BTN_SIZE, BTN_SIZE)
'        oBtn.name = NAME_BTN_CHANGE
'    End With
'
'    Dim SHUFFLE_EMOJI$
'    SHUFFLE_EMOJI = ChrW(&HD83D) & ChrW(&HDD00)
'    With oBtn
'        .TextFrame.Characters.Text = SHUFFLE_EMOJI
'        .AlternativeText = "change shape"
'        .OnAction = "on_click_change_shape"
'    End With
'
'End Sub

'################################################################################
'# hyperlink
Sub remove_link_button()
    If shape_exists(NAME_BTN_LINK) Then moMapSheet.Shapes(NAME_BTN_LINK).Delete
    If shape_exists(NAME_BTN_GO) Then moMapSheet.Shapes(NAME_BTN_GO).Delete
End Sub

Sub add_link_button()
    Call remove_link_button
    
    Dim oNode As Shape, iID&, iRow&, sLink$
    
    Set oNode = selected_node()
    iID = get_node_id(oNode.name)
    iRow = get_node_row(oNode.name)
    
    'create a button
    Dim oBtnCoord As tCoordinate, oBtn As Shape
    oBtnCoord = getShapeBottomRight(oNode)
    With oBtnCoord
        .x = .x + BTN_MARGIN
        .y = .y + BTN_MARGIN
        Set oBtn = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x + BTN_SIZE / 2, .y - BTN_SIZE, BTN_SIZE, BTN_SIZE)
        oBtn.name = NAME_BTN_LINK
    End With
    
    'change the button behaviour depending on whether there is a link or not
    Dim CREATE_LINK_EMOJI$, BREAK_LINK_EMOJI$, GO_LINK_EMOJI$, GO_EMOJI$
    
    GO_EMOJI = ChrW(&HD83D) & ChrW(&HDC46)        'finger
    CREATE_LINK_EMOJI = ChrW(&HD83C) & ChrW(&HDF10) 'globe
    BREAK_LINK_EMOJI = ChrW(&H26D3) & ChrW(&HFE0F) & ChrW(&H200D) & ChrW(&HD83D) & ChrW(&HDCA5) ' break link
    
    With oBtn.TextFrame.Characters
        Dim bIsCreateBtn As Boolean
        sLink = moNodeSheet.Cells(iRow, COL_NODE_LINK).Value
        If sLink = "" Then
            bIsCreateBtn = True
            .Text = CREATE_LINK_EMOJI
            oBtn.AlternativeText = "create link"
            oBtn.OnAction = "on_click_make_link"
        Else
            bIsCreateBtn = False
            .Text = BREAK_LINK_EMOJI
            oBtn.AlternativeText = "break link"
            oBtn.OnAction = "on_click_unlink"
        End If
    End With
    
    If Not bIsCreateBtn Then
        Dim oBtnGo As Shape
        oBtnCoord = getShapeRight(oNode)
        With oBtnCoord
            .x = .x + BTN_MARGIN
            .y = .y + BTN_MARGIN
            Set oBtn = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x + BTN_SIZE / 2, .y - BTN_SIZE, BTN_SIZE, BTN_SIZE)
        End With
        
        With oBtn
            .name = NAME_BTN_GO
            .OnAction = "on_click_go"
            .TextFrame.Characters.Text = GO_EMOJI
        End With
    

    End If
       
End Sub

Sub on_click_go()
    Call init_sheet_vars
    Call remove_all_buttons
    
    Dim oNode As Shape, iID&, iRow&, sLink$
    
    Set oNode = selected_node()
    iID = get_node_id(oNode.name)
    iRow = get_node_row(oNode.name)
    Range("a1").Select
    
    sLink = moNodeSheet.Cells(iRow, COL_NODE_LINK).Value
    If Left(sLink, 1) = "#" Then
        Dim sPage$, iPageRow&
        sPage = Right(sLink, Len(sLink) - 1)
        Call select_page_in_list(sPage)
    Else
        ThisWorkbook.FollowHyperlink sLink
    End If
    
End Sub

Sub on_click_make_link()
    Dim sLink$
    Call init_sheet_vars
    
    sLink = InputBox("please enter link (use # for page name)", APP_TITLE)
    
    If sLink <> "" Then
        '----------------- local link
        If Left(sLink, 1) = "#" Then
            Dim sPage$, iPageRow&
            sPage = Right(sLink, Len(sLink) - 1)
            On Error Resume Next
                iPageRow = get_page_row(sPage)
                If Err.Number <> 0 Then
                    MsgBox "unknown page: " & sPage, vbOKOnly + vbExclamation, APP_TITLE
                    Exit Sub
                End If
            On Error GoTo 0
        End If
        
        '-----------
        Dim iRow&, oNode As Shape
        Set oNode = selected_node
        iRow = get_node_row(oNode.name)
        moNodeSheet.Cells(iRow, COL_NODE_LINK).Value = sLink
    End If
    
    Call remove_all_buttons
    Range("a1").Select

End Sub

Sub on_click_unlink()
    Call init_sheet_vars
    
    Dim iRow&, oNode As Shape
    Set oNode = selected_node
    iRow = get_node_row(oNode.name)
    moNodeSheet.Cells(iRow, COL_NODE_LINK).Value = ""

    Call remove_link_button
    Range("a1").Select
End Sub

'################################################################################
'# collapse nodes
Sub remove_collapse_button()
    If shape_exists(NAME_BTN_COLLAPSE) Then moMapSheet.Shapes(NAME_BTN_COLLAPSE).Delete
End Sub

Sub add_collapse_button()
    Dim iPage&, iID&, oBtn As Shape, oNode As Shape
    Dim oBtnCoord As tCoordinate
    
    Call remove_all_buttons
    
    'get the selected shape
    Set oNode = selected_node()
        
    'if the node doesnt have outward connections ignore
    Dim bCollapsed As Boolean, iRow&, iCount&, sCol$
    iID = get_node_id(oNode.name)
    iCount = Application.WorksheetFunction.CountIf(moEdgesSheet.Columns(COL_EDGE_SOURCE), iID)
    If iCount = 0 Then Exit Sub
    
    'create the collapse button
    oBtnCoord = getShapeTopRight(oNode)
    With oBtnCoord
        .x = .x + BTN_MARGIN
        .y = .y - BTN_MARGIN
        Set oBtn = moMapSheet.Shapes.AddFormControl(xlButtonControl, .x - BTN_SIZE / 2, .y - BTN_SIZE / 2, BTN_SIZE, BTN_SIZE)
        oBtn.name = NAME_BTN_COLLAPSE
    End With
    
    
    'check if node is collapsed
    Dim OPEN_UMBRELLA$, CLOSED_UMBRELLA$
    OPEN_UMBRELLA$ = ChrW(&H2602)
    CLOSED_UMBRELLA$ = ChrW(&HD83C) & ChrW(&HDF02)
    
    iRow = get_node_row(oNode.name)
    bCollapsed = moNodeSheet.Cells(iRow, COL_NODE_COLLAPSED).Value = 1
    With oBtn.TextFrame.Characters
        If bCollapsed Then
            .Text = OPEN_UMBRELLA
            oBtn.AlternativeText = "expand"
            oBtn.OnAction = "onClickExpand"
        Else
            .Text = CLOSED_UMBRELLA
            oBtn.AlternativeText = "collapse"
            oBtn.OnAction = "onClickCollapse"
        End If
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 0, 255)
    End With
    
    oBtn.ZOrder msoBringToFront
    'form controls cant be rotated
End Sub

Sub onClickCollapse()
    Call prCollapse(True)
End Sub

Sub onClickExpand()
    Call prCollapse(False)
End Sub

Private Sub prCollapse(pbCollapsing As Boolean)
   Call init_sheet_vars
    Set moCollapsed = New Scripting.Dictionary
    
    Dim iID&, oNode As Shape
    Set oNode = selected_node()
    iID = get_node_id(oNode.name)
    
    'hide the button
    remove_all_buttons
    
    'mark node as collapsed
    Dim iRow&
    moCollapsed(iID) = 1
    iRow = get_node_row(oNode.name)
    
    If pbCollapsing Then
        moNodeSheet.Cells(iRow, COL_NODE_COLLAPSED).Value = 1
    Else
        moNodeSheet.Cells(iRow, COL_NODE_COLLAPSED).Value = 0
    End If
    Call collapse_outgoing_edges(iID, pbCollapsing)
End Sub

Public Sub collapse_node(piID&, pbCollapsing As Boolean)
    If moCollapsed.Exists(piID) Then Exit Sub
    moCollapsed(piID) = 1
    
    ' hide the shape
    Dim oInfo As tNodeInfo
    oInfo = get_node_info(piID)
    If oInfo.shp Is Nothing Then Exit Sub
    
    If pbCollapsing Then
        oInfo.shp.Visible = msoFalse
    Else
        oInfo.shp.Visible = msoTrue
    End If
    Call collapse_outgoing_edges(piID, pbCollapsing)
End Sub

Sub collapse_nodes()
    'now hide nodes that are collapsed
    Dim iLastRow&, iRow&, iRowPageID&, iNodeID&, iPageID&
    iLastRow = get_last_row(moNodeSheet, ColumnLetter(COL_NODE_ID))
    iPageID = get_current_page_id
    
    Dim bCollapsed As Boolean
    Set moCollapsed = New Scripting.Dictionary

    For iRow = 2 To iLastRow
        iRowPageID = moNodeSheet.Cells(iRow, COL_NODE_PAGE).Value
        iNodeID = moNodeSheet.Cells(iRow, COL_NODE_ID).Value
        bCollapsed = (moNodeSheet.Cells(iRow, COL_NODE_COLLAPSED).Value = 1)
        If iRowPageID = iPageID And bCollapsed Then Call collapse_outgoing_edges(iNodeID, True)
    Next iRow
End Sub

