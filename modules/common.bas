Attribute VB_Name = "common"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey&) As Integer
#Else
    Public Declare Function GetKeyState Lib "user32" (ByVal vKey&) As Integer
#End If
Public Const VK_SHIFT& = &H10
Public Const EXACT_MATCH = 0
Public Const APP_TITLE = "Excel Mindmap experiment - "

Public Type tCoordinate
    x As Long
    y As Long
End Type


Public Function get_list_selection(poSheet As Worksheet, poList As Shape) As String
    If poList Is Nothing Then
        MsgBox "unknown list", vbOKOnly + vbExclamation, APP_TITLE
        Exit Function
    End If
    
    Dim iIndex
    iIndex = poList.ControlFormat.Value
    If iIndex < 0 Then Exit Function
    
    On Error Resume Next
    get_list_selection = poList.ControlFormat.List(iIndex)
    On Error GoTo 0
End Function


'############################################################################################
'# Points
'############################################################################################
Function rotate_point(poCentre As tCoordinate, poPoint As tCoordinate, piDegrees As Integer) As tCoordinate
    Dim iAng As Double
    Dim oNew As tCoordinate
    
    iAng = piDegrees * WorksheetFunction.Pi / 180 'convert degrees to radians
    
    With oNew
        .x = ((poPoint.x - poCentre.x) * Math.Cos(iAng)) - ((poPoint.y - poCentre.y) * Math.Sin(iAng))
        .y = ((poPoint.x - poCentre.x) * Math.Sin(iAng)) + ((poPoint.y - poCentre.y) * Math.Cos(iAng))
        
        .x = .x + poCentre.x
        .y = .y + poCentre.y
    End With
    
    rotate_point = oNew
End Function



'############################################################################################
'# soudex
'############################################################################################
Function Soundex(ByVal txt As String) As String
    Dim i As Integer
    Dim result As String
    Dim prevCode As String
    Dim currCode As String
    Dim firstLetter As String
    
    txt = UCase(txt)
    If Len(txt) = 0 Then Exit Function
    
    firstLetter = Left(txt, 1)
    result = firstLetter
    prevCode = soundex_mapcode(firstLetter)
    
    For i = 2 To Len(txt)
        currCode = soundex_mapcode(Mid(txt, i, 1))
        
        If currCode <> prevCode And currCode <> "" Then
            result = result & currCode
        End If
        
        If currCode <> "" Then prevCode = currCode
    Next i
    
    ' Pad or trim to 4 characters
    result = result & "0000"
    Soundex = Left(result, 4)
End Function

Private Function soundex_mapcode(ByVal ch As String) As String
    Select Case ch
        Case "B", "F", "P", "V": soundex_mapcode = "1"
        Case "C", "G", "J", "K", "Q", "S", "X", "Z": soundex_mapcode = "2"
        Case "D", "T": soundex_mapcode = "3"
        Case "L": soundex_mapcode = "4"
        Case "M", "N": soundex_mapcode = "5"
        Case "R": soundex_mapcode = "6"
        Case Else: soundex_mapcode = ""
    End Select
End Function


'############################################################################################
'# Sheets
'############################################################################################
Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Function sheet_must_exist(psSheet$) As Boolean
    Dim oSheet As Worksheet
    sheet_must_exist = False
    On Error Resume Next
        Set oSheet = Worksheets(psSheet)
        If oSheet Is Nothing Then
            MsgBox "sheet is missing " & psSheet, vbOKOnly + vbCritical, APP_TITLE
            Exit Function
        End If
    On Error GoTo 0
    sheet_must_exist = True
End Function

Sub delete_all_rows(psPass$, poSheet As Worksheet, piCol&)
    If Not psPass = CONFIRM_PASS Then Exit Sub
    
    Dim iLastRow&
    iLastRow = get_last_row(poSheet, ColumnLetter(piCol))
    If iLastRow < 2 Then Exit Sub
    poSheet.Rows("2:" & iLastRow).Delete
End Sub

Public Function get_last_row(poSheet As Worksheet, psColumn$) As Integer
    Dim iRows&
    Dim oLast As Range
    
    iRows = poSheet.Rows.Count
    Set oLast = poSheet.Cells(iRows, psColumn)
    
    
    get_last_row = oLast.End(xlUp).Row
End Function

'############################################################################################
'# Shapes
'############################################################################################
Public Sub delete_all_shapes()
    Dim oShp As Shape
    
    Call init_sheet_vars

    For Each oShp In moMapSheet.Shapes
        If (Not moReservedControls.Exists(oShp.name)) Then oShp.Delete
    Next oShp
End Sub

Public Function make_a_shape(poSheet As Worksheet, piType As MsoAutoShapeType, piX&, piY&, piwidth&, piheight&) As Shape
    Dim oShape As Shape
    Set oShape = poSheet.Shapes.AddShape(piType, piX, piY, piwidth, piheight)

    With oShape.TextFrame2
        .TextRange.Text = "idea"
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    
    Set make_a_shape = oShape
End Function

Function CountSelectedShapes() As Long
    Dim iCount&

    iCount = 0
    Dim sType$
    If TypeName(Selection) = "DrawingObjects" Then
        iCount = Selection.ShapeRange.Count
    Else
        iCount = 1
    End If
    
    CountSelectedShapes = iCount
        
End Function

Public Function shape_middle(poShape As Shape) As tCoordinate
    Dim oCentre As tCoordinate
    
    With poShape
        oCentre.x = .Left + .Width / 2
        oCentre.y = .Top + .Height / 2
    End With
    
    shape_middle = oCentre
End Function

Function shape_exists(psName$, Optional poSheet As Worksheet) As Boolean
    shape_exists = False
    Dim oSheet As Worksheet
    If Not poSheet Is Nothing Then
        Set oSheet = poSheet
    Else
        Set oSheet = moMapSheet
    End If
    
    On Error Resume Next
        shape_exists = Not oSheet.Shapes(psName$) Is Nothing
    On Error GoTo 0
End Function

Function getShapeTopRight(poShape As Shape) As tCoordinate
    Dim oPoint As tCoordinate, oCentre As tCoordinate, oRotated As tCoordinate
    With poShape
        oCentre.x = .Left + (.Width / 2)
        oCentre.y = .Top + (.Height / 2)
        
        oPoint.x = .Left + .Width
        oPoint.y = .Top
        
        oRotated = rotate_point(oCentre, oPoint, .Rotation)
    End With
    
    
    getShapeTopRight = oRotated
End Function

Function getShapeBottomRight(poShape As Shape) As tCoordinate
    Dim oPoint As tCoordinate, oCentre As tCoordinate, oRotated As tCoordinate
    With poShape
        oCentre.x = .Left + (.Width / 2)
        oCentre.y = .Top + (.Height / 2)
        
        oPoint.x = .Left + .Width
        oPoint.y = .Top + .Height
        
        oRotated = rotate_point(oCentre, oPoint, .Rotation)
    End With
    
    
    getShapeBottomRight = oRotated
End Function

Function getShapeRight(poShape As Shape) As tCoordinate
    Dim oPoint As tCoordinate, oCentre As tCoordinate, oRotated As tCoordinate
    With poShape
        oCentre.x = .Left + (.Width / 2)
        oCentre.y = .Top + (.Height / 2)
        
        oPoint.x = .Left + .Width
        oPoint.y = .Top + .Height / 2
        
        oRotated = rotate_point(oCentre, oPoint, .Rotation)
    End With
    
    
    getShapeRight = oRotated
End Function

Function getShapeLeft(poShape As Shape) As tCoordinate
    Dim oPoint As tCoordinate, oCentre As tCoordinate, oRotated As tCoordinate
    With poShape
        oCentre.x = .Left + (.Width / 2)
        oCentre.y = .Top + (.Height / 2)
        
        oPoint.x = .Left
        oPoint.y = .Top + .Height / 2
        
        oRotated = rotate_point(oCentre, oPoint, .Rotation)
    End With
    
    
    getShapeLeft = oRotated
End Function

Sub deselect_shapes()
    Range("a1").Select
End Sub


Public Sub remove_all_buttons()
    On Error Resume Next
        remove_collapse_button
        remove_edge_button
        remove_link_button
        remove_add_node_button
    On Error GoTo 0
End Sub


Public Function get_named_value(psName$, Optional pvDefault As Variant) As Variant
    Dim vVal As Variant
    
    vVal = Range(psName).Value
    
    If IsEmpty(vVal) Or vVal = "" Then
        If IsMissing(pvDefault) Then
            get_named_value = Null
        Else
            get_named_value = pvDefault
        End If
    Else
        get_named_value = vVal
    End If
End Function
