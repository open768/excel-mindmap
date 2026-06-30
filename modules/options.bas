Attribute VB_Name = "options"
Option Explicit

Public Const NAME_OPTIONS_SHAPE = "default_shape"
Public Const NAME_OPTIONS_BORDER = "default_border"
Public Const NAME_OPTIONS_COLOUR = "default_colour"
Public Const NAME_OPTIONS_BORDER_COLOUR = "default_border_colour"
Public Const NAME_OPTIONS_FONT_COLOUR = "default_font_colour"
Public Const NAME_OPTIONS_FONT_SIZE = "default_font_size"

Public Const NAME_VERSION = "version"
Public Const NAME_APP = "app"

Public DEFAULT_COLOUR&

Public Function get_default_shape() As MsoAutoShapeType
    Dim iShape As MsoAutoShapeType
    On Error Resume Next
        iShape = get_named_value(NAME_OPTIONS_SHAPE, msoShapeOval)
        If Err.Number <> 0 Then iShape = msoShapeOval
    On Error GoTo 0
    get_default_shape = iShape
End Function

Public Function get_default_colour() As Long
    Dim iColour&, iDefaultColour&
    
    iDefaultColour = RGB(173, 216, 230)
    On Error Resume Next
        iColour = CLng(get_named_value(NAME_OPTIONS_COLOUR, iDefaultColour))
        If iColour = 0 Or Err.Number <> 0 Then iColour = iDefaultColour     ' light blue
    On Error GoTo 0
    
    get_default_colour = iColour
End Function

Public Function get_default_border_colour() As Long
    Dim iColour&, iDefaultColour&
    
    On Error Resume Next
        iColour = CLng(get_named_value(NAME_OPTIONS_BORDER_COLOUR, 0))
        If Err.Number <> 0 Then iColour = 0    ' light blue
    On Error GoTo 0
    
    get_default_border_colour = iColour
End Function
Public Function get_default_font_colour() As Long
    Dim iColour&, iDefaultColour&
    
    On Error Resume Next
        iColour = CLng(get_named_value(NAME_OPTIONS_FONT_COLOUR, 0))
        If Err.Number <> 0 Then iColour = 0    ' black
    On Error GoTo 0
    
    get_default_font_colour = iColour
End Function

Public Function get_default_border_size() As Integer
    Dim iBorder%
    On Error Resume Next
        iBorder = CInt(Range(NAME_OPTIONS_BORDER).Text)
        If Err.Number <> 0 Then iBorder = 0
    On Error GoTo 0
    If iBorder < 0 Or iBorder > 5 Then iBorder = 0
    get_default_border_size = iBorder
End Function

Public Function get_default_font_size() As Integer
    Dim iSize%
    On Error Resume Next
        iSize = CInt(Range(NAME_OPTIONS_FONT_SIZE).Text)
        If Err.Number <> 0 Then iSize = 12
    On Error GoTo 0
    If iSize <= 0 Then iSize = 12
    If iSize < 9 Then iSize = 9
    If iSize > 20 Then iSize = 20
    get_default_font_size = iSize
End Function

Public Sub btn_options_click()
    remove_all_buttons
    Range("a9").Select
    frm_options.Show vbModal
End Sub

Public Sub on_click_search()
    remove_all_buttons
    Range("c9").Select
    frm_search.Show vbModal
End Sub
