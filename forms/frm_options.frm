VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_options 
   Caption         =   "mindmap options"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   OleObjectBlob   =   "frm_options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_import_Click()
    on_click_import
End Sub

Private Sub cmd_border_colour_Click()
    Dim iColour&
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        iColour = ActiveWorkbook.Colors(1)
        Range(NAME_OPTIONS_BORDER_COLOUR).Value = iColour
        cmd_border_colour.BackColor = iColour
    End If
End Sub

Private Sub cmd_colour_Click()
    Dim iColour&
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        iColour = ActiveWorkbook.Colors(1)
        Range(NAME_OPTIONS_COLOUR).Value = iColour
        cmd_colour.BackColor = iColour
    End If
End Sub

Private Sub cmd_new_Click()
    btn_new_page_Click
End Sub

Private Sub cmd_rename_Click()
    onPageRenameClick
    lbl_page.Caption = current_page()
End Sub

Private Sub cmd_reset_Click()
    on_click_reset
End Sub

Private Sub cmd_save_Click()
    on_click_save
End Sub

Private Sub cmd_text_colour_Click()
    Dim iColour&
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        iColour = ActiveWorkbook.Colors(1)
        Range(NAME_OPTIONS_FONT_COLOUR).Value = iColour
        cmd_text_colour.BackColor = iColour
    End If
End Sub

Private Sub opt_box_Click()
    Range(NAME_OPTIONS_SHAPE).Value = msoShapeRoundedRectangle
End Sub

Private Sub opt_oval_Click()
    Range(NAME_OPTIONS_SHAPE).Value = msoShapeOval
End Sub

Private Sub txt_border_Change()
    Dim iBorder%, bOK As Boolean
    
    bOK = True
    On Error Resume Next
        iBorder = CInt(txt_border.Text)
        If Err <> 0 Then bOK = False
    On Error GoTo 0
    
    If bOK And (iBorder < 0 Or iBorder > 5) Then bOK = False
    
    With txt_border
        If bOK Then
            .BorderColor = vbBlack
            .BorderStyle = fmBorderStyleNone
            Range(NAME_OPTIONS_BORDER).Value = iBorder
        Else
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbRed
        End If
    End With
    
End Sub

Private Sub txt_size_Change()
    Dim iSize%, bOK As Boolean
    
    bOK = True
    On Error Resume Next
        iSize = CInt(txt_size.Text)
        If Err <> 0 Then bOK = False
    On Error GoTo 0
    
    If bOK And (iSize < 9 Or iSize > 20) Then bOK = False
    
    With txt_size
        If bOK Then
            .BorderColor = vbBlack
            .BorderStyle = fmBorderStyleNone
            Range(NAME_OPTIONS_FONT_SIZE).Value = iSize
        Else
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbRed
        End If
    End With
End Sub

Private Sub UserForm_Initialize()
    lbl_page.Caption = current_page()
    
    '------------------shape type
    Dim iShape As MsoAutoShapeType
    iShape = get_default_shape()
    
    If iShape = msoShapeOval Then
        opt_oval.Value = 1
    Else
        opt_box.Value = 1
    End If
    
    '-------------------------colour
    cmd_colour.BackColor = get_default_colour()
    cmd_border_colour.BackColor = get_default_border_colour()
    txt_border.Text = get_default_border_size()
    txt_size.Text = get_default_font_size()
    cmd_text_colour.BackColor = get_default_font_colour()
    
End Sub

