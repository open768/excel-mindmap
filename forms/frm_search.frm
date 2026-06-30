VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_search 
   Caption         =   "mindmap search"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   OleObjectBlob   =   "frm_search.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_search_Click()
    Dim oResults As Collection
    Dim oHistory As New Scripting.Dictionary
    
    lst_results.Clear
    Set oResults = search_all_nodes(txt_search.Value)
    If oResults Is Nothing Or oResults.Count = 0 Then
        MsgBox "no results found", vbInformation + vbOKOnly, APP_TITLE & "search"
        Exit Sub
    End If
    
    With lst_results
        .ColumnCount = 2
        .ColumnWidths = "150 pt;" & .Width - 150 & " pt"
        
        ' Header row
        .AddItem "-Match-"
        .List(.ListCount - 1, 1) = "- Page Name -"
    End With

    ' Add results
    Dim oItem As cSearchResult
    For Each oItem In oResults
        Dim sKey$
        sKey = oItem.match & oItem.page_name
        If Not oHistory.Exists(sKey) Then
            lst_results.AddItem oItem.match
            lst_results.List(lst_results.ListCount - 1, 1) = oItem.page_name
            oHistory.Add sKey, 1
        End If
    Next oItem
End Sub

Private Sub lst_results_Click()
    Dim iRow As Long
    
    iRow = lst_results.ListIndex
    
    ' Ignore header row (row 0)
    If iRow <= 0 Then Exit Sub
    
    frm_search.Hide
    Dim sPage$
    sPage = lst_results.List(iRow, 1)
    select_page_in_list sPage
End Sub
