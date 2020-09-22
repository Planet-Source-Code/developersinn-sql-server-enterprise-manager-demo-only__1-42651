Attribute VB_Name = "mdlHelper"
Option Explicit

'This module will contain the functions which may be needed for the frmMain
'It will not have any SQLDMO related function, for SQLDMO, See mdlMain


Public Sub ClearTreeView()
    With frmMain.tvmain
        Dim temp As Integer
        For temp = .Nodes.Count To 1 Step -1
            .Nodes.Remove (temp)
        Next
    End With
End Sub
Public Sub AddInitialItemsToTreeView()
    Call ClearTreeView      'First clear the tree view
    Call ClearListView      'Also Clear the List View
    
    Dim thisNode As Node
    With frmMain.tvmain
        Set thisNode = .Nodes.Add(, , "DataBases", "DataBases on " & thisServer)
        'Set thisNode = .Nodes.Add(, , "DataBases", "DataBases")
        '
    End With
    
    frmMain.Show
End Sub
Public Sub ClearListView()
    'Clear Items and Also Clear Columns
    frmMain.lvMain.ListItems.Clear
    frmMain.lvMain.ColumnHeaders.Clear
End Sub
Public Sub CreateListViewColumn(ByVal col1 As String, ByVal col2 As String, Optional ByVal col3 As String = "")
    Call ClearListView
    frmMain.lvMain.View = lvwReport
    With frmMain.lvMain
        .ColumnHeaders.Add 1, , col1, 3000
        .ColumnHeaders.Add 2, , col2, 1500
        If col3 <> "" Then .ColumnHeaders.Add 3, , col3, 3000
    End With
End Sub
Public Sub RemoveListComumns()
    Call ClearListView
    frmMain.lvMain.View = lvwIcon
End Sub
Public Sub ClearTreeViewNode(ByVal nodeKey As String)
    'Remove everyhting under this node
    Dim temp As Integer
    Dim thisNode As Node
    Set thisNode = frmMain.tvmain.Nodes(nodeKey)
    If thisNode.Children = 0 Then Exit Sub
    For temp = 1 To thisNode.Children
        frmMain.tvmain.Nodes.Remove thisNode.Child.Key
    Next
End Sub
