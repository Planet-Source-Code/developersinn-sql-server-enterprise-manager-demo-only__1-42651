VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Enterprise Manager (Demo only)"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4545
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "dflt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":031A
            Key             =   "udt"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0634
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":094E
            Key             =   "sps"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C68
            Key             =   "table"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F82
            Key             =   "role"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":129C
            Key             =   "udf"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D0
            Key             =   "user"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BEA
            Key             =   "view"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F04
            Key             =   "db"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3195
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3186
            Key             =   "view"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3322
            Key             =   "role"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34CB
            Key             =   "udf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3760
            Key             =   "table"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39F6
            Key             =   "sps"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F38
            Key             =   "db"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41DB
            Key             =   "user"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4604
            Key             =   "dflt"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A39
            Key             =   "udt"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E6F
            Key             =   "rule"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   7440
      Begin MSComctlLib.ListView lvMain 
         Height          =   4305
         Left            =   2835
         TabIndex        =   7
         Top             =   0
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   7594
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Owener"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Created"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView tvmain 
         Height          =   3990
         Left            =   45
         TabIndex        =   8
         Top             =   135
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   7038
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5850
         TabIndex        =   5
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   285
         Left            =   4725
         TabIndex        =   4
         Top             =   135
         Width           =   1140
      End
      Begin VB.ComboBox cmbServers 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Top             =   135
         Width           =   2985
      End
      Begin VB.Label Label1 
         Caption         =   "Available Servers"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   195
         Width           =   1635
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12876
            MinWidth        =   12876
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "status"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
      Begin VB.Menu mnuNewDB 
         Caption         =   "&Data Base"
      End
      Begin VB.Menu mnuNewTable 
         Caption         =   "Table"
      End
      Begin VB.Menu mnuNEwSP 
         Caption         =   "Stored Procedure"
      End
   End
   Begin VB.Menu mnuMaitain 
      Caption         =   "Maintanance"
      Begin VB.Menu mnuBackupDB 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteDB 
         Caption         =   "Delete Database"
      End
      Begin VB.Menu mnuDeleteTable 
         Caption         =   "Delete Table"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuAboutAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
    
    If cmbServers.Text = "" Then
        MsgBox "Select a Server or type the Address of Server.", vbCritical
        Exit Sub
    End If
    mdlMain.thisServer = cmbServers.Text
    Load frmConnect
    frmConnect.Caption = "Connect to " & thisServer
    frmConnect.Frame1.Caption = frmConnect.Caption
    frmConnect.Show vbModal
    
    If mdlMain.mIsConnected Then
        Me.sbStatus.Panels("status").Text = "Connected to " & thisServer
        'List all objects under this Server
        mdlHelper.AddInitialItemsToTreeView
        Me.sbStatus.Panels("status") = "Getting Databases from Server " & thisServer
        mdlMain.GetAllDataBases
        Me.sbStatus.Panels("status") = "Done"
        Me.cmdConnect.Enabled = False
        Me.cmdDisconnect.Enabled = True
    End If
End Sub

Private Sub cmdDisconnect_Click()
    If mdlMain.DisconnectFromServer Then
        Me.cmdConnect.Enabled = True
        Me.cmdDisconnect.Enabled = False
        Me.lvMain.ListItems.Clear
        Me.tvmain.Nodes.Clear
    End If
    
End Sub

Private Sub Form_Resize()
    'REsize The Controls
    On Error Resume Next
    With Me.Frame2
        .Top = Frame1.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - Frame1.Height - Me.sbStatus.Height
    End With
    With Me.lvMain
        .Left = Me.tvmain.Width + 10
        .Top = Frame1.Top
        .Width = Frame2.Width - Me.tvmain.Width
        .Height = Frame2.Height
    End With
    With Me.tvmain
        .Left = Frame2.Left
        .Top = Frame1.Top
        .Height = Frame2.Height
    End With
    
    Me.lvMain.Refresh
End Sub

Private Sub lvMain_DblClick()
    'Check Which was selected
    On Error Resume Next
    Dim thisKey As String
    thisKey = lvMain.SelectedItem.Key
    If Err Then Exit Sub            'No Item was selected
    thisKey = Left(thisKey, 12)
    Dim thisDB As String
    thisDB = lvMain.SelectedItem.Key
    thisDB = Mid(thisDB, 13)
    Select Case thisKey
        Case "dbroot_tble_"
            mdlMain.GetAllTables thisDB
        Case "dbroot_view_"
            mdlMain.GetAllViews thisDB
        Case "dbroot_stps_"
            mdlMain.GetAllStoredProcedures thisDB
        Case "dbroot_user_"
            mdlMain.GetAllUsers thisDB
        Case "dbroot_role_"
            mdlMain.GetAllRoles thisDB
        Case "dbroot_rule_"
            mdlMain.GetAllRules thisDB
        Case "dbroot_dflt_"
            mdlMain.GetAllDefaults thisDB
        Case "dbroot_udts_"
            mdlMain.GetAllUserDefinedDataTypes thisDB
        Case "dbroot_udfs_"
            mdlMain.GetAllUserDefinedFunctions thisDB
    End Select
   
    tvmain.Nodes(thisKey & thisDB).Selected = True
    tvmain.Nodes(thisKey & thisDB).Expanded = True
End Sub

Private Sub mnuAboutAbout_Click()
    MsgBox "Created by Sameers (theAngrycodeR)" & vbCrLf & "theAngrycoder@Yahoo.com", vbInformation
End Sub

Private Sub mnuBackupDB_Click()
    'Check if the DB is selected
    Dim thisKey As String
    On Error Resume Next
    
    thisKey = tvmain.SelectedItem.Key
    If tvmain.SelectedItem.Parent.Key = "DataBases" Then
        'Some DataBase was selected, Get it's name
        thisDB = tvmain.SelectedItem.Key
        If thisDB = "" Or thisDB = "DataBases" Then
            MsgBox "Select a database to create backup", vbCritical
            Exit Sub
        End If
        frmBackup.Show vbModal
    Else
        MsgBox "Select a database to create it's backup.", vbCritical
    End If
End Sub

Private Sub mnuDeleteDB_Click()
    Dim thisKey As String
    On Error Resume Next
    
    thisKey = tvmain.SelectedItem.Key
    If tvmain.SelectedItem.Parent.Key = "DataBases" Then
        'Some DataBase was selected, Get it's name
        If MsgBox("Are you sure you want to delete DataBase " & thisKey, vbYesNo) = vbYes Then
            If mdlMain.DeleteDataBase(thisKey) Then
                MsgBox "Database " & thisKey & " Deleted Successfully.", vbInformation
                mdlMain.GetAllDataBases
            End If
        End If
    Else
        MsgBox "Select a database to Delete.", vbCritical
    End If
End Sub

Private Sub mnuDeleteTable_Click()
    'Check if Some table is seleted
    ' Key is in   tbleroot_DBName_TableName Format
    On Error Resume Next
    Dim mKey As String
    mKey = lvMain.SelectedItem.Key
    If mKey = "" Then
        MsgBox "Select some table to delete.", vbCritical
        Exit Sub
    End If
    If Left(mKey, 9) <> "tbleroot_" Then
        MsgBox "Select some table to delete.", vbCritical
        Exit Sub
    End If
    
    'OK some table is seleted, Get the Table and DB Name
    Dim mTableName As String, mDBName As String
    
    mTableName = lvMain.SelectedItem.Text
    mDBName = tvmain.SelectedItem.Parent.Text
    If MsgBox("Are you sure you want to delete Table " & mTableName, vbYesNo) = vbYes Then
        If DeleteTable(mTableName, mDBName) Then
            MsgBox "Table Delete Successfully.", vbInformation
            mdlMain.GetAllTables mDBName
        End If
    End If
End Sub

Private Sub mnuNewDB_Click()
    frmNewDB.Show vbModal
End Sub

Private Sub mnuNEwSP_Click()
    Dim thisKey As String
    On Error Resume Next
    
    thisKey = tvmain.SelectedItem.Key
    If tvmain.SelectedItem.Parent.Key = "DataBases" Then
        'Some DataBase was selected, Get it's name
        thisDB = tvmain.SelectedItem.Key
        frmNewSP.Show vbModal
    ElseIf Left(thisKey, 12) = "dbroot_stps_" Then
        thisDB = tvmain.SelectedItem.Parent.Key
        frmNewSP.Show vbModal
    Else
        MsgBox "Invalid place to Add New Procedure. Select Stored Procedure From Tree View or Database and then try again.", vbCritical
    End If
End Sub

Private Sub mnuNewTable_Click()
    Dim thisKey As String
    On Error Resume Next
    
    thisKey = tvmain.SelectedItem.Key
    If tvmain.SelectedItem.Parent.Key = "DataBases" Then
        'Some DataBase was selected, Get it's name
        thisDB = tvmain.SelectedItem.Key
        frmNewTable.Show vbModal
    ElseIf Left(thisKey, 12) = "dbroot_tble_" Then
        thisDB = tvmain.SelectedItem.Parent.Key
        frmNewTable.Show vbModal
    Else
        MsgBox "Invalid place to Add New Table. Select Tables From Tree View and then try again.", vbCritical
    End If
    
End Sub

Private Sub mnuRestore_Click()
    'Check if the DB is selected
    Dim thisKey As String
    On Error Resume Next
    
    thisKey = tvmain.SelectedItem.Key
    If tvmain.SelectedItem.Parent.Key = "DataBases" Then
        thisDB = tvmain.SelectedItem.Key
        If thisDB = "" Or thisDB = "DataBases" Then
            MsgBox "Select a database to create backup", vbCritical
            Exit Sub
        End If
        
        'Some DataBase was selected, Get it's name
        
        frmRestore.Show vbModal
    Else
        MsgBox "Select a database to Restore backup.", vbCritical
    End If
End Sub

'Private Sub tvmain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
'    If Button = vbRightButton Then
'        Dim thisKey As String
'        thisKey = tvmain.SelectedItem.Key
'        If thisKey = "DataBases" Then PopupMenu mnutemp, , x, y: Exit Sub
'        If Left(thisKey, 12) = "dbroot_tble_" Then
'            Me.mnuNew.Caption = "New Table"
'            PopupMenu mnutemp, , x, y
'        End If
'    End If
'End Sub

Private Sub tvmain_NodeClick(ByVal Node As MSComctlLib.Node)
    'Process Node, Like, If table Node clicked, then Show all tables
    If Node.Key = "DataBases" Then
        'DataBase Root Node was clicked, List all DBs in the ListView
        GetAllDataBases
        Exit Sub
    End If
    
    'Check if Clicked on the Root of DB
    If Node.Parent.Key = "DataBases" Then
        'Show all Objects in the ListView
        mdlMain.ShowObjectsInListView Node.Key
        Exit Sub
    End If
    Dim thisKey As String
    thisKey = Left(Node.Key, 12)
    Screen.MousePointer = vbHourglass
    Select Case thisKey
        Case "dbroot_tble_"
            mdlMain.GetAllTables Node.Parent.Key
        Case "dbroot_view_"
            mdlMain.GetAllViews Node.Parent.Key
        Case "dbroot_stps_"
            mdlMain.GetAllStoredProcedures Node.Parent.Key
        Case "dbroot_user_"
            mdlMain.GetAllUsers Node.Parent.Key
        Case "dbroot_role_"
            mdlMain.GetAllRoles Node.Parent.Key
        Case "dbroot_rule_"
            mdlMain.GetAllRules Node.Parent.Key
        Case "dbroot_dflt_"
            mdlMain.GetAllDefaults Node.Parent.Key
        Case "dbroot_udts_"
            mdlMain.GetAllUserDefinedDataTypes Node.Parent.Key
        Case "dbroot_udfs_"
            mdlMain.GetAllUserDefinedFunctions Node.Parent.Key
    End Select
    Node.Expanded = True
Screen.MousePointer = vbNormal
End Sub
