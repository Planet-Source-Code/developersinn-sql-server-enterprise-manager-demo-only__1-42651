Attribute VB_Name = "mdlMain"
Option Explicit

'Global Object of Server, Reused for each connection
Public mServer As SQLDMO.SQLServer
Public mIsConnected As Boolean  'Are you connected to Some Server ?
Public thisServer As String     'Current Server to which you are connected
Public thisDB As String

Public Sub main()
    'ListAllSErvers will also List all DataBases to the TreeView
    Call ListAllServers
    Load frmMain
    Set frmMain.tvmain.ImageList = frmMain.ImageList1
    Set frmMain.lvMain.SmallIcons = frmMain.ImageList1
    Set frmMain.lvMain.Icons = frmMain.ImageList2
    frmMain.Show
End Sub

Public Sub ListAllServers()
    Dim mApplication As SQLDMO.Application
    Set mApplication = New SQLDMO.Application
    Dim mNames As SQLDMO.NameList
    On Error GoTo ServerListError
    Screen.MousePointer = vbHourglass
    DoEvents
    Set mNames = mApplication.ListAvailableSQLServers
    DoEvents
    Dim temp As Integer
    Load frmMain
    For temp = 1 To mNames.Count
        frmMain.cmbServers.AddItem mNames(temp)
    Next
    
    
ServerListError:
    frmMain.Show
    Screen.MousePointer = vbNormal
    Exit Sub
End Sub
Public Sub ShowObjectsInListView(ByVal dbKey As String)
    Dim thisDB As SQLDMO.Database
    mdlHelper.RemoveListComumns
    'For Each thisDB In mServer.Databases
        frmMain.lvMain.ListItems.Add , "dbroot_tble_" & dbKey, "Tables", "table"
        frmMain.lvMain.ListItems.Add , "dbroot_view_" & dbKey, "Views", "view"
        frmMain.lvMain.ListItems.Add , "dbroot_stps_" & dbKey, "Stored Procedures", "sps"
        frmMain.lvMain.ListItems.Add , "dbroot_user_" & dbKey, "Users", "user"
        frmMain.lvMain.ListItems.Add , "dbroot_role_" & dbKey, "Roles", "role"
        frmMain.lvMain.ListItems.Add , "dbroot_rule_" & dbKey, "Rules", "rule"
        frmMain.lvMain.ListItems.Add , "dbroot_dflt_" & dbKey, "Defaults", "dflt"
        frmMain.lvMain.ListItems.Add , "dbroot_udts_" & dbKey, "User Defined Data Types", "udt"
        frmMain.lvMain.ListItems.Add , "dbroot_udfs_" & dbKey, "User Defined Functions", "udf"
    'Next
End Sub
Public Function GetAllDataBases()
    'Note that we are connected using mServer Object, Same will be used to Enum DataBases
    Dim thisDB As SQLDMO.Database
    DoEvents
    Screen.MousePointer = vbHourglass
    mdlHelper.RemoveListComumns
    ClearTreeViewNode "DataBases"
    For Each thisDB In mServer.Databases
        DoEvents
        frmMain.tvmain.Nodes.Add "DataBases", tvwChild, thisDB.Name, thisDB.Name, "db"
        'Also Add this DB in the ListView
        frmMain.lvMain.ListItems.Add , thisDB.Name, thisDB.Name, "db"
        'Add Nodes for other Objects like tables/views etc
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_tble_" & thisDB.Name, "Tables", "table"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_view_" & thisDB.Name, "Views", "view"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_stps_" & thisDB.Name, "Stored Procedures", "sps"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_user_" & thisDB.Name, "Users", "user"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_role_" & thisDB.Name, "Roles", "role"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_rule_" & thisDB.Name, "Rules", "rule"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_dflt_" & thisDB.Name, "Defaults", "dflt"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_udts_" & thisDB.Name, "User Defined Data Types", "udt"
        frmMain.tvmain.Nodes.Add thisDB.Name, tvwChild, "dbroot_udfs_" & thisDB.Name, "User Defined Functions", "udf"
        DoEvents
    Next
    Screen.MousePointer = vbNormal
End Function
Public Sub GetAllTables(ByVal mDBName As String)
    'Gets all tables under the Database
    Dim mTable As SQLDMO.Table
    'ClearTreeViewNode "dbroot_tble_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    For Each mTable In mServer.Databases(mDBName).Tables
        
         'frmMain.tvmain.Nodes.Add "dbroot_tble_" & mDBName, tvwChild, "tbleroot_" & mDBName & "_" & mTable.Name, mTable.Name
         
         Set thisItem = frmMain.lvMain.ListItems.Add(, "tbleroot_" & mDBName & "_" & mTable.Name, mTable.Name)
         thisItem.SubItems(1) = mTable.Owner
         thisItem.SubItems(2) = mTable.CreateDate
    
    Next
End Sub
Public Sub GetAllViews(ByVal mDBName As String)
    'Gets all Views under the Database
    Dim mView As SQLDMO.View
    'ClearTreeViewNode "dbroot_view_" & mDBName
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    For Each mView In mServer.Databases(mDBName).Views
        'frmMain.tvmain.Nodes.Add "dbroot_view_" & mDBName, tvwChild, "viewroot_" & mDBName & "_" & mView.Name, mView.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "viewroot_" & mDBName & "_" & mView.Name, mView.Name)
        thisItem.SubItems(1) = mView.Owner
        thisItem.SubItems(2) = mView.CreateDate
    Next
End Sub
Public Sub GetAllStoredProcedures(ByVal mDBName As String)
    'Gets all Stored Procedures under the Database
    Dim mSP As SQLDMO.StoredProcedure
    'ClearTreeViewNode "dbroot_stps_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    
    For Each mSP In mServer.Databases(mDBName).StoredProcedures
        'frmMain.tvmain.Nodes.Add "dbroot_stps_" & mDBName, tvwChild, "stpsroot_" & mDBName & "_" & mSP.Name, mSP.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "stpsroot_" & mDBName & "_" & mSP.Name, mSP.Name)
        thisItem.SubItems(1) = mSP.Owner
        thisItem.SubItems(2) = mSP.CreateDate
    
    Next
End Sub
Public Sub GetAllUsers(ByVal mDBName As String)
    'Gets all Users under the Database
    Dim mUser As SQLDMO.User
    'ClearTreeViewNode "dbroot_user_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Login Name", "DataBase Access"
    Dim thisItem As ListItem
    For Each mUser In mServer.Databases(mDBName).Users
        'frmMain.tvmain.Nodes.Add "dbroot_user_" & mDBName, tvwChild, "userroot_" & mDBName & "_" & mUser.Name, mUser.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "userroot_" & mDBName & "_" & mUser.Name, mUser.Name)
        thisItem.SubItems(1) = mUser.Login
        thisItem.SubItems(2) = mUser.HasDBAccess
    Next
End Sub
Public Sub GetAllRoles(ByVal mDBName As String)
    'Gets all Roles under the Database
    Dim mRole As SQLDMO.DatabaseRole
    'ClearTreeViewNode "dbroot_role_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Role Type"
    Dim thisItem As ListItem
    For Each mRole In mServer.Databases(mDBName).DatabaseRoles
        'frmMain.tvmain.Nodes.Add "dbroot_role_" & mDBName, tvwChild, "roleroot_" & mDBName & "_" & mRole.Name, mRole.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "roleroot_" & mDBName & "_" & mRole.Name, mRole.Name)
        thisItem.SubItems(1) = IIf(mRole.IsFixedRole, "Standard", "Non Standard")
        'thisItem.SubItems(2) = mRole.IsFixedRole
    Next
End Sub
Public Sub GetAllRules(ByVal mDBName As String)
    'Gets all tables under the Database
    Dim mRule As SQLDMO.Rule
    'ClearTreeViewNode "dbroot_rule_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    For Each mRule In mServer.Databases(mDBName).Rules
        'frmMain.tvmain.Nodes.Add "dbroot_rule_" & mDBName, tvwChild, "ruleroot_" & mDBName & "_" & mRule.Name, mRule.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "ruleroot_" & mDBName & "_" & mRule.Name, mRule.Name)
        thisItem.SubItems(1) = mRule.Owner
        thisItem.SubItems(2) = mRule.CreateDate
    
    Next
If Err Then MsgBox Err.Description
End Sub
Public Sub GetAllDefaults(ByVal mDBName As String)
    'Gets all tables under the Database
    Dim mDflt As SQLDMO.Default
    'ClearTreeViewNode "dbroot_dflt_" & mDBName
    
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    For Each mDflt In mServer.Databases(mDBName).Defaults
        'frmMain.tvmain.Nodes.Add "dbroot_dflt_" & mDBName, tvwChild, "dfltroot_" & mDBName & "_" & mDflt.Name, mDflt.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "dfltroot_" & mDBName & "_" & mDflt.Name, mDflt.Name)
        thisItem.SubItems(1) = mDflt.Owner
        thisItem.SubItems(2) = mDflt.CreateDate
    
    Next
    
End Sub
Public Sub GetAllUserDefinedDataTypes(ByVal mDBName As String)
    'Gets all tables under the Database
    Dim mUDT As SQLDMO.UserDefinedDatatype
    'ClearTreeViewNode "dbroot_udts_" & mDBName
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Base Type"
    Dim thisItem As ListItem
    For Each mUDT In mServer.Databases(mDBName).UserDefinedDatatypes
        'frmMain.tvmain.Nodes.Add "dbroot_udts_" & mDBName, tvwChild, "udtsroot_" & mDBName & "_" & mUDT.Name, mUDT.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "udtsroot_" & mDBName & "_" & mUDT.Name, mUDT.Name)
        thisItem.SubItems(1) = mUDT.Owner
        thisItem.SubItems(2) = mUDT.BaseType
    
    Next
    
End Sub
Public Sub GetAllUserDefinedFunctions(ByVal mDBName As String)
    'Gets all tables under the Database
    Dim mUDF As SQLDMO.UserDefinedFunction
    'ClearTreeViewNode "dbroot_udfs_" & mDBName
    'Add Columns by removing previous cols in Tree view. This will also remove all data in the ListView
    CreateListViewColumn "Name", "Owner", "Date Created"
    Dim thisItem As ListItem
    'Note: User Defined Functions are available in SQLServer 7.0 or later
    'DataBase Object will not let you access UserDefinedFunctions, But you can Use DataBase2 Object
    Dim mDB As SQLDMO.Database2
    Set mDB = mServer.Databases(mDBName)
    For Each mUDF In mDB.UserDefinedFunctions
        'frmMain.tvmain.Nodes.Add "dbroot_udfs_" & mDBName, tvwChild, "udfsroot_" & mDBName & "_" & mUDF.Name, mUDF.Name
        
        Set thisItem = frmMain.lvMain.ListItems.Add(, "udfsroot_" & mDBName & "_" & mUDF.Name, mUDF.Name)
        thisItem.SubItems(1) = mUDF.Owner
        thisItem.SubItems(2) = mUDF.CreateDate
    
    Next
End Sub


Public Function ConnectToServer(ByVal mAddress As String, ByVal mUName As String, ByVal mPWD As String, ByVal isNTAuth As Boolean) As Boolean
    Set mServer = New SQLDMO.SQLServer
    On Error GoTo ConnectError
    Screen.MousePointer = vbHourglass
    With mServer
        If isNTAuth Then
            'If Use NT Authenticaion then
            .LoginSecure = True
            .Connect mAddress
        Else
            .Connect mAddress, mUName, mPWD
        End If
    End With
    
    Screen.MousePointer = vbNormal
    ConnectToServer = True
    Exit Function
ConnectError:
    ConnectToServer = False
    Screen.MousePointer = vbNormal
End Function

Public Function DisconnectFromServer() As Boolean
    On Error GoTo DisconnectError
    Screen.MousePointer = vbHourglass
    mServer.DisConnect
    
    DisconnectFromServer = True
    Screen.MousePointer = vbNormal
    Exit Function
DisconnectError:
    Screen.MousePointer = vbNormal
    DisconnectFromServer = False
End Function


Public Function DeleteTable(ByVal mTableName As String, ByVal mDBName As String) As Boolean
    On Error GoTo TableDeleteError
    Screen.MousePointer = vbHourglass
    mServer.Databases(mDBName).Tables(mTableName).Remove
    Screen.MousePointer = vbNormal
    
    DeleteTable = True
    Exit Function
TableDeleteError:
    Screen.MousePointer = vbNormal
    MsgBox "Error : " & Err.Description
    DeleteTable = False
End Function
Public Function DeleteDataBase(ByVal mDBName As String) As Boolean
    On Error GoTo DBDeleteError
    Screen.MousePointer = vbHourglass
    mServer.Databases(mDBName).Remove
    Screen.MousePointer = vbNormal
    
    DeleteDataBase = True
    Exit Function
DBDeleteError:
    Screen.MousePointer = vbNormal
    MsgBox "Error : " & Err.Description
    DeleteDataBase = False
End Function
