VERSION 5.00
Begin VB.Form frmNewTable 
   Caption         =   "New Table"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTableName 
      Height          =   285
      Left            =   2025
      TabIndex        =   38
      Top             =   720
      Width           =   2490
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2970
      TabIndex        =   37
      Top             =   3375
      Width           =   1410
   End
   Begin VB.CheckBox chkID 
      Height          =   240
      Index           =   4
      Left            =   4860
      TabIndex        =   36
      Top             =   2955
      Width           =   330
   End
   Begin VB.CheckBox chkID 
      Height          =   240
      Index           =   3
      Left            =   4860
      TabIndex        =   35
      Top             =   2655
      Width           =   330
   End
   Begin VB.CheckBox chkID 
      Height          =   240
      Index           =   2
      Left            =   4860
      TabIndex        =   34
      Top             =   2340
      Width           =   330
   End
   Begin VB.CheckBox chkID 
      Height          =   240
      Index           =   1
      Left            =   4860
      TabIndex        =   33
      Top             =   1980
      Width           =   330
   End
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      Index           =   4
      ItemData        =   "frmNewTable.frx":0000
      Left            =   3600
      List            =   "frmNewTable.frx":0016
      TabIndex        =   32
      Top             =   2910
      Width           =   1050
   End
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      Index           =   3
      ItemData        =   "frmNewTable.frx":004A
      Left            =   3600
      List            =   "frmNewTable.frx":0060
      TabIndex        =   31
      Top             =   2625
      Width           =   1050
   End
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      Index           =   2
      ItemData        =   "frmNewTable.frx":0094
      Left            =   3600
      List            =   "frmNewTable.frx":00AA
      TabIndex        =   30
      Top             =   2310
      Width           =   1050
   End
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      Index           =   1
      ItemData        =   "frmNewTable.frx":00DE
      Left            =   3600
      List            =   "frmNewTable.frx":00F4
      TabIndex        =   29
      Top             =   1950
      Width           =   1050
   End
   Begin VB.CheckBox chkNull 
      Height          =   240
      Index           =   4
      Left            =   3195
      TabIndex        =   28
      Top             =   2955
      Width           =   330
   End
   Begin VB.CheckBox chkNull 
      Height          =   240
      Index           =   3
      Left            =   3195
      TabIndex        =   27
      Top             =   2655
      Width           =   330
   End
   Begin VB.CheckBox chkNull 
      Height          =   240
      Index           =   2
      Left            =   3195
      TabIndex        =   26
      Top             =   2340
      Width           =   330
   End
   Begin VB.CheckBox chkNull 
      Height          =   240
      Index           =   1
      Left            =   3195
      TabIndex        =   25
      Top             =   1980
      Width           =   330
   End
   Begin VB.CheckBox chkPrimary 
      Height          =   240
      Index           =   4
      Left            =   2700
      TabIndex        =   24
      Top             =   2955
      Width           =   330
   End
   Begin VB.CheckBox chkPrimary 
      Height          =   240
      Index           =   3
      Left            =   2700
      TabIndex        =   23
      Top             =   2655
      Width           =   330
   End
   Begin VB.CheckBox chkPrimary 
      Height          =   240
      Index           =   2
      Left            =   2700
      TabIndex        =   22
      Top             =   2340
      Width           =   330
   End
   Begin VB.CheckBox chkPrimary 
      Height          =   240
      Index           =   1
      Left            =   2700
      TabIndex        =   21
      Top             =   1980
      Width           =   330
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   20
      Top             =   2925
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   3
      Left            =   1260
      TabIndex        =   19
      Top             =   2640
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   2
      Left            =   1260
      TabIndex        =   18
      Top             =   2325
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   17
      Top             =   1965
      Width           =   1365
   End
   Begin VB.CheckBox chkID 
      Height          =   240
      Index           =   0
      Left            =   4860
      TabIndex        =   16
      Top             =   1620
      Width           =   330
   End
   Begin VB.CheckBox chkPrimary 
      Height          =   240
      Index           =   0
      Left            =   2700
      TabIndex        =   14
      Top             =   1620
      Width           =   330
   End
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      Index           =   0
      ItemData        =   "frmNewTable.frx":0128
      Left            =   3600
      List            =   "frmNewTable.frx":013E
      TabIndex        =   13
      Top             =   1590
      Width           =   1050
   End
   Begin VB.CheckBox chkNull 
      Height          =   240
      Index           =   0
      Left            =   3195
      TabIndex        =   11
      Top             =   1620
      Width           =   330
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   0
      Left            =   1260
      TabIndex        =   7
      Top             =   1605
      Width           =   1365
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   420
      Left            =   1530
      TabIndex        =   1
      Top             =   3375
      Width           =   1410
   End
   Begin VB.Label Label5 
      Caption         =   "Table Name"
      Height          =   195
      Left            =   315
      TabIndex        =   39
      Top             =   765
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Identity"
      Height          =   240
      Index           =   8
      Left            =   4770
      TabIndex        =   15
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Data Type"
      Height          =   285
      Left            =   3690
      TabIndex        =   12
      Top             =   1305
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Null"
      Height          =   240
      Index           =   6
      Left            =   3195
      TabIndex        =   10
      Top             =   1305
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Primary"
      Height          =   240
      Index           =   7
      Left            =   2610
      TabIndex        =   9
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "Field Name"
      Height          =   240
      Left            =   1305
      TabIndex        =   8
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Field 5"
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   2955
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Field 4"
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   2655
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Field 3"
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   2340
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Field 2"
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1980
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Field 1"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is a Demo Only that how to Create Tables. So no full functions available here."
      Height          =   465
      Left            =   855
      TabIndex        =   0
      Top             =   135
      Width           =   3660
   End
End
Attribute VB_Name = "frmNewTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkID_Click(Index As Integer)
    If Me.chkID(Index).Value = vbChecked Then
        'UnCheck all others
        Dim temp As Integer
        For temp = 0 To Me.chkID.UBound
            If temp <> Index Then _
                chkID(temp).Value = vbUnchecked
        Next
        
    End If
End Sub

Private Sub chkPrimary_Click(Index As Integer)
    If Me.chkPrimary(Index).Value = vbChecked Then
        'UnCheck all others
        Dim temp As Integer
        For temp = 0 To Me.chkPrimary.UBound
            If temp <> Index Then _
                chkPrimary(temp).Value = vbUnchecked
        Next
        
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    If Not ValidateForm Then Exit Sub
        
    On Error GoTo AddTablError
    
    Dim mTable As New SQLDMO.Table
    Dim mCol As SQLDMO.Column
    Dim mKey As New SQLDMO.Key
    
    'mTable.BeginAlter
    
    
    Dim temp As Integer     'Iterate through each Field you typed in the form
    For temp = 0 To Me.txtName.UBound
        Set mCol = New Column
        If txtName(temp).Text = "" Then Exit For
        Set mCol = New SQLDMO.Column
        mCol.Name = txtName(temp).Text
        mCol.AllowNulls = CBool(Me.chkNull(temp).Value) 'Allow Null ?
        mCol.Identity = CBool(Me.chkID(temp).Value)     'Is Identity
        
        mCol.Datatype = Me.cmbDBType(temp).Text     'DAtaType of Column
        Select Case Me.cmbDBType(temp).Text
            Case "Bit"
                mCol.Length = 1
            Case "DateTime"
                mCol.Length = 8
            Case "Decimal"
                mCol.Length = 9
            Case "Int"
                mCol.Length = 4
            Case "Numeric"
                mCol.Length = 9
            Case "nvarchar"
                mCol.Length = 50
        End Select
        If Me.chkPrimary(temp).Value = vbChecked Then
            mKey.Type = SQLDMOKey_Primary
            mKey.KeyColumns.Add txtName(temp).Text

        End If
        mTable.Columns.Add mCol
    Next
    mTable.Keys.Add mKey
    
    mTable.Name = Me.txtTableName.Text
    'mTable.DoAlter
    mServer.Databases(thisDB).Tables.Add mTable
    MsgBox "Table " & mTable.Name & " has been added successfully to " & thisDB & " Database.", vbInformation
    'Refresh the Tables List
    GetAllTables thisDB
    Unload Me
AddTablError:
    MsgBox "Error in adding new table. Please review your entries.", vbCritical
End Sub

Private Function ValidateForm() As Boolean
    Dim temp As Integer
    For temp = 0 To txtName.UBound
        If txtName(temp).Text = "" Then
            If temp = 0 Then
                MsgBox "You should type field names in Sequence. Atlease one field is required.", vbCritical
                ValidateForm = False
                Exit Function
            Else
                ValidateForm = True
                Exit Function
            End If
        Else
            'Check if DBType is Selected
            If Me.cmbDBType(temp).Text = "" Then
                MsgBox "Select the dAtabase type for this Field", vbCritical
                cmbDBType(temp).SetFocus
                ValidateForm = False
                Exit Function
            End If
        End If
    Next
    ValidateForm = True
    Exit Function
AddTablError:
    MsgBox "Error adding new table. " & Err.Description, vbCritical
End Function
