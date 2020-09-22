VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNewDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Database"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3915
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   1935
      TabIndex        =   10
      Top             =   1440
      Width           =   1050
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create DB"
      Height          =   420
      Left            =   900
      TabIndex        =   9
      Top             =   1440
      Width           =   1050
   End
   Begin VB.TextBox txtGrowth 
      Height          =   285
      Left            =   1935
      TabIndex        =   8
      Text            =   "5"
      Top             =   1080
      Width           =   870
   End
   Begin VB.TextBox txtLFSize 
      Height          =   285
      Left            =   1935
      TabIndex        =   7
      Text            =   "10"
      Top             =   765
      Width           =   870
   End
   Begin VB.TextBox txtDFSize 
      Height          =   285
      Left            =   1935
      TabIndex        =   6
      Text            =   "10"
      Top             =   450
      Width           =   870
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1935
      TabIndex        =   5
      Top             =   135
      Width           =   1860
   End
   Begin VB.CheckBox Check1 
      Caption         =   "In MBs"
      Height          =   240
      Left            =   2880
      TabIndex        =   4
      Top             =   1125
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "MB"
      Height          =   240
      Index           =   5
      Left            =   2835
      TabIndex        =   12
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "MB"
      Height          =   240
      Index           =   4
      Left            =   2835
      TabIndex        =   11
      Top             =   495
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "File Growth"
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   3
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Log File Init Size"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   720
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Data File Init Size"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   450
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "New DB Name"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   1770
   End
End
Attribute VB_Name = "frmNewDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "Enter Database Name", vbCritical
        txtName.SetFocus
        Exit Sub
    ElseIf Val(Me.txtDFSize.Text) < 1 Then
        MsgBox "Invalid Data File Size", vbCritical
        txtDFSize.SetFocus
        Exit Sub
    ElseIf Val(Me.txtLFSize.Text) < 1 Then
        MsgBox "Invalid Log File Size", vbCritical
        txtLFSize.SetFocus
        Exit Sub
    ElseIf Val(Me.txtGrowth.Text) < 1 Then
        MsgBox "Invalid File Growth Size", vbCritical
        txtGrowth.SetFocus
        Exit Sub
    End If
    If Val(Me.txtGrowth.Text) > 100 Or Val(Me.txtGrowth.Text) < 1 Then
        MsgBox "Invalid File Growth size, can not exceed 100 and less than 1", vbCritical
        txtGrowth.SetFocus
        Exit Sub
    End If
    'OK, Now Create Databse
    If CreateDatabase Then
        mdlMain.GetAllDataBases
        Unload Me
        Exit Sub
    End If
End Sub


Private Function CreateDatabase() As Boolean
    'Use the Same SErver Object we already have
    'On Error GoTo CreateError
    Dim mDB As New SQLDMO.Database
    Dim mDBFile As New SQLDMO.DBFile
    Dim mLogFile As New SQLDMO.LogFile
    
    mDB.Name = txtName.Text     'Set the DB Name
    
    'SEt properties of DB File
    With mDBFile
        .Name = txtName.Text
        .PhysicalName = mServer.Registry.SQLDataRoot & "\DATA\" & txtName.Text & ".mdf"
        .PrimaryFile = True
        .Size = Val(Me.txtDFSize.Text)
        If Me.Check1.Value = vbChecked Then
            .FileGrowthType = SQLDMOGrowth_MB
        Else
            .FileGrowthType = SQLDMOGrowth_Percent
        End If
        .FileGrowth = Val(Me.txtGrowth.Text)
    End With
    'Add to dB Object
    mDB.FileGroups.Item("PRIMARY").DBFiles.Add mDBFile
    
    'Define the DB Transaction Log File
    mLogFile.Name = txtName.Text & "log"      'Set Name
    'Set Physical Path
    mLogFile.PhysicalName = mServer.Registry.SQLDataRoot & "\DATA\" & mLogFile.Name & ".log"
    mLogFile.Size = Val(Me.txtLFSize.Text)    'Initial Size, Optional
    If Me.Check1.Value = vbChecked Then
        mLogFile.FileGrowthType = SQLDMOGrowth_MB
    Else
        mLogFile.FileGrowthType = SQLDMOGrowth_Percent
    End If
    mLogFile.FileGrowth = Val(Me.txtGrowth.Text)

    'Add to  DB Object
    mDB.TransactionLog.LogFiles.Add mLogFile
    
    mServer.Databases.Add mDB
    MsgBox "Database Created Successfully.", vbInformation
    
    CreateDatabase = True
    Exit Function
CreateError:
    MsgBox "Error in creating database. " & Err.Description, vbCritical
End Function
