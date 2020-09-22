VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestore 
   Caption         =   "Restore DataBase Backup"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Replace DB if necessary"
      Height          =   285
      Left            =   1125
      TabIndex        =   6
      Top             =   1440
      Width           =   2580
   End
   Begin VB.OptionButton optFiles 
      Caption         =   "Restore Files"
      Height          =   330
      Left            =   1125
      TabIndex        =   5
      Top             =   990
      Width           =   2445
   End
   Begin VB.OptionButton optLog 
      Caption         =   "Restore Log"
      Height          =   330
      Left            =   1125
      TabIndex        =   4
      Top             =   630
      Width           =   2445
   End
   Begin VB.OptionButton optComplete 
      Caption         =   "DB Complete"
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   270
      Value           =   -1  'True
      Width           =   2445
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   4185
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   600
      Left            =   2385
      TabIndex        =   2
      Top             =   2295
      Width           =   1455
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Select and Restore"
      Height          =   600
      Left            =   945
      TabIndex        =   1
      Top             =   2295
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   675
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents mRestore As SQLDMO.Restore
Attribute mRestore.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    Set mRestore = New SQLDMO.Restore
    With dlgOpen
        .FileName = ""
        .ShowOpen
        If .FileName = "" Then Exit Sub
    End With
    
    'If Me.optComplete.Value Then mRestore.Action = SQLDMORestore_Database
    If Me.optLog.Value Then mRestore.Action = SQLDMORestore_Log
    If Me.optFiles.Value Then mRestore.Action = SQLDMORestore_Files
    
    mRestore.ReplaceDatabase = Me.Check1.Value
    mRestore.Database = thisDB
    mRestore.Files = dlgOpen.FileName
    
    Me.ProgressBar1.Visible = True
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.Min = 0
    On Error GoTo RestoreError
    mRestore.SQLRestore mServer
RestoreError:
    MsgBox "Error : " & Err.Description
    Me.ProgressBar1.Visible = False
End Sub

Private Sub Form_Load()
    Me.Caption = "Restore Backup of " & thisDB
End Sub

Private Sub mRestore_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    Me.ProgressBar1.Value = Percent
End Sub
