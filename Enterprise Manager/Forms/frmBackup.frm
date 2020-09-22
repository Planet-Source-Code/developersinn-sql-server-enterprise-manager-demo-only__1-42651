VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup 
   Caption         =   "Backup/Generate SQL"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "DB Backup Options"
      Height          =   1455
      Left            =   90
      TabIndex        =   6
      Top             =   1035
      Width           =   2310
      Begin VB.OptionButton optLog 
         Caption         =   "Database Log"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   900
         Width           =   1860
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "Differntial"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   607
         Width           =   1770
      End
      Begin VB.OptionButton optComplete 
         Caption         =   "Database Complete"
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Save To"
      Height          =   330
      Left            =   4860
      TabIndex        =   3
      Top             =   225
      Width           =   960
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   6210
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   2700
      TabIndex        =   2
      Top             =   1665
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4365
      TabIndex        =   1
      Top             =   2160
      Width           =   960
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label lblMsg 
      Height          =   600
      Left            =   2790
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lblPath 
      Caption         =   "C:\"
      Height          =   600
      Left            =   495
      TabIndex        =   4
      Top             =   225
      Width           =   4335
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mBackUp As SQLDMO.Backup
Attribute mBackUp.VB_VarHelpID = -1

Private Sub cmdBackup_Click()
    Set mBackUp = New SQLDMO.Backup
    mBackUp.Database = thisDB   'Set the DB Name to backup
    On Error GoTo BackUpError
    With mBackUp
        'Complete Database
        If Me.optComplete.Value Then .Action = SQLDMOBackup_Database
        'Differntial Database
        If Me.optDiff.Value Then .Action = SQLDMOBackup_Differential
        'Log Database
        If Me.optLog.Value Then .Action = SQLDMOBackup_Log
        
        .Files = lblPath.Caption        'Where to Save
        Me.ProgressBar1.Visible = True
        Me.ProgressBar1.Max = 100
        Me.ProgressBar1.Min = 0
        
        .SQLBackup mServer
    End With
    Exit Sub
BackUpError:
    MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub cmdBrowse_Click()
    With dlgSave
        .FileName = ""
        .ShowSave
        If .FileName = "" Then Exit Sub
        
        Me.lblPath.Caption = .FileName
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    lblPath.Caption = "C:\SQLBackup\" & thisDB & ".bck"
End Sub

Private Sub mBackUp_Complete(ByVal Message As String)
    MsgBox "DataBase Backup is completed.", vbInformation
    Me.ProgressBar1.Visible = False
    Me.lblMsg.Visible = False
End Sub

Private Sub mBackUp_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    Me.lblMsg.Caption = Message
'    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.Value = Percent
End Sub
