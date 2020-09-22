VERSION 5.00
Begin VB.Form frmNewSP 
   Caption         =   "New Stored Procedure"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Your_Proc_Name"
      Top             =   135
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2655
      TabIndex        =   2
      Top             =   3915
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   1530
      TabIndex        =   1
      Top             =   3915
      Width           =   1140
   End
   Begin VB.TextBox txtSP 
      Height          =   2895
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmNewSP.frx":0000
      Top             =   765
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Procedure Name"
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   180
      Width           =   1545
   End
End
Attribute VB_Name = "frmNewSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim mProc As New SQLDMO.StoredProcedure
    On Error GoTo ProcedureError
    With mProc
        'Get the Name of the SP
        .Name = txtName.Text
        .Text = "Create Procedure " & .Name & " As " & Me.txtSP.Text
        MsgBox .Text
    End With
    
    mServer.Databases(thisDB).StoredProcedures.Add mProc
    MsgBox "Stored Procedure Created Successfully.", vbInformation
    
    mdlMain.GetAllStoredProcedures thisDB
    Unload Me
    Exit Sub
ProcedureError:
    MsgBox "Error during creation or Stored Procedure. " & Err.Description, vbCritical
    
End Sub
