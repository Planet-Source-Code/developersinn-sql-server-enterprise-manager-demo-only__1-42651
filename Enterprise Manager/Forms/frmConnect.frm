VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1950
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4155
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2160
         TabIndex        =   7
         Top             =   1305
         Width           =   1140
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   420
         Left            =   990
         TabIndex        =   6
         Top             =   1305
         Width           =   1140
      End
      Begin VB.TextBox txtPWD 
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   585
         Width           =   2085
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   225
         Width           =   2085
      End
      Begin VB.CheckBox chkNTAuth 
         Caption         =   "Use Windows NT Authentication"
         Height          =   330
         Left            =   360
         TabIndex        =   1
         Top             =   945
         Width           =   3435
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   240
         Left            =   225
         TabIndex        =   3
         Top             =   585
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   270
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkNTAuth_Click()
    If chkNTAuth.Value = vbChecked Then
        Me.txtName.Enabled = False
        Me.txtPWD.Enabled = False
    Else
        Me.txtName.Enabled = True
        Me.txtPWD.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    
    
    Dim isConnected As Boolean
    isConnected = mdlMain.ConnectToServer(thisServer, txtName.Text, txtPWD.Text, Me.chkNTAuth.Value)
    If isConnected Then
        Unload Me
        mIsConnected = True
    Else
        MsgBox "Connection to " & thisServer & " Failed. Make sure the User Name, Password and Server Address.", vbCritical
    End If
    
End Sub
