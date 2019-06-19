VERSION 5.00
Begin VB.Form frmSysInf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统信息"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   Icon            =   "frmSysInf.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   900
      TabIndex        =   0
      Top             =   3480
      Width           =   1200
   End
End
Attribute VB_Name = "frmSysInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me

End Sub
