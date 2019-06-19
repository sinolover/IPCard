VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
CenterForm Me
SetMousePtr vbHourglass
'DoEvents
End Sub
Public Sub ShowMsg(strMsg As String)
On Error Resume Next
Label1.Caption = Trim(strMsg)
Label1.Left = Int((Me.Width - Label1.Width) / 2)
Label1.Top = Int((Me.Height - Label1.Height) / 2)
'Me.SetFocus
frmStatus.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetMousePtr
End Sub
