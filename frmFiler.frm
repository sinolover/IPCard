VERSION 5.00
Begin VB.Form frmFiler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "筛选"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmFiler.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   400
      Left            =   3390
      TabIndex        =   6
      Top             =   1410
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3390
      TabIndex        =   5
      Top             =   877
      Width           =   1200
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3390
      TabIndex        =   4
      Top             =   345
      Width           =   1200
   End
   Begin VB.TextBox txtCorName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   2
      Top             =   1125
      Width           =   1620
   End
   Begin VB.TextBox txtOprName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   3
      Top             =   1590
      Width           =   1620
   End
   Begin VB.TextBox txtWkrName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   1
      Top             =   660
      Width           =   1620
   End
   Begin VB.TextBox txtCountNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Top             =   195
      Width           =   1620
   End
   Begin VB.CheckBox chkOprName 
      Caption         =   "操作员"
      Height          =   315
      Left            =   315
      TabIndex        =   10
      Top             =   1620
      Width           =   1050
   End
   Begin VB.CheckBox chkCorName 
      Caption         =   "单位名称"
      Height          =   315
      Left            =   315
      TabIndex        =   9
      Top             =   1160
      Width           =   1050
   End
   Begin VB.CheckBox chkWkrName 
      Caption         =   "业务员"
      Height          =   315
      Left            =   315
      TabIndex        =   8
      Top             =   700
      Width           =   1050
   End
   Begin VB.CheckBox chkCountNO 
      Caption         =   "账号"
      Height          =   315
      Left            =   315
      TabIndex        =   7
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmFiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
strSql = ""
Unload Me
End Sub

Private Sub Find()
'If chkCountNO.Value = 1 Then
CenterForm Me
strSql = ""
If Len(Trim(txtCountNo.Text)) > 0 Then
   strSql = "Countno = '" & Trim(txtCountNo.Text) & "'"
End If
If Len(Trim(txtWkrName.Text)) > 0 Then
   If Len(strSql) > 3 Then
      strSql = " and WkrName = '" & Trim(txtWkrName.Text) & "'"
     Else
      strSql = " WkrName = '" & Trim(txtWkrName.Text) & "'"
   End If
End If
If Len(Trim(txtCorName.Text)) > 0 Then
   If Len(strSql) > 3 Then
      strSql = " and corName = '" & Trim(txtCorName.Text) & "'"
     Else
      strSql = " corName = '" & Trim(txtCorName.Text) & "'"
   End If
End If
Unload Me
End Sub
Private Sub cmdFind_Click()
'If chkCountNO.Value = 1 Then
strFilerSql = ""
If Len(Trim(txtCountNo.Text)) > 0 Then
   strFilerSql = "Countno like '" & Trim(txtCountNo.Text) & "%'"
End If
If Len(Trim(txtWkrName.Text)) > 0 Then
   If Len(strFilerSql) > 3 Then
      strFilerSql = " and WkrName like '" & Trim(txtWkrName.Text) & "%'"
     Else
      strFilerSql = " WkrName like '" & Trim(txtWkrName.Text) & "%'"
   End If
End If
If Len(Trim(txtCorName.Text)) > 0 Then
   If Len(strFilerSql) > 3 Then
      strFilerSql = " and corName like '" & Trim(txtCorName.Text) & "%'"
     Else
      strFilerSql = " corName like '" & Trim(txtCorName.Text) & "%'"
   End If
End If
Unload Me
With mfrmMain.Toolbar1
 .Buttons("Filer").Enabled = False
 .Buttons("CFiler").Enabled = True
End With
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub
