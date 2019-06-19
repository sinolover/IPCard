VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表修改"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optV 
      Caption         =   "纵向"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1725
      TabIndex        =   2
      Top             =   765
      Width           =   1095
   End
   Begin VB.OptionButton optH 
      Caption         =   "横向"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   330
      TabIndex        =   1
      Top             =   735
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSetFont 
      Caption         =   "字体(&F)"
      Height          =   400
      Left            =   2910
      TabIndex        =   0
      Top             =   675
      Width           =   1200
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CenterForm Me
'rpt1.
End Sub
