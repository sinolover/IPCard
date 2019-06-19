VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印对话框"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmPrint_old.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtPrintTile 
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
      Left            =   1065
      TabIndex        =   1
      Top             =   210
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   450
      Left            =   3570
      TabIndex        =   10
      Top             =   2250
      Width           =   1020
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "字体"
      Height          =   450
      Left            =   3570
      TabIndex        =   8
      Top             =   910
      Width           =   1020
   End
   Begin VB.OptionButton optCurrent 
      Caption         =   "从当前记录开始"
      Height          =   405
      Left            =   420
      TabIndex        =   3
      Top             =   1350
      Width           =   1755
   End
   Begin VB.OptionButton optAll 
      Caption         =   "打印所有记录"
      Height          =   360
      Left            =   435
      TabIndex        =   2
      Top             =   930
      Value           =   -1  'True
      Width           =   1770
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "开始"
      Height          =   450
      Left            =   3570
      TabIndex        =   9
      Top             =   1580
      Width           =   1020
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "打印当前日期"
      Height          =   315
      Left            =   2010
      TabIndex        =   6
      Top             =   2790
      Width           =   1455
   End
   Begin VB.CheckBox chkBG 
      Caption         =   "打印表格"
      Height          =   285
      Left            =   345
      TabIndex        =   5
      Top             =   2775
      Width           =   1080
   End
   Begin VB.TextBox Text1 
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
      Left            =   660
      TabIndex        =   4
      Top             =   1890
      Width           =   1245
   End
   Begin VB.CommandButton cmdSetFont 
      Caption         =   "字体"
      Height          =   450
      Left            =   3570
      TabIndex        =   7
      Top             =   240
      Width           =   1020
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   2715
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "D:\Source\VB\IPCard\addmoney2.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   20
   End
   Begin VB.Label Label2 
      Caption         =   "条"
      Height          =   255
      Left            =   2085
      TabIndex        =   11
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "打印标题"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   285
      Width           =   780
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

