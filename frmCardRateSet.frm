VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCardRateSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   4635
      TabIndex        =   1
      Top             =   -15
      Width           =   1425
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   158
         TabIndex        =   5
         Top             =   2805
         Width           =   1095
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "����(&A)"
         Default         =   -1  'True
         Height          =   400
         Left            =   105
         TabIndex        =   4
         Top             =   495
         Width           =   1200
      End
      Begin VB.CommandButton cmdChange 
         Cancel          =   -1  'True
         Caption         =   "�޸�(&C)"
         Height          =   400
         Left            =   105
         TabIndex        =   3
         Top             =   1245
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   400
         Left            =   105
         TabIndex        =   2
         Top             =   1995
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   630
      Left            =   2700
      Top             =   3210
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1111
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCardRateSet.frx":0000
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   8043
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCardRateSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

