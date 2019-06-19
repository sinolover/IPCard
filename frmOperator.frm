VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOperator 
   Caption         =   "操作员信息"
   ClientHeight    =   4065
   ClientLeft      =   2835
   ClientTop       =   3450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   6045
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgOperator 
      Bindings        =   "frmOperator.frx":0000
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   40
      BeginProperty Column00 
         DataField       =   "OprNO"
         Caption         =   "OprNO"
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
         DataField       =   "OprName"
         Caption         =   "OprName"
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
      BeginProperty Column02 
         DataField       =   "OprPW"
         Caption         =   "OprPW"
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
      BeginProperty Column03 
         DataField       =   "OprType"
         Caption         =   "OprType"
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
      BeginProperty Column04 
         DataField       =   "OprRights"
         Caption         =   "OprRights"
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
      BeginProperty Column05 
         DataField       =   "OprSex"
         Caption         =   "OprSex"
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
      BeginProperty Column06 
         DataField       =   "OprAge"
         Caption         =   "OprAge"
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
      BeginProperty Column07 
         DataField       =   "OprBirthday"
         Caption         =   "OprBirthday"
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
      BeginProperty Column08 
         DataField       =   "OprPeriod"
         Caption         =   "OprPeriod"
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
      BeginProperty Column09 
         DataField       =   "OprAddr"
         Caption         =   "OprAddr"
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
      BeginProperty Column10 
         DataField       =   "OprTel"
         Caption         =   "OprTel"
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
      BeginProperty Column11 
         DataField       =   "OprPage"
         Caption         =   "OprPage"
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
      BeginProperty Column12 
         DataField       =   "OprPhone"
         Caption         =   "OprPhone"
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
      BeginProperty Column13 
         DataField       =   "OprNative"
         Caption         =   "OprNative"
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
      BeginProperty Column14 
         DataField       =   "OprXL"
         Caption         =   "OprXL"
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
      BeginProperty Column15 
         DataField       =   "OprCollage"
         Caption         =   "OprCollage"
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
      BeginProperty Column16 
         DataField       =   "OprZJ"
         Caption         =   "OprZJ"
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
      BeginProperty Column17 
         DataField       =   "Money"
         Caption         =   "Money"
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
      BeginProperty Column18 
         DataField       =   "AddMoney"
         Caption         =   "AddMoney"
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
      BeginProperty Column19 
         DataField       =   "EditMoney"
         Caption         =   "EditMoney"
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
      BeginProperty Column20 
         DataField       =   "DelMoney"
         Caption         =   "DelMoney"
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
      BeginProperty Column21 
         DataField       =   "Opr"
         Caption         =   "Opr"
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
      BeginProperty Column22 
         DataField       =   "AddOpr"
         Caption         =   "AddOpr"
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
      BeginProperty Column23 
         DataField       =   "EditOpr"
         Caption         =   "EditOpr"
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
      BeginProperty Column24 
         DataField       =   "DelOpr"
         Caption         =   "DelOpr"
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
      BeginProperty Column25 
         DataField       =   "Count"
         Caption         =   "Count"
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
      BeginProperty Column26 
         DataField       =   "AddCount"
         Caption         =   "AddCount"
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
      BeginProperty Column27 
         DataField       =   "EditCount"
         Caption         =   "EditCount"
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
      BeginProperty Column28 
         DataField       =   "DelCount"
         Caption         =   "DelCount"
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
      BeginProperty Column29 
         DataField       =   "Cor"
         Caption         =   "Cor"
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
      BeginProperty Column30 
         DataField       =   "AddCor"
         Caption         =   "AddCor"
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
      BeginProperty Column31 
         DataField       =   "EditCor"
         Caption         =   "EditCor"
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
      BeginProperty Column32 
         DataField       =   "DelCor"
         Caption         =   "DelCor"
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
      BeginProperty Column33 
         DataField       =   "Restore"
         Caption         =   "Restore"
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
      BeginProperty Column34 
         DataField       =   "UsedMoney"
         Caption         =   "UsedMoney"
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
      BeginProperty Column35 
         DataField       =   "Import"
         Caption         =   "Import"
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
      BeginProperty Column36 
         DataField       =   "ImportRz"
         Caption         =   "ImportRz"
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
      BeginProperty Column37 
         DataField       =   "ZZ"
         Caption         =   "ZZ"
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
      BeginProperty Column38 
         DataField       =   "remark"
         Caption         =   "remark"
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
      BeginProperty Column39 
         DataField       =   "TypeName"
         Caption         =   "TypeName"
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
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcOperator 
      Height          =   495
      Left            =   1860
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Connect         =   "DSN=IpCard"
      OLEDBString     =   "DSN=IpCard"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   "iamchinese"
      RecordSource    =   "select * from cxoperator"
      Caption         =   "Operator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub adcOperator_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Load()
writeLog Me.Name
Me.Icon = LoadResPicture(101, 1)
IniMenu 1
'adcoperatorRecordSource = "Select * from Corporation where rtrim(OprNo)=  '" & gOpr.strOprNo & "'"
'adcOperator.Refresh
'RefreshWindow Me.dgOperator.Columns, Me.Name
setADODC Me.adcOperator
'RefreshWindow Me.dgOperator.Columns, "frmOperator"
RefreshOpr Me.dgOperator.Columns
With adcOperator.Recordset
 If Not .EOF Then .MoveLast
End With
End Sub
Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgOperator.Width = Me.Width - 120
dgOperator.Height = Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
writeLog Me.Name, "Exit"
IniMenu 11
End Sub
Private Sub RefreshOpr(colColumns As Columns)
On Error Resume Next
With colColumns
 .Item("OprNo").Width = 6 * 135
 .Item("OprNo").Caption = "编号"
 .Item("OprName").Width = 10 * 130
 .Item("OprName").Caption = "操作员"
 .Item("OprSex").Width = 4 * 130
 .Item("OprSex").Caption = "性别"
 .Item("OprAge").Width = 4 * 130
 .Item("OprAge").Caption = "年龄"
 .Item("OprBirthday").Width = 10 * 130
 .Item("OprBirthday").Caption = "生日"
 .Item("OprPeriod").Width = 10 * 130
 .Item("OprPeriod").Caption = "工作时间"
 .Item("OprAddr").Width = 20 * 130
 .Item("OprAddr").Caption = "家庭住址"
 .Item("OprNative").Width = 4 * 130
 .Item("OprNative").Caption = "民族"
 .Item("OprXL").Width = 4 * 130
 .Item("OprXL").Caption = "学历"
 .Item("OprCollage").Width = 20 * 130
 .Item("OprCollage").Caption = "毕业院校"
 .Item("OprZJ").Width = 20 * 130
 .Item("OprZJ").Caption = "证件号码"
 .Item("OprTel").Width = 10 * 130
 .Item("OprTel").Caption = "手机号码"
 .Item("OprPage").Width = 10 * 130
 .Item("OprPage").Caption = "传呼号码"
 .Item("OprPhone").Width = 10 * 130
 .Item("OprPhone").Caption = "家庭电话"
 .Item("OprPW").Width = 0 ' 10 * 130
 .Item("OprPW").Caption = "操作员"
 .Item("TypeName").Width = 10 * 130
 .Item("TypeName").Caption = "操作员类型"
 .Item("OprType").Width = 0 ' 10 * 130
 .Item("OprRights").Width = 0 ' 10 * 130
 .Item("Money").Width = 0 ' 10 * 130
 .Item("AddMoney").Width = 0 ' 10 * 130
 .Item("EditMoney").Width = 0 ' 10 * 130
 .Item("DelMoney").Width = 0 ' 10 * 130
 .Item("Opr").Width = 0 ' 10 * 130
 .Item("AddOpr").Width = 0 ' 10 * 130
 .Item("EditOpr").Width = 0 ' 10 * 130
 .Item("DelOpr").Width = 0 ' 10 * 130
 .Item("Count").Width = 0 ' 10 * 130
 .Item("AddCount").Width = 0 ' 10 * 130
 .Item("EditCount").Width = 0 ' 10 * 130
 .Item("DelCount").Width = 0 ' 10 * 130
 .Item("Cor").Width = 0 ' 10 * 130
 .Item("AddCor").Width = 0 ' 10 * 130
 .Item("EditCor").Width = 0 ' 10 * 130
 .Item("DelCor").Width = 0 ' 10 * 130
 .Item("Restore").Width = 0 ' 10 * 130
 .Item("UsedMoney").Width = 0 ' 10 * 130
 .Item("Import").Width = 0 ' 10 * 130
 .Item("ImportRZ").Width = 0 ' 10 * 130
 .Item("ZZ").Width = 0 ' 10 * 130
 .Item("Remark").Width = 0 ' 40 * 120
 .Item("Remark").Caption = "备注"
End With
End Sub



