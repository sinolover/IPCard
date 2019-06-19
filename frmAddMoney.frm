VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddMoney 
   Caption         =   "充值信息"
   ClientHeight    =   4995
   ClientLeft      =   1875
   ClientTop       =   2145
   ClientWidth     =   7485
   Icon            =   "frmAddMoney.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7485
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adcAddMoney 
      Height          =   495
      Left            =   660
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "DSN=IPCard"
      OLEDBString     =   "DSN=IPCard"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   "iamchinese"
      RecordSource    =   "select * from cxAddMoney "
      Caption         =   "AddMoney"
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
   Begin MSDataGridLib.DataGrid dgAddMoney 
      Bindings        =   "frmAddMoney.frx":29A2
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "AddDate"
         Caption         =   "AddDate"
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
         DataField       =   "CountNo"
         Caption         =   "CountNo"
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
      BeginProperty Column04 
         DataField       =   "NowMoney"
         Caption         =   "NowMoney"
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
         DataField       =   "WithTel"
         Caption         =   "WithTel"
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
         DataField       =   "EditName"
         Caption         =   "EditName"
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
      BeginProperty Column08 
         DataField       =   "WkrNO"
         Caption         =   "WkrNO"
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
         DataField       =   "WkrName"
         Caption         =   "WkrName"
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
         DataField       =   "CorNo"
         Caption         =   "CorNo"
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
         DataField       =   "CorName"
         Caption         =   "CorName"
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
         DataField       =   "FromCount"
         Caption         =   "FromCount"
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
         DataField       =   "ReMark"
         Caption         =   "ReMark"
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
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub adcAddMoney_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Load()
'Me.Icon = LoadResPicture(101, 1)
IniMenu 1
adcAddMoney.RecordSource = strSql
setADODC Me.adcAddMoney
RefreshAddMoney Me.dgAddMoney.Columns
Me.adcAddMoney.Refresh
If Not adcAddMoney.Recordset.EOF Then adcAddMoney.Recordset.MoveLast
writeLog Me.Name
End Sub
Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgAddMoney.Width = Me.Width - 120
dgAddMoney.Height = Me.Height - 400
End Sub
Private Sub Form_Unload(Cancel As Integer)
writeLog Me.Name, "Exit"
IniMenu 11
End Sub
Private Sub RefreshAddMoney(colColumns As Columns)
With colColumns
 .Item("ID").Width = IIf(GetReg("ID", "1") = "1", 6 * 135, 0)
 .Item("ID").Caption = "编号"
 .Item("AddDate").Width = 10 * 130
 .Item("AddDate").Caption = "充值日期"
 .Item("CountNo").Width = 10 * 130
 .Item("CountNo").Caption = "账号"
 .Item("FromCount").Width = 10 * 130
 .Item("FromCount").Caption = "转出卡账号"
 .Item("AddMoney").Width = 8 * 135
 .Item("AddMoney").Alignment = dbgRight
 .Item("AddMoney").NumberFormat = "Fixed"
 .Item("AddMoney").Caption = "充值总额"
' .Item("UsedMoney").Width = 8 * 135
' .Item("UsedMoney").Caption = "使用金额"
 .Item("NowMoney").Width = 8 * 135
 .Item("NowMoney").Alignment = dbgRight
 .Item("NowMoney").NumberFormat = "Fixed"
 .Item("NowMoney").Caption = "当前余额"
 '.Item("AlertMoney").Width = 8 * 135
 '.Item("AlertMoney").Caption = "警戒金额"
 .Item("EditName").Width = 8 * 135
 .Item("EditName").Caption = "编辑标记"
 .Item("WithTel").Width = 10 * 130
 .Item("WithTel").Caption = "邦定电话"
' .Item("OprNo").Width = 0
 .Item("OprName").Width = IIf(GetReg("OprName", "1") = "1", 10 * 130, 0)
 .Item("OprName").Caption = "操作员"
 .Item("WkrNo").Width = 0
 .Item("WkrName").Width = 10 * 130
 .Item("WkrName").Caption = "业务员"
 .Item("CorNo").Width = 0
 .Item("CorName").Width = 20 * 130
 .Item("CorName").Caption = "业务单位"
' .Item("LastDate").Width = 10 * 130
' .Item("LastDate").Caption = "使用日期"
 .Item("Remark").Width = 40 * 120
 .Item("Remark").Caption = "备注"
End With
End Sub

