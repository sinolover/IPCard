VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCorDaily 
   Caption         =   "公司日报表"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adcUsedMoney 
      Height          =   675
      Left            =   360
      Top             =   2235
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1191
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
      RecordSource    =   "select * from cxCorDaily"
      Caption         =   "UsedMoney"
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
   Begin MSDataGridLib.DataGrid dgUsedMoney 
      Bindings        =   "frmCorDaily.frx":0000
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   3201
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "useddate"
         Caption         =   "useddate"
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
         DataField       =   "sumusedMoney"
         Caption         =   "sumusedMoney"
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
         DataField       =   "sumusedtime"
         Caption         =   "sumusedtime"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCorDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adcUsedMoney_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Load()
writeLog Me.Name
Me.Icon = LoadResPicture(101, 1)
IniMenu 1
setADODC Me.adcUsedMoney
RefreshUsedMoney Me.dgUsedMoney.Columns
adcUsedMoney.RecordSource = strSql ' "Select * from cxUsedMoney where rtrim(WkrName)=  '" & gOpr.strOprName & "'"
adcUsedMoney.Refresh
If adcUsedMoney.Recordset.RecordCount > 0 Then adcUsedMoney.Recordset.MoveLast
End Sub

Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgUsedMoney.Width = Me.Width - 120
dgUsedMoney.Height = Me.Height - 400

End Sub

Private Sub Form_Unload(Cancel As Integer)
writeLog Me.Name, "Exit"
IniMenu 11

End Sub
Private Sub RefreshUsedMoney(colColumns As Columns)
With colColumns
 .Item("UsedDate").Width = 18 * 130
 .Item("UsedDate").Caption = "使用日期"
 .Item("sumUsedMoney").Width = 16 * 135
 .Item("sumUsedMoney").Alignment = dbgRight
 .Item("sumUsedMoney").NumberFormat = "Fixed"
 .Item("sumUsedMoney").Caption = "使用金额"
' .Item("NowMoney").Width = 8 * 135
' .Item("NowMoney").Caption = "当前余额"
' .Item("AlertMoney").Width = 8 * 135
' .Item("AlertMoney").Caption = "警戒金额"
' .Item("UsedName").Width = 8 * 135
' .Item("UsedName").Caption = "费用标记"
 .Item("sumUsedTime").Width = 16 * 135
 .Item("sumUsedTime").Caption = "使用时长"
' .Item("WithTel").Width = 10 * 130
' .Item("WithTel").Caption = "邦定电话"
' .Item("OprName").Width = IIf(GetReg("OprName", "1") = "1", 10 * 130, 0)
' .Item("OprName").Caption = "操作员"
' .Item("WkrName").Width = 10 * 130
' .Item("WkrName").Caption = "业务员"
 .Item("CorName").Width = 40 * 130
 .Item("CorName").Caption = "业务单位"
' .Item("LastDate").Width = 10 * 130
' .Item("LastDate").Caption = "使用日期"
' .Item("Remark").Width = 40 * 120
' .Item("Remark").Caption = "备注"
End With
End Sub



