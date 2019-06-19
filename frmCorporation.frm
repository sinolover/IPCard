VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCorporation 
   Caption         =   "往来单位信息"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCorporation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgCor 
      Bindings        =   "frmCorporation.frx":1472
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3731
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
      ColumnCount     =   15
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "WkrNo"
         Caption         =   "WkrNo"
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
         DataField       =   "CorAddr"
         Caption         =   "CorAddr"
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
         DataField       =   "CorTel"
         Caption         =   "CorTel"
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
         DataField       =   "CorFax"
         Caption         =   "CorFax"
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
         DataField       =   "CorMan"
         Caption         =   "CorMan"
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
         DataField       =   "CorAge"
         Caption         =   "CorAge"
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
         DataField       =   "CorSex"
         Caption         =   "CorSex"
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
         DataField       =   "CorPhone"
         Caption         =   "CorPhone"
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
         DataField       =   "CorManager"
         Caption         =   "CorManager"
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
         DataField       =   "OprNo"
         Caption         =   "OprNo"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1814.74
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
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcCor 
      Height          =   555
      Left            =   840
      Top             =   2280
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   979
      ConnectMode     =   4
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
      RecordSource    =   "select * from cxcorporation"
      Caption         =   "Corporation"
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
Attribute VB_Name = "frmCorporation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub adcCor_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub dgCor_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
Private Sub Form_Load()
writeLog Me.Name
Me.Icon = LoadResPicture(101, 1)
'adcCor.RecordSource = "Select * from cxCorporation where rtrim(WkrName)=  '" & gCor.strCorName & "'"
'adcCor.RecordSource = "Select * from cxCorporation"
'MsgShow Me.Tag
setADODC Me.adcCor
adcCor.RecordSource = strSql
adcCor.Refresh
RefreshCor Me.dgCor.Columns
'RefreshWindow Me.dgCor.Columns, "frmCorporation"
If adcCor.Recordset.RecordCount > 0 Then adcCor.Recordset.MoveLast
IniMenu 1
End Sub

Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgCor.Width = frmCorporation.Width - 120
dgCor.Height = frmCorporation.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
writeLog Me.Name, "Exit"
IniMenu 11
End Sub
Private Sub RefreshCor(colColumns As Columns)
On Error Resume Next
With colColumns
 .Item("CorNo").Width = 6 * 135
 .Item("CorNo").Caption = "编号"
 .Item("CorName").Width = 10 * 130
 .Item("CorName").Caption = "公司名称"
 .Item("CorAge").Width = 0 ' 4 * 130
 .Item("CorAge").Caption = "年龄"
 .Item("CorMan").Width = 8 * 130
 .Item("CorMan").Caption = "联系人"
 .Item("CorSex").Width = 0 ' 4 * 130
 .Item("CorSex").Caption = "性别"
 .Item("CorManager").Width = 8 * 130
 .Item("CorManager").Caption = "公司经理"
 .Item("CorAddr").Width = 20 * 130
 .Item("CorAddr").Caption = "公司地址"
 .Item("CorTel").Width = 10 * 130
 .Item("CorTel").Caption = "公司电话"
 .Item("CorFax").Width = 10 * 130
 .Item("CorFax").Caption = "公司传真"
 .Item("CorPhone").Width = 10 * 130
 .Item("CorPhone").Caption = "联系人电话"
 .Item("OprNo").Width = 0
 .Item("OprName").Width = 10 * 130
 .Item("OprName").Caption = "操作员"
 .Item("WkrNo").Width = 0
 .Item("WkrName").Width = 10 * 130
 .Item("WkrName").Caption = "业务员"
 .Item("Remark").Width = 40 * 120
 .Item("Remark").Caption = "备注"
End With
End Sub

