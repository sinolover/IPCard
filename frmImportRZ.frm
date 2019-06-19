VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImportRZ 
   Caption         =   "导入日志"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   8070
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adcImportRZ 
      Height          =   555
      Left            =   900
      Top             =   3420
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   979
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
      RecordSource    =   "select * from cxImportRZ"
      Caption         =   "adcImportRZ"
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
   Begin MSDataGridLib.DataGrid dgImportRZ 
      Bindings        =   "frmImportRZ.frx":0000
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5847
      _Version        =   393216
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
      ColumnCount     =   12
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
         DataField       =   "Importdate"
         Caption         =   "Importdate"
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
         DataField       =   "FileDate"
         Caption         =   "FileDate"
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
         DataField       =   "FilePath"
         Caption         =   "FilePath"
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
         DataField       =   "FileName"
         Caption         =   "FileName"
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
         DataField       =   "RecNum"
         Caption         =   "RecNum"
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
         DataField       =   "TimeTotal"
         Caption         =   "TimeTotal"
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
         DataField       =   "CountTotal"
         Caption         =   "CountTotal"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "UsedName"
         Caption         =   "UsedName"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1454.74
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
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImportRZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub adcImportRZ_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Load()
Me.Icon = LoadResPicture(101, 1)
IniMenu 1
setADODC Me.adcImportRZ
adcImportRZ.RecordSource = "Select * from cxImportRZ " ' order by ID" ' where rtrim(wkrNo)=  '" & gOpr.strOprName & "'"
adcImportRZ.Refresh
RefreshImportRZ Me.dgImportRZ.Columns
With adcImportRZ.Recordset
 If Not .EOF Then .MoveLast
End With
End Sub

Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgImportRZ.Width = frmImportRZ.Width - 120
dgImportRZ.Height = frmImportRZ.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
IniMenu 11
End Sub
Private Sub RefreshImportRZ(colColumns As Columns)
On Error Resume Next
With colColumns
 .Item("ID").Width = IIf(GetReg("ID", "1") = "1", 6 * 135, 0)
 .Item("ID").Caption = "编号"
 .Item("ImportDate").Width = 10 * 130
 .Item("ImportDate").Caption = "导入日期"
 .Item("FileDate").Width = 10 * 130
 .Item("FileDate").Caption = "解析日期"
 .Item("FilePath").Width = 20 * 130
 .Item("FilePath").Caption = "文件路径"
 .Item("FileName").Width = 20 * 135
 .Item("FileName").Caption = "文件名称"
 .Item("OprNo").Width = 0 ' IIf(GetReg("OprNo", "1") = "1", 6 * 130, 0)
 .Item("OprName").Width = 10 * 130
 .Item("OprName").Caption = "操作员"
 .Item("EditName").Width = 8 * 130
 .Item("EditName").Caption = "编辑标志"
 .Item("UsedName").Width = 8 * 130
 .Item("UsedName").Caption = "使用标志"
 .Item("RecNum").Width = 8 * 130
 .Item("RecNum").Caption = "总记录数"
 .Item("TimeTotal").Width = 6 * 130
 .Item("TimeTotal").Caption = "总时长"
 .Item("CountTotal").Width = 8 * 120
 .Item("CountTotal").Caption = "总金额"
End With
End Sub

