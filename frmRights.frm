VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRights 
   Caption         =   "操作员权限明细表"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmRights.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   6390
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adccxRight 
      Height          =   495
      Left            =   720
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "Select * from cxRight"
      Caption         =   "adccxRight"
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
   Begin MSDataGridLib.DataGrid dgRight 
      Bindings        =   "frmRights.frx":0442
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   4128
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
      ColumnCount     =   17
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
         DataField       =   "ImportRZ"
         Caption         =   "ImportRZ"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1335.118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adccxRight_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True

End Sub

Private Sub Form_Load()
writeLog Me.Name
Me.Icon = LoadResPicture(101, 1)
IniMenu 1
adccxRight.RecordSource = "select * from cxRight"
setADODC Me.adccxRight
'adccxRight.Refresh
RefreshOpr Me.dgRight.Columns
With adccxRight.Recordset
 If Not .EOF Then .MoveLast
End With
End Sub
Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgRight.Width = frmRights.Width - 120
dgRight.Height = frmRights.Height - 400
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
 '.Item("OprSex").Width = 0 ' 4 * 130
 '.Item("OprSex").Caption = "性别"
 '.Item("OprAge").Width = 0 ' 4 * 130
 '.Item("OprAge").Caption = "年龄"
 '.Item("OprBirthday").Width = 0 ' 10 * 130
 '.Item("OprBirthday").Caption = "生日"
 '.Item("OprPeriod").Width = 0 '10 * 130
 '.Item("OprPeriod").Caption = "工作时间"
 '.Item("OprAddr").Width = 0 ' 20 * 130
 '.Item("OprAddr").Caption = "家庭住址"
 '.Item("OprNative").Width = 0 ' 4 * 130
 '.Item("OprNative").Caption = "民族"
 '.Item("OprXL").Width = 0 '4 * 130
 '.Item("OprXL").Caption = "学历"
 '.Item("OprCollage").Width = 0 '20 * 130
' .Item("OprCollage").Caption = "毕业院校"
 '.Item("OprZJ").Width = 0 '20 * 130
 '.Item("OprZJ").Caption = "证件号码"
 '.Item("OprTel").Width = 0 '10 * 130
 '.Item("OprTel").Caption = "手机号码"
 '.Item("OprPage").Width = 0 '10 * 130
 '.Item("OprPage").Caption = "传呼号码"
 '.Item("OprPhone").Width = 0 '10 * 130
 '.Item("OprPhone").Caption = "家庭电话"
 '.Item("OprPW").Width = 0 ' 10 * 130
 '.Item("OprPW").Caption = "操作员"
 .Item("TypeName").Width = 10 * 130
 .Item("TypeName").Caption = "操作员类型"
 '.Item("OprType").Width = 0 ' 10 * 130
 '.Item("OprRights").Width = 0 ' 10 * 130
 '.Item("Money").Width = 8 * 130
 '.Item("Money").Caption = "充值浏览"
 .Item("AddMoney").Width = 8 * 130
 .Item("AddMoney").Caption = "充值增加"
 '.Item("EditMoney").Width = 8 * 130
 '.Item("EditMoney").Caption = "充值修改"
 '.Item("DelMoney").Width = 8 * 130
 '.Item("DelMoney").Caption = "充值删除"
 .Item("Opr").Width = 10 * 130
 .Item("Opr").Caption = "操作员浏览"
 .Item("AddOpr").Width = 10 * 130
 .Item("AddOpr").Caption = "操作员增加"
 .Item("EditOpr").Width = 10 * 130
 .Item("EditOpr").Caption = "操作员修改"
 .Item("DelOpr").Width = 10 * 130
 .Item("DelOpr").Caption = "操作员删除"
 '.Item("Count").Width = 10 * 130
 '.Item("Count").Caption = "帐号浏览"
 .Item("AddCount").Width = 8 * 130
 .Item("AddCount").Caption = "帐号增加"
 .Item("EditCount").Width = 8 * 130
 .Item("EditCount").Caption = "帐号修改"
 .Item("DelCount").Width = 8 * 130
 .Item("DelCount").Caption = "帐号删除"
 '.Item("Cor").Width = 10 * 130
 '.Item("Cor").Caption = "往来单位浏览"
 .Item("AddCor").Width = 10 * 130
 .Item("AddCor").Caption = "往来单位增加"
 .Item("EditCor").Width = 10 * 130
 .Item("EditCor").Caption = "往来单位修改"
 .Item("DelCor").Width = 10 * 130
 .Item("DelCor").Caption = "往来单位删除"
 '.Item("Restore").Width = 8 * 130
 '.Item("Restore").Caption = "数据恢复"
 .Item("UsedMoney").Width = 8 * 130
 .Item("UsedMoney").Caption = "费用浏览"
 '.Item("Import").Width = 8 * 130
 '.Item("Import").Caption = "导入数据"
 .Item("ImportRZ").Width = 10 * 130
 .Item("ImportRZ").Caption = "导入日志"
 .Item("ZZ").Width = 8 * 130
 .Item("ZZ").Caption = "内部转帐"
' .Item("Remark").Width = 20 * 120
' .Item("Remark").Caption = "备注"
End With
End Sub

