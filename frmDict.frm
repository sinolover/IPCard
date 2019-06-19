VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDict 
   Caption         =   "字典设置"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "frmDict.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   7680
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgDict 
      Bindings        =   "frmDict.frx":27A2
      Height          =   1320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2328
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
      ColumnCount     =   23
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
         DataField       =   "Field_Name"
         Caption         =   "Field_Name"
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
         DataField       =   "Field_Type"
         Caption         =   "Field_Type"
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
         DataField       =   "Field_Len"
         Caption         =   "Field_Len"
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
         DataField       =   "Field_Dec"
         Caption         =   "Field_Dec"
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
         DataField       =   "Disp_YN"
         Caption         =   "Disp_YN"
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
         DataField       =   "Disp_Name"
         Caption         =   "Disp_Name"
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
         DataField       =   "Disp_Type"
         Caption         =   "Disp_Type"
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
         DataField       =   "Disp_Len"
         Caption         =   "Disp_Len"
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
         DataField       =   "Disp_Dec"
         Caption         =   "Disp_Dec"
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
         DataField       =   "Disp_Align"
         Caption         =   "Disp_Align"
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
         DataField       =   "Disp_WrapText"
         Caption         =   "Disp_WrapText"
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
         DataField       =   "Disp_Index"
         Caption         =   "Disp_Index"
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
         DataField       =   "Print_YN"
         Caption         =   "Print_YN"
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
         DataField       =   "Print_Name"
         Caption         =   "Print_Name"
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
         DataField       =   "Print_Type"
         Caption         =   "Print_Type"
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
         DataField       =   "Print_Len"
         Caption         =   "Print_Len"
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
         DataField       =   "Print_Dec"
         Caption         =   "Print_Dec"
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
         DataField       =   "Print_Index"
         Caption         =   "Print_Index"
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
         DataField       =   "Print_Font_Name"
         Caption         =   "Print_Font_Name"
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
         DataField       =   "Print_Fone_Size"
         Caption         =   "Print_Fone_Size"
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
         DataField       =   "Module_Name"
         Caption         =   "Module_Name"
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
         DataField       =   "Key"
         Caption         =   "Key"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcDict 
      Height          =   615
      Left            =   690
      Top             =   2175
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1085
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
      RecordSource    =   "select * from sysDict"
      Caption         =   "Adodc1"
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
Attribute VB_Name = "frmDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adcDict_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub dgDict_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
MsgShow "无法编辑！"
Cancel = True
End Sub
Private Sub Form_Load()
RefreshDict Me.dgDict.Columns
setADODC adcDict
Me.adcDict.Refresh
'CenterForm Me
End Sub

Private Sub Form_Resize()
If Me.Width < 100 Then Exit Sub
If Me.Height < 500 Then Exit Sub
dgDict.Width = Me.Width - 120
dgDict.Height = Me.Height - 400
End Sub
Private Sub RefreshDict(colColumns As Columns)
On Error Resume Next
With colColumns
 .Item("ID").Width = IIf(GetReg("ID", "1") = "1", 6 * 135, 0)
 .Item("ID").Caption = "编号"
 .Item("Field_Name").Width = 0 ' 10 * 130
' .Item("AddDate").NumberFormat = "Long Date"
 .Item("Field_Name").Caption = "字段名称"
 .Item("Field_type").Width = 0 ' 10 * 130
 .Item("Field_type").Caption = "字段类型"
 .Item("Field_len").Width = 0 ' 10 * 130
 .Item("Field_len").Caption = "字段长度"
 .Item("Field_dec").Width = 0 ' 8 * 135
 .Item("Field_dec").Caption = "小数位"
 .Item("Disp_YN").Width = 8 * 135
 .Item("Disp_YN").Caption = "是否显示"
 .Item("Disp_name").Width = 8 * 135
 .Item("Disp_name").Caption = "显示名称"
 .Item("Disp_type").Width = 8 * 135
 .Item("Disp_type").Caption = "显示类型"
 .Item("Disp_len").Width = 8 * 135
 .Item("Disp_len").Caption = "显示宽度"
 .Item("Disp_dec").Width = 8 * 135
 .Item("Disp_dec").Caption = "小数位"
 .Item("Disp_align").Width = 8 * 135
 .Item("Disp_align").Caption = "对齐方式"
 .Item("Disp_index").Width = 8 * 135
 .Item("Disp_index").Caption = "显示顺序"
 .Item("Disp_wraptext").Width = 10 * 130
' .Item("Disp_index").Caption = "是否排序"
' .Item("Disp_wraptext").Width = 10 * 130
 .Item("Disp_wraptext").Caption = "是否折行"
 .Item("Print_YN").Width = 8 * 135
 .Item("Print_YN").Caption = "是否打印"
 .Item("Print_name").Width = 8 * 135
 .Item("Print_name").Caption = "打印名称"
 .Item("Print_type").Width = 8 * 135
 .Item("Print_type").Caption = "打印类型"
 .Item("Print_len").Width = 8 * 135
 .Item("Print_len").Caption = "打印宽度"
 .Item("Print_dec").Width = 8 * 135
 .Item("Print_dec").Caption = "小数位"
' .Item("Print_align").Width = 8 * 135
 '.Item("Print_align").Caption = "对齐"
 .Item("Print_index").Width = 8 * 135
 .Item("Print_index").Caption = "打印顺序"
' .Item("Print_wraptext").Width = 10 * 130
' .Item("Print_wraptext").Caption = "折行"
 .Item("key").Width = 0 ' 10 * 130
 .Item("key").Caption = "关键字"
' .Item("Remark").Width = 40 * 120
' .Item("Remark").Caption = "备注"
End With
End Sub

