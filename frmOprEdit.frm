VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOprEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "操作员明细"
   ClientHeight    =   5520
   ClientLeft      =   3120
   ClientTop       =   2505
   ClientWidth     =   7440
   Icon            =   "frmOprEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtOprSex 
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
      Left            =   720
      TabIndex        =   55
      Top             =   937
      Width           =   870
   End
   Begin VB.ComboBox txtOprNative 
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
      Left            =   720
      TabIndex        =   54
      Top             =   1353
      Width           =   870
   End
   Begin VB.ComboBox txtOprXL 
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
      Left            =   2730
      TabIndex        =   53
      Top             =   1353
      Width           =   1530
   End
   Begin VB.TextBox txtOprZJ 
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
      Left            =   5430
      TabIndex        =   10
      Top             =   1770
      Width           =   1905
   End
   Begin VB.TextBox txtOprCollage 
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
      Left            =   5430
      TabIndex        =   8
      Top             =   1353
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker dtpOprBirthday 
      Height          =   360
      Left            =   2730
      TabIndex        =   4
      Top             =   525
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59768833
      CurrentDate     =   37429
   End
   Begin MSComCtl2.DTPicker dtpOprPeriod 
      Height          =   360
      Left            =   2730
      TabIndex        =   6
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59768833
      CurrentDate     =   37429
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "修改(&C)"
      Height          =   400
      Left            =   6045
      TabIndex        =   22
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Frame fraRights 
      Caption         =   "权限设定"
      Height          =   2415
      Left            =   210
      TabIndex        =   36
      Top             =   3075
      Width           =   5535
      Begin VB.CheckBox chkZZ 
         Caption         =   "内部转账"
         Height          =   240
         Left            =   2835
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkCor 
         Caption         =   "往来单位"
         Height          =   255
         Left            =   4305
         TabIndex        =   16
         Top             =   270
         Width           =   1065
      End
      Begin VB.CheckBox chkOpr 
         Caption         =   "操 作 员"
         Height          =   255
         Left            =   2940
         TabIndex        =   15
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkMoney 
         Caption         =   "充值管理"
         Height          =   255
         Left            =   1545
         TabIndex        =   14
         Top             =   270
         Width           =   1080
      End
      Begin VB.CheckBox chkCount 
         Caption         =   "账号维护"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkCountAdd 
         Caption         =   "增加"
         Height          =   255
         Left            =   420
         TabIndex        =   48
         Top             =   615
         Width           =   675
      End
      Begin VB.CheckBox chkCountEdit 
         Caption         =   "修改"
         Height          =   255
         Left            =   420
         TabIndex        =   47
         Top             =   1035
         Width           =   675
      End
      Begin VB.CheckBox chkCountDel 
         Caption         =   "删除"
         Height          =   255
         Left            =   420
         TabIndex        =   46
         Top             =   1425
         Width           =   675
      End
      Begin VB.CheckBox chkMoneyAdd 
         Caption         =   "增加"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   615
         Width           =   675
      End
      Begin VB.CheckBox chkMoneyEdit 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   44
         Top             =   1035
         Width           =   675
      End
      Begin VB.CheckBox chkMoneyDel 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   1425
         Width           =   675
      End
      Begin VB.CheckBox chkOprAdd 
         Caption         =   "增加"
         Height          =   255
         Left            =   3180
         TabIndex        =   42
         Top             =   615
         Width           =   675
      End
      Begin VB.CheckBox chkOprEdit 
         Caption         =   "修改"
         Height          =   255
         Left            =   3180
         TabIndex        =   41
         Top             =   1035
         Width           =   675
      End
      Begin VB.CheckBox chkOprDel 
         Caption         =   "删除"
         Height          =   255
         Left            =   3180
         TabIndex        =   40
         Top             =   1425
         Width           =   675
      End
      Begin VB.CheckBox chkCorAdd 
         Caption         =   "增加"
         Height          =   255
         Left            =   4620
         TabIndex        =   39
         Top             =   615
         Width           =   675
      End
      Begin VB.CheckBox chkCorEdit 
         Caption         =   "修改"
         Height          =   255
         Left            =   4620
         TabIndex        =   38
         Top             =   1035
         Width           =   675
      End
      Begin VB.CheckBox chkCorDel 
         Caption         =   "删除"
         Height          =   255
         Left            =   4635
         TabIndex        =   37
         Top             =   1425
         Width           =   675
      End
      Begin VB.CheckBox chkUsedMoney 
         Caption         =   "费用查询"
         Height          =   315
         Left            =   4290
         TabIndex        =   20
         Top             =   2025
         Width           =   1095
      End
      Begin VB.CheckBox chkRestore 
         Caption         =   "数据恢复"
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   1980
         Width           =   1020
      End
      Begin VB.CheckBox chkImport 
         Caption         =   "数据导入"
         Height          =   315
         Left            =   1650
         TabIndex        =   18
         Top             =   2010
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   15
         X2              =   5520
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line2 
         X1              =   15
         X2              =   5520
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "类型"
      Height          =   750
      Left            =   240
      TabIndex        =   35
      Top             =   2250
      Width           =   5460
      Begin VB.OptionButton optWorker 
         Caption         =   "业务员"
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   285
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton optManager 
         Caption         =   "管理员"
         Height          =   270
         Left            =   3165
         TabIndex        =   12
         Top             =   300
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc adcOperator 
      Height          =   495
      Left            =   3525
      Top             =   2385
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
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
      Password        =   ""
      RecordSource    =   "select * from Operator"
      Caption         =   "adcOperator"
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
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(Q)"
      Height          =   400
      Left            =   6060
      TabIndex        =   24
      Top             =   4620
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(D)"
      Height          =   400
      Left            =   6060
      TabIndex        =   23
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "增加(&A)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6060
      TabIndex        =   21
      Top             =   2445
      Width           =   1200
   End
   Begin VB.TextBox txtOprAddr 
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
      Left            =   720
      TabIndex        =   9
      Top             =   1770
      Width           =   3540
   End
   Begin VB.TextBox txtOprPage 
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
      Left            =   5430
      TabIndex        =   7
      Top             =   937
      Width           =   1905
   End
   Begin VB.TextBox txtOprTel 
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
      Left            =   5430
      TabIndex        =   5
      Top             =   521
      Width           =   1905
   End
   Begin VB.TextBox txtOprAge 
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
      Left            =   720
      TabIndex        =   3
      Top             =   521
      Width           =   870
   End
   Begin VB.TextBox txtOprPhone 
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
      Left            =   5430
      TabIndex        =   2
      Top             =   105
      Width           =   1905
   End
   Begin VB.TextBox txtOprName 
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
      Left            =   2730
      TabIndex        =   1
      Top             =   105
      Width           =   1530
   End
   Begin VB.TextBox txtOprNo 
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
      Left            =   720
      TabIndex        =   0
      Top             =   105
      Width           =   870
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "证件号码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   52
      Top             =   1845
      Width           =   840
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "毕业院校"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   51
      Top             =   1425
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "最高学历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1725
      TabIndex        =   50
      Top             =   1428
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "民族"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   49
      Top             =   1428
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "手机号码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   34
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "呼机号码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   33
      Top             =   1005
      Width           =   840
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "联系电话"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   32
      Top             =   180
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "住址"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   31
      Top             =   1845
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "工作时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1725
      TabIndex        =   30
      Top             =   1005
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1725
      TabIndex        =   29
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   28
      Top             =   596
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   27
      Top             =   1012
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " 姓  名 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1725
      TabIndex        =   26
      Top             =   180
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   25
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmOprEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub adcOperator_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub cmdAppend_Click()
If Len(Trim(txtOprName)) < 1 Then
   MsgShow "请输入业务员姓名！"
   txtOprName.SetFocus
   GoTo Error_Exit
End If
If Not IsNumeric(txtOprAge.Text) Then
   MsgShow "年龄必须为数字！"
   txtOprAge.SetFocus
   GoTo Error_Exit
End If
If funOprNo() Then Exit Sub
With frmOperator.adcOperator.Recordset
 .AddNew
 !oprNo = Trim(txtOprNo)
 !OprName = Trim(txtOprName)
 If optManager.Value Then
    !OprType = 1
   Else
    !OprType = 0
 End If
 !OprPW = "888888"
 !OprSex = IIf(IsNull(txtOprSex.Text), "男", Trim(txtOprSex))
 !OprAge = IIf(IsNull(txtOprAge.Text), 23, CInt(txtOprAge.Text))
 !OprBirthday = dtpOprBirthday.Value
 !OprPeriod = dtpOprPeriod.Value
 !OprAddr = IIf(IsNull(txtOprAddr.Text), "未知", Trim(txtOprAddr))
 !OprTel = IIf(IsNull(txtOprTel.Text), "未知", Trim(txtOprTel.Text))
 !OprPage = IIf(IsNull(txtOprPage.Text), "未知", Trim(txtOprPage.Text))
 !OprPhone = IIf(IsNull(txtOprPhone.Text), "未知", Trim(txtOprPhone.Text))
 !OprNative = txtOprNative.Text
 !OprXL = txtOprXL.Text
 !OprCollage = txtOprCollage.Text
 !OprZJ = txtOprZJ.Text
 !Money = chkMoney.Value
 !AddCount = chkCountAdd.Value
 !EditCount = chkCountEdit.Value
 !DelCount = chkCountDel.Value
 !AddMoney = chkMoneyAdd.Value
 !EditMoney = chkMoneyEdit.Value
 !DelMoney = chkMoneyDel.Value
 !Opr = chkOpr.Value
 !AddOpr = chkOprAdd.Value ' IIf(chkOprAdd.check, 1, 0)
 !EditOpr = chkOprEdit.Value
 !DelOpr = chkOprDel.Value
 !Cor = chkCor.Value
 !AddCor = chkCorAdd.Value
 !EditCor = chkCorEdit.Value
 !DelCor = chkCorDel.Value
 !Count = chkCount.Value
 !Restore = chkRestore.Value
 !Import = chkImport.Value
 !ZZ = chkZZ.Value
 !UsedMoney = chkUsedMoney.Value
 '!RZ=chkrz.value
 .Update
 writeLog Me.Name, "Append", !OprName
 End With
' MsgShow frmOperator.adcOperator.RecordSource
'With frmOperator.adcOperator
' .Refresh
' If Not .Recordset.EOF Then .Recordset.MoveLast
'End With
Unload Me
Error_Exit:
' txtOprNo.SetFocus
End Sub

Private Sub cmdChange_Click()
If Trim(txtOprNo.Text) = "1001" Then
   MsgShow "对不起，您不可以修改机器管理员！"
   GoTo Error_Exit
End If
If Len(Trim(txtOprName)) < 1 Then
   MsgShow "请输入业务员姓名！"
   txtOprName.SetFocus
   GoTo Error_Exit
End If
If Not IsNumeric(txtOprAge.Text) Then
   MsgShow "年龄必须位数字数字！"
   txtOprAge.SetFocus
   GoTo Error_Exit
End If
With frmOperator.adcOperator.Recordset
If Not .EOF And Not .BOF Then
 .Fields("OprName") = Trim(txtOprName)
 If optManager.Value Then
    !OprType = 1
   Else
    !OprType = 0
 End If
 .Fields("OprPW") = "888888"
 .Fields("OprSex") = IIf(IsNull(txtOprSex.Text), "男", Trim(txtOprSex))
 .Fields("OprAge") = IIf(IsNull(txtOprAge.Text), 23, CInt(txtOprAge.Text))
 .Fields("OprBirthday") = dtpOprBirthday.Value
 .Fields("OprPeriod") = dtpOprPeriod.Value
 .Fields("OprAddr") = IIf(IsNull(txtOprAddr.Text), "未知", Trim(txtOprAddr))
 .Fields("OprTel") = IIf(IsNull(txtOprTel.Text), "未知", Trim(txtOprTel.Text))
 .Fields("OprPage") = IIf(IsNull(txtOprPage.Text), "未知", Trim(txtOprPage.Text))
 .Fields("OprPhone") = IIf(IsNull(txtOprPhone.Text), "未知", Trim(txtOprPhone.Text))
 !OprNative = txtOprNative.Text
 !OprXL = txtOprXL.Text
 !OprCollage = txtOprCollage.Text
 !OprZJ = txtOprZJ.Text
 !Money = chkMoney.Value
 !AddCount = chkCountAdd.Value
 !EditCount = chkCountEdit.Value
 !DelCount = chkCountDel.Value
 !AddMoney = chkMoneyAdd.Value
 !EditMoney = chkMoneyEdit.Value
 !DelMoney = chkMoneyDel.Value
 !Opr = chkOpr.Value
 !AddOpr = chkOprAdd.Value ' IIf(chkOprAdd.check, 1, 0)
 !EditOpr = chkOprEdit.Value
 !DelOpr = chkOprDel.Value
 !Cor = chkCor.Value
 !AddCor = chkCorAdd.Value
 !EditCor = chkCorEdit.Value
 !DelCor = chkCorDel.Value
 !Count = chkCount.Value
 !Restore = chkRestore.Value
 !Import = chkImport.Value
 !ZZ = chkZZ.Value
 !UsedMoney = chkUsedMoney.Value
 '!RZ=chkrz.value
 .Update
 writeLog Me.Name, "Change", !OprName
Else
 MsgShow "清先选定待修改记录！"
End If
 End With
Error_Exit:
'Dim strID As String
'With frmOperator.adcOperator
' strID = Str(.Recordset!OprNo)
' .Refresh
' .Refresh
' .Recordset.Find "OprNO=" & strID, , adSearchForward, 1
' If Not .Recordset.EOF Then .Recordset.MoveLast
'End With
 Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim cmdIPCount As New ADODB.Command
Dim lngrow As Long
On Error Resume Next
lngrow = 1
With frmOperator
If IsNumeric(.dgOperator.Row) Then lngrow = .dgOperator.Row
If blnRegstered Then
   cmdIPCount.ActiveConnection = gCnn
   cmdIPCount.CommandText = "delete from Operator where rtrim(oprno)='" & Trim(txtOprNo.Text) & "'"
   cmdIPCount.Execute
   '.Delete adAffectCurrent
   .adcOperator.Refresh
   .adcOperator.Recordset.Move lngrow
 Else
   MsgShow "这是一个试用版本，您将不能够修改账号信息！"
End If
End With
'With frmOperator.adcOperator
' .Refresh
' If Not .Recordset.EOF Then .Recordset.MoveLast
'End With
Unload Me

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
'txtOprNo.SetFocus
'Me.Icon = LoadResPicture(103, 1)
setADODC Me.adcOperator
With txtOprSex
 .AddItem "男"
 .AddItem "女"
 .Text = "男"
End With
With txtOprNative
 .AddItem "汉"
 .AddItem "满"
 .AddItem "彝"
 .AddItem "其他"
 .Text = "汉"
End With
With txtOprXL
 .AddItem "小学"
 .AddItem "初中"
 .AddItem "高中"
 .AddItem "大学"
 .AddItem "学士"
 .AddItem "硕士"
 .AddItem "博士"
 .AddItem "其他"
 .Text = "大学"
End With
With frmOperator.adcOperator.Recordset
If Not .EOF And Not .BOF Then
   txtOprNo.Text = !oprNo
   txtOprName = .Fields("OprName") & ""
   If !OprType = 1 Then
      optManager.Value = True
     Else
      optWorker.Value = True
   End If
  '.Fields("OprPW") = "888888"
   txtOprSex.Text = .Fields("OprSex") & ""
   txtOprAge.Text = .Fields("OprAge")
   dtpOprBirthday.Value = IIf(IsDate(!OprBirthday), !OprBirthday, Date)
   dtpOprPeriod.Value = IIf(IsDate(!OprPeriod), !OprPeriod, Date)
   txtOprAddr.Text = !OprAddr & ""
   txtOprTel.Text = !OprTel & ""
   txtOprPage.Text = !OprPage & ""
   txtOprPhone.Text = !OprPhone & ""
   txtOprXL.Text = !OprXL & ""
   txtOprNative.Text = !OprNative & ""
   txtOprCollage.Text = !OprCollage & ""
   txtOprZJ.Text = !OprZJ & ""
   chkMoney.Value = !Money
   chkCountAdd.Value = !AddCount
   chkCountEdit.Value = !EditCount
   chkCountDel.Value = !DelCount
   chkMoneyAdd.Value = !AddMoney
   chkMoneyEdit.Value = !EditMoney
   chkMoneyDel.Value = !DelMoney
   chkOpr.Value = !Opr
   chkOprAdd.Value = !AddOpr
   chkOprEdit.Value = !EditOpr
   chkOprDel.Value = !DelOpr
   chkCor.Value = !Cor
   chkCorAdd.Value = !AddCor
   chkCorEdit.Value = !EditCor
   chkCorDel.Value = !DelCor
   chkCount.Value = !Count
   chkRestore.Value = !Restore
   chkImport.Value = !Import
   chkZZ.Value = !ZZ
   chkUsedMoney.Value = !UsedMoney
 '!RZ=chkrz.value
 End If
 End With
End Sub
'Private Sub txtOprNo_LostFocus()
'If Len(txtOprNo) < 1 Then
' MsgShow "操作员编号错！"
' txtOprNo.SetFocus
' Exit Sub
'End If
'adcOperator.RecordSource = "select * from operator where OprNo= '" & Trim(txtOprNo) & "'"
'adcOperator.Refresh
'If adcOperator.Recordset.RecordCount > 0 Then
' MsgShow "该编号已经存在！"
' txtOprNo.SetFocus
' Exit Sub
'End If
'End Sub
Private Function funOprNo() As Boolean
If Len(txtOprNo) < 1 Then
 MsgShow "操作员编号错！"
 txtOprNo.SetFocus
 funOprNo = True
 Exit Function
End If
adcOperator.RecordSource = "select * from operator where OprNo= '" & Trim(txtOprNo) & "'"
adcOperator.Refresh
If adcOperator.Recordset.RecordCount > 0 Then
 MsgShow "该编号已经存在！"
 txtOprNo.SetFocus
 funOprNo = True
 Exit Function
End If
funOprNo = False
End Function
Public Sub SetStatus(strStatus As String)
Select Case Trim(strStatus)
       Case "Append"
'        txtCorNo.Text = ""
        cmdAppend.Enabled = Enabled ' IIf(Trim(gOpr.chrAddOpr) = "1", True, False)
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
        Me.Caption = "增加新的业务员信息"
       Case "Change"
        txtOprNo.Enabled = False
        cmdAppend.Enabled = False
        cmdChange.Enabled = Enabled ' IIf(Trim(gOpr.chrEditWkr) = "1", True, False)
        cmdDelete.Enabled = False
        Me.Caption = "对当前的业务员信息进行修改"
       Case "Delete"
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = Enabled ' IIf(Trim(gOpr.chrDelWkr) = "1", True, False)
        Me.Caption = "删除当前的业务员信息"
       Case Else
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
End Select
End Sub
