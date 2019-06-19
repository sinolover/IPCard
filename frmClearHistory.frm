VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClearHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "清除历史数据记录"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmClearHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   400
      Left            =   3015
      TabIndex        =   3
      Top             =   825
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   400
      Left            =   3015
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "范围"
      Height          =   1215
      Left            =   45
      TabIndex        =   1
      Top             =   1410
      Width           =   4020
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1545
         TabIndex        =   9
         Top             =   750
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   59375617
         CurrentDate     =   37582
      End
      Begin VB.OptionButton optBefore 
         Caption         =   "此时间以前"
         Height          =   390
         Left            =   195
         TabIndex        =   8
         Top             =   735
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optAllData 
         Caption         =   "全部历史数据"
         Height          =   495
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据表"
      Height          =   1290
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2850
      Begin VB.OptionButton optUsedMoney 
         Caption         =   "明细历史数据"
         Height          =   285
         Left            =   255
         TabIndex        =   6
         Top             =   915
         Width           =   1395
      End
      Begin VB.OptionButton optAddMoney 
         Caption         =   "充值历史数据"
         Height          =   285
         Left            =   255
         TabIndex        =   5
         Top             =   600
         Width           =   1395
      End
      Begin VB.OptionButton optIPCount 
         Caption         =   "帐务历史数据"
         Height          =   285
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "0 条"
         Height          =   210
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "0 条"
         Height          =   210
         Left            =   1680
         TabIndex        =   11
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "0 条"
         Height          =   210
         Left            =   1695
         TabIndex        =   10
         Top             =   345
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmClearHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim cmdClearHistory As New ADODB.Command
Dim strSql
strSql = ""
On Error GoTo ErrorExit
If optIPCount.Value Then
   strSql = "delete from IPCountHistory"
   If optBefore Then strSql = strSql & " where AddDate < #" & Format(DTPicker1.Value, "YYYY-MM-DD") & "#"
End If
If optAddMoney.Value Then
   strSql = "delete from AddMoneyHistory"
   If optBefore Then strSql = strSql & " where AddDate < #" & Format(DTPicker1.Value, "YYYY-MM-DD") & "#"
End If
If optUsedMoney.Value Then
   strSql = "delete from UsedMoneyHistory"
   If optBefore Then strSql = strSql & " where UsedDate < #" & Format(DTPicker1.Value, "YYYY-MM-DD") & "#"
End If
With cmdClearHistory
 .ActiveConnection = gCnn
 .CommandText = strSql
 .Execute
End With
Unload Me
ErrorExit:
End Sub

Private Sub Form_Load()
Dim rstCount As New ADODB.Recordset
CenterForm Me
With rstCount
 .Open "select count(*) as lngCount from IPCountHistory", gCnn, adOpenDynamic, adLockOptimistic
 Label1.Caption = Format(!lngCount) & " 条"
 .Close
 .Open "select count(*) as lngCount from AddMoneyHistory", gCnn, adOpenDynamic, adLockOptimistic
 Label2.Caption = Format(!lngCount) & " 条"
 .Close
 .Open "select count(*) as lngCount from UsedMoneyHistory", gCnn, adOpenDynamic, adLockOptimistic
 Label3.Caption = Format(!lngCount) & " 条"
 .Close
End With
End Sub
