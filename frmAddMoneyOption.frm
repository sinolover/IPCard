VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddMoneyOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "充值结转"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmAddMoneyOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3090
      TabIndex        =   7
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "开始(&B)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3090
      TabIndex        =   6
      Top             =   300
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   360
      Left            =   2790
      TabIndex        =   5
      Top             =   1575
      Width           =   1560
      _ExtentX        =   2752
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
      Format          =   59310081
      CurrentDate     =   37424
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   1155
      TabIndex        =   4
      Top             =   1575
      Width           =   1560
      _ExtentX        =   2752
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
      Format          =   59310081
      CurrentDate     =   37424
   End
   Begin VB.OptionButton optBetween 
      Caption         =   "此间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1605
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpBefore 
      Height          =   360
      Left            =   1155
      TabIndex        =   2
      Top             =   930
      Width           =   1560
      _ExtentX        =   2752
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
      Format          =   59310081
      CurrentDate     =   37424
   End
   Begin VB.OptionButton optBefore 
      Caption         =   "此前"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   810
   End
   Begin VB.OptionButton optAll 
      Caption         =   "结转全部充值信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   315
      Value           =   -1  'True
      Width           =   2265
   End
End
Attribute VB_Name = "frmAddMoneyOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim rstAddMoney As New ADODB.Recordset
Dim cmdAddMoney As New ADODB.Command
Dim rstAddMoneyHistory As New ADODB.Recordset
Dim i As Integer
Dim lngrow As Long
Dim strSql As String
writeLog Me.Name, "Begin"
If optAll Then
   strSql = " * from AddMoney"
End If
If optBefore Then
   strSql = " * from AddMoney where  AddDate <= #" & Format(dtpBefore.Value, "YYYY-MM-DD") & "#"
End If
If optBetween Then
   MsgShow "对不起，您不能采用此方式转出！"
   Exit Sub
End If
Me.Hide
frmStatus.Show
rstAddMoney.Open "select " & strSql, gCnn, adOpenStatic, adLockReadOnly
rstAddMoneyHistory.Open "select top 1 * from AddMoneyHistory", gCnn, adOpenStatic, adLockOptimistic
lngrow = 1
With rstAddMoneyHistory
While Not rstAddMoney.EOF
 .AddNew
 For i = 1 To .Fields.Count - 2
  .Fields(i) = rstAddMoney.Fields(i)
 Next i
 !ZZDate = Date
 frmStatus.ShowMsg "正在转出使用使用日志，账号：" & !CountNo
 .Update
 rstAddMoney.MoveNext
 lngrow = lngrow + 1
Wend
.Close
rstAddMoney.Close
cmdAddMoney.CommandText = "delete " & strSql
cmdAddMoney.ActiveConnection = gCnn
cmdAddMoney.Execute
Unload frmStatus
MsgShow "转账完毕！"
Unload Me
End With
writeLog Me.Name, "End", "共计：" & Str(lngrow)
End Sub
Private Sub Form_Load()
CenterForm Me
End Sub
