VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   4380.183
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdvance 
      Caption         =   "高级>>"
      Height          =   400
      Left            =   3060
      TabIndex        =   6
      Top             =   1065
      Width           =   1200
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1650
      TabIndex        =   1
      Text            =   "1001"
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   300
      TabIndex        =   4
      Top             =   1065
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   1680
      TabIndex        =   5
      Top             =   1065
      Width           =   1200
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   540
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:sllin@QingDaoNews.com"
      Height          =   180
      Left            =   1065
      TabIndex        =   8
      Top             =   2100
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "非常感谢您的使用，若有任何问题请"
      Height          =   180
      Left            =   885
      TabIndex        =   7
      Top             =   1830
      Width           =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   0
      X2              =   4380.183
      Y1              =   948.287
      Y2              =   948.287
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   4380.183
      Y1              =   966.012
      Y2              =   966.012
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "用户名称(&U):"
      Height          =   180
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   225
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "密码(&P):"
      Height          =   180
      Index           =   1
      Left            =   465
      TabIndex        =   2
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adcLogin_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub Form_Load()
Dim i As Integer
CenterForm Me
'setADODC Me.adcLogin
frmLogin.Icon = LoadResPicture(102, 1)
frmLogo.Hide
Unload frmLogo
End Sub
Private Sub cmdAdvance_Click()
If Trim(cmdAdvance.Caption) = "高级>>" Then
   cmdAdvance.Caption = "简化<<"
   Me.Height = 2760
 Else
   Me.Height = 1920
   cmdAdvance.Caption = "高级>>"
End If
End Sub
Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    Unload Me
    Unload mfrmMain
    End
End Sub
Private Sub cmdOK_Click()
Dim rst As New ADODB.Recordset
Dim strSql As String
Dim rstReg As New ADODB.Recordset
'       mfrmMain.Enabled = True
'       Unload Me
'       GoTo Exit_Sub
'**********Verify Password Beginning
'On Error GoTo Err_Exit
'gCnn.BeginTrans
If Not IsNumeric(txtUserName) Or Len(txtUserName) = 0 Then GoTo Clear_Exit
strSql = "select * from operator where OprNo =" & Chr(39) & Trim(txtUserName) & Chr(39)
rst.Open strSql, gCnn, adOpenForwardOnly, adLockReadOnly
With rst
    If Not .EOF Then
       If Trim(txtPassword) = Trim(.Fields("oprpw") & "") Then
    '       mfrmMain.Enabled = True
           gOpr.strOprNo = Trim(txtUserName)
           gOpr.strOprName = Trim(!OprName) & ""
           gOpr.strOprSex = Trim(!OprSex) & ""
           gOpr.intOprAge = Val(!OprAge)
           gOpr.chrOprType = Trim(!OprType) & ""
           If GetReg("WkrLogin", "1") = "0" Then
              If gOpr.chrOprType <> "1" Then
                 MsgShow "对不起，您不可以登陆，请联系管理员！"
                 txtUserName.SetFocus
                 Exit Sub
              End If
           End If
            If Not blnRegstered Then
               Me.Hide
               Load frmWait
               frmWait.Show vbModal
               Unload frmWait
               With rstReg
                If .State <> 0 Then .Close
                .Open "select RegCount from RegID", gCnn, adOpenKeyset, adLockOptimistic
                !RegCount = !RegCount + 1
                .Update
                If !RegCount > 30 Then
                   If !RegCount Mod (Val(Format(Time, "SS")) Mod 30 + 1) = 0 Then
                      MsgShow "该程序执行了非法操作,即将被关闭", , "错误"
                      End
                   End If
                   MsgShow "对不起，您已经超过试用期了！"
                   End
                End If
                .Close
                End With
            End If
           'gOpr.chrAddCount = !AddCount
           'If !OprType = "1" Then
           '  blnOprType = True
           ' Else
           '  blnOprType = False
           'End If
           'gOpr.chrOprType = !OprType
           mfrmMain.Enabled = True
           writeLog
           Unload Me
           'mfrmMain.Show
           GoTo Exit_Sub
       End If
    End If
End With

Clear_Exit:
MsgShow "无效的用户名密码，请重试!"
txtPassword.SetFocus
SendKeys "{Home}+{End}"
Exit_Sub:
AppActivate App.Title
'Exit Sub
On Error Resume Next
rst.Close
Set rst = Nothing
Set rstReg = Nothing
End Sub

'*****************Verify Password End
