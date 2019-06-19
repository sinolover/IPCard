VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSysSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统参数设置"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSysSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3390
      TabIndex        =   15
      Top             =   3840
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6482
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "视图设置"
      TabPicture(0)   =   "frmSysSetup.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkToolShow"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkStatusShow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Internet"
      TabPicture(1)   =   "frmSysSetup.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "邮箱设置"
      TabPicture(2)   =   "frmSysSetup.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "其他"
      TabPicture(3)   =   "frmSysSetup.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chkStatusShow 
         Caption         =   "状态栏"
         Height          =   300
         Left            =   480
         TabIndex        =   22
         Top             =   1455
         Width           =   1620
      End
      Begin VB.CheckBox chkToolShow 
         Caption         =   "工具栏"
         Height          =   300
         Left            =   480
         TabIndex        =   21
         Top             =   885
         Width           =   1950
      End
      Begin VB.Frame Frame1 
         Caption         =   "帐务系统设置"
         Height          =   2955
         Left            =   -74880
         TabIndex        =   2
         Top             =   555
         Width           =   4470
         Begin VB.CheckBox chkNo 
            Caption         =   "显示编号"
            Height          =   285
            Left            =   315
            TabIndex        =   23
            Top             =   2190
            Width           =   1155
         End
         Begin VB.OptionButton optIngnore 
            Caption         =   "忽略"
            Height          =   300
            Left            =   3330
            TabIndex        =   20
            Top             =   1710
            Width           =   720
         End
         Begin VB.OptionButton optAppend 
            Caption         =   "加入"
            Height          =   300
            Left            =   1800
            TabIndex        =   19
            Top             =   1710
            Width           =   720
         End
         Begin VB.OptionButton optTW 
            Caption         =   "提示"
            Height          =   300
            Left            =   270
            TabIndex        =   18
            Top             =   1725
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.CheckBox chkNewCount 
            Caption         =   "当发现新的账号"
            Height          =   300
            Left            =   315
            TabIndex        =   17
            Top             =   1215
            Width           =   1560
         End
         Begin VB.CheckBox chkStartUp 
            Caption         =   "启动警戒员"
            Height          =   255
            Left            =   2325
            TabIndex        =   16
            Top             =   1215
            Width           =   1335
         End
         Begin VB.CheckBox chkWkrLogin 
            Caption         =   "允许业务员登陆"
            Height          =   285
            Left            =   315
            TabIndex        =   8
            Top             =   780
            Width           =   1620
         End
         Begin VB.TextBox txtAlertMoney 
            Alignment       =   1  'Right Justify
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
            Left            =   3885
            TabIndex        =   7
            Text            =   "0"
            Top             =   720
            Width           =   435
         End
         Begin VB.CheckBox chkAlertMoney 
            Caption         =   "设置警戒金额"
            Height          =   330
            Left            =   2325
            TabIndex        =   6
            Top             =   750
            Width           =   1410
         End
         Begin VB.CheckBox chkCountPW 
            Caption         =   "显示账号密码"
            Height          =   285
            Left            =   330
            TabIndex        =   5
            Top             =   330
            Width           =   1380
         End
         Begin VB.CheckBox chkAlertDate 
            Caption         =   "设置警戒天数"
            Height          =   375
            Left            =   2325
            TabIndex        =   4
            Top             =   270
            Width           =   1395
         End
         Begin VB.TextBox txtAlertDate 
            Alignment       =   1  'Right Justify
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
            Left            =   3885
            TabIndex        =   3
            Text            =   "3"
            Top             =   255
            Width           =   435
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "确认(&O)"
      Height          =   400
      Left            =   1815
      TabIndex        =   0
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "警戒员设置"
      Height          =   1440
      Left            =   120
      TabIndex        =   9
      Top             =   2085
      Visible         =   0   'False
      Width           =   4470
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   3885
         TabIndex        =   14
         Text            =   "3"
         Top             =   255
         Width           =   435
      End
      Begin VB.CheckBox Check4 
         Caption         =   "设置警戒天数"
         Height          =   375
         Left            =   2325
         TabIndex        =   13
         Top             =   270
         Width           =   1395
      End
      Begin VB.CheckBox chkTrayCountPW 
         Caption         =   "显示账号密码"
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1380
      End
      Begin VB.CheckBox Check2 
         Caption         =   "设置警戒金额"
         Height          =   330
         Left            =   2340
         TabIndex        =   11
         Top             =   900
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   3885
         TabIndex        =   10
         Text            =   "0"
         Top             =   885
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmSysSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Load()
CenterForm Me
'If IsConnected = True Then
'    msgshow ("已经连接了Internet!")
'End If
'If IsConnected = False Then
'    msgshow ("还没有连接 Internet!")
'End If
'Me.Icon = LoadResPicture(103, 1)
Dim strRegValue As String
If GetReg("IPolice") = "1" Then
   chkStartUp.Value = 1
  Else
   chkStartUp.Value = 0
End If
If GetReg("ID") = "1" Then
    chkNo.Value = 1
Else
    chkNo.Value = 0
End If
If GetReg("CountPW") = "1" Then
    chkCountPW.Value = 1
Else
    chkCountPW = 0
End If

chkNewCount.Value = 1
Select Case GetReg("NewCount")
    Case "0"
        optTW.Value = True
    Case "1"
        optIngnore.Value = True
    Case "2"
        optAppend.Value = True
    Case Else
        chkNewCount.Value = 0
End Select

strRegValue = GetReg("AlertMoney", "0")
'chkAlertMoney.Value = IIf(IsNumeric(strRegValue), CInt(strRegValue), 0)
chkAlertMoney.Value = IIf(IsNumeric(strRegValue), 1, 0)
txtAlertMoney.Text = strRegValue
strRegValue = GetReg("AlertDate", "0")
'chkAlertDate.Value = IIf(IsNumeric(strRegValue), CInt(strRegValue), 0)
chkAlertDate.Value = IIf(IsNumeric(strRegValue), 1, 0)
txtAlertDate.Text = strRegValue
strRegValue = GetReg("WkrLogin", "0")
chkWkrLogin.Value = IIf(IsNumeric(strRegValue), CInt(strRegValue), 0)
strRegValue = GetReg("NewCount", "0")
Select Case strRegValue
       Case "0"
        optTW.Value = True
       Case "1"
        optIngnore.Value = True
       Case "2"
        optAppend.Value = True
       Case Else
        optTW.Value = True
End Select
strRegValue = GetReg("ToolShow", "1")
If strRegValue = "1" Then
   chkToolShow.Value = 1
  Else
   chkToolShow.Value = 0
End If
strRegValue = GetReg("StatusShow", "1")
If strRegValue = "1" Then
   chkStatusShow.Value = 1
  Else
   chkStatusShow.Value = 0
End If
writeLog Me.Name
End Sub
Private Sub cmdOK_Click()
'SaveReg "IPolice", "3"
If chkStartUp.Value = 1 Then
    SaveReg "IPolice", "1"
    UpdateKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "IPolice", App.Path & "\IPolice.exe", &H80000002
Else
    SaveReg "IPolice", "0"
    DelKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "IPolice", &H80000002
End If
SaveReg "WkrLogin", chkWkrLogin.Value
SaveReg "AlertMoney", txtAlertMoney.Text
SaveReg "AlertDate", txtAlertDate.Text
SaveReg "CountPW", chkCountPW.Value
SaveReg "ToolShow", chkToolShow.Value
SaveReg "StatusShow", chkStatusShow.Value
SaveReg "ID", chkNo.Value
With mfrmMain
If chkToolShow.Value = 1 Then
   .Toolbar1.Visible = True
  Else
   .Toolbar1.Visible = False
End If
If chkStatusShow.Value = 1 Then
   .sta.Visible = True
  Else
   .sta.Visible = False
End If
End With
'SaveReg "CountPW", chkTrayCountPW, "IPTray"
If optTW.Value Then SaveReg "NewCount", "0"
If optIngnore.Value Then SaveReg "NewCount", "1"
If optAppend.Value Then SaveReg "NewCount", "2"
   
Unload Me
End Sub

Private Sub txtAlertDate_Change()
If Not IsNumeric(txtAlertDate.Text) Then
   MsgShow "请输入数字！"
   txtAlertDate.SetFocus
   Exit Sub
End If
'If Val(txtAlertDate.Text) < 2 Or Val(txtAlertDate.Text) > 10 Then
'   MsgShow "请输入介于 3 到 10 之间的数字！"
'   txtAlertDate.SetFocus
'   Exit Sub
'End If
End Sub
Private Sub txtAlertMoney_Change()
If Not IsNumeric(txtAlertMoney.Text) Then
   MsgShow "请输入数字！"
   txtAlertDate.SetFocus
   Exit Sub
End If
If Val(txtAlertDate.Text) < -10 Or Val(txtAlertDate.Text) > 10 Then
   MsgShow "请输入介于 -10 到 10 之间的数字！"
   txtAlertDate.SetFocus
   Exit Sub
End If
End Sub
