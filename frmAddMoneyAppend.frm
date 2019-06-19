VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddMoneyAppend 
   Caption         =   "充值操作"
   ClientHeight    =   3750
   ClientLeft      =   3750
   ClientTop       =   3135
   ClientWidth     =   6150
   Icon            =   "frmAddMoneyAppend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6150
   Begin MSComctlLib.ListView ListView1 
      Height          =   1710
      Left            =   1425
      TabIndex        =   16
      Top             =   1350
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   3016
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "退出(Q)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3165
      TabIndex        =   20
      Top             =   3195
      Width           =   1200
   End
   Begin VB.TextBox txtCountNo 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2085
      Width           =   4650
   End
   Begin VB.TextBox txtOpr 
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1147
      Width           =   1455
   End
   Begin VB.TextBox txtAddCount 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Top             =   686
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adcIPCount 
      Height          =   435
      Left            =   3510
      Top             =   2940
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   767
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
      MaxRecords      =   2
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
      RecordSource    =   "select * from IpCount"
      Caption         =   "adcIPCount"
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
   Begin MSAdodcLib.Adodc adcAddMoney 
      Height          =   435
      Left            =   4230
      Top             =   2955
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
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
      RecordSource    =   "select * from AddMOney"
      Caption         =   "adcAddMoney"
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4620
      TabIndex        =   5
      Top             =   3195
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   1710
      TabIndex        =   4
      Top             =   3195
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   255
      TabIndex        =   3
      Top             =   3195
      Width           =   1200
   End
   Begin VB.TextBox txtCor 
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1608
      Width           =   4635
   End
   Begin VB.TextBox txtWkr 
      Enabled         =   0   'False
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
      Left            =   4500
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1147
      Width           =   1440
   End
   Begin VB.TextBox txtNowCount 
      Enabled         =   0   'False
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
      Left            =   4500
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   686
      Width           =   1440
   End
   Begin VB.TextBox txtAddDate 
      Enabled         =   0   'False
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
      Left            =   4500
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   225
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "充值帐号"
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
      Left            =   165
      TabIndex        =   19
      Top             =   300
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "充值日期"
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
      Left            =   3285
      TabIndex        =   18
      Top             =   300
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "其他说明"
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
      Left            =   165
      TabIndex        =   17
      Top             =   2145
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "使用单位"
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
      Left            =   165
      TabIndex        =   12
      Top             =   1683
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "充前余额"
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
      Left            =   3285
      TabIndex        =   11
      Top             =   761
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "充值金额"
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
      Left            =   165
      TabIndex        =   10
      Top             =   761
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "业 务 员"
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
      Left            =   3285
      TabIndex        =   9
      Top             =   1222
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "操 作 员"
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
      Left            =   165
      TabIndex        =   6
      Top             =   1222
      Width           =   840
   End
End
Attribute VB_Name = "frmAddMoneyAppend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lst As ListItem
Dim intLen As Integer
Dim blnManul As Boolean
Private Sub adcAddMoney_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adcIPCount_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
'Me.Icon = LoadResPicture(103, 1)
blnManul = False
txtAddDate.Text = gSysDate
txtOpr.Text = gOpr.strOprName
setADODC Me.adcAddMoney
setADODC Me.adcIPCount
With frmAddMoney.adcAddMoney.Recordset
If Not .EOF Then
 txtCountNo.Text = .Fields("countno").Value & ""
 txtWkr.Text = .Fields("WkrName").Value & ""
 txtCor.Text = .Fields("CorName").Value & ""
End If
End With
txtCountNo.Enabled = True
End Sub
Private Sub cmdOK_Click()
'If Len(txtCountNo) <> 10 Then
'   mfrmMain.sta.Panels(1).Text = "账号应为10位!"
'   MsgShow "账号应为10位!"
'   txtCountNo.SetFocus
'   GoTo Exit_Sub
'End If
If txtCountNo_Validate_old Then Exit Sub
If Not IsNumeric(txtAddCount) Then
   MsgShow "请输入充值额！"
   txtAddCount.SetFocus
   GoTo Exit_Sub
End If
adcAddMoney.RecordSource = "Select * from AddMoney"
adcAddMoney.MaxRecords = 1
adcAddMoney.Refresh
'adcIPCount.MaxRecords = 2
'adcIPCount.Refresh
With adcAddMoney.Recordset
   .AddNew
  .Fields("CountNo") = txtCountNo & ""
  .Fields("AddDate") = CDate(txtAddDate.Text)
  '.Fields("PreMoney") = IpCount.curPreMoney
  .Fields("AddMoney") = Val(txtAddCount.Text)
  .Fields("NowMoney") = Val(txtNowCount) + Val(txtAddCount.Text)
  .Fields("OprNo") = gOpr.strOprNo
  .Fields("EditNo") = "1"
  If Not blnRegstered Then
     If !AddMoney > 88 Then
        MsgShow "这是一个试用版本，因此充值总额额不可以大于 88 元"
        txtAddCount.SetFocus
        Exit Sub
     End If
  End If
  !Remark = Trim(txtRemark)
  .Update
  writeLog Me.Name, "Append", !CountNo
End With
adcIPCount.Recordset.Update "AddMoney", CCur(txtAddCount.Text) + adcIPCount.Recordset!AddMoney
txtCountNo.SetFocus
Exit_Sub:
'mfrmMain.sta.Panels(1).Text = "状态"
With frmAddMoney.adcAddMoney
 .Refresh
 If Not .Recordset.EOF Then .Recordset.MoveLast
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not frmAddMoney.adcAddMoney.Recordset.EOF Then frmAddMoney.adcAddMoney.Recordset.MoveLast
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListView1.Visible = False
txtCountNo.Text = lst.Text
End Sub
Private Sub txtCountNo_Change()
If Not IsNumeric(txtCountNo.Text & "0") Then
   MsgShow "账号必须为数字！"
   txtCountNo.SetFocus
   Exit Sub
End If
intLen = Len(Trim(txtCountNo.Text))
If Not blnManul Then Exit Sub
If intLen = 0 Then
   ListView1.Visible = False
   Exit Sub
End If
'If intLen = 10 Then
'   ListView1.Visible = False
'   With adcIPCount
'   .RecordSource = "select * from cxIPCount where countno= " _
'           & "'" & Trim(txtCountNo) & "'"
'   .Refresh
'   If .Recordset.RecordCount <> 1 Then
'       MsgShow "该账号不存在，请确认！"
'       Exit Sub
'   End If
'   txtWkr = .Recordset!WkrName & ""
'   txtCor = .Recordset!CorName & ""
'   txtNowCount = .Recordset!NowMoney & ""
'   End With
'   Exit Sub
'End If
With ListView1
 .ListItems.Clear
 .Top = txtCountNo.Top + txtCountNo.Height
 .Left = txtCountNo.Left + intLen * 120
 If intLen <> 10 Then .Visible = True
End With
With adcIPCount
.RecordSource = "select top 20 CountNo,WkrName from cxipcount where rtrim(countno) like '" & Trim(txtCountNo.Text) & "%'"
.Refresh
With .Recordset
'If .RecordCount = 1 Then
'   txtCountNo.Text = !CountNo
'   Exit Sub
'End If
While Not .EOF
 Set lst = ListView1.ListItems.Add(, , !CountNo)
 lst.SubItems(1) = !WkrName
 .MoveNext
Wend
If .RecordCount = 1 Then
   If Trim(txtCountNo.Text) = Trim(lst.Text) Then
      With adcIPCount
     .RecordSource = "select * from cxIPCount where countno= " _
            & "'" & Trim(txtCountNo) & "'"
     .Refresh
'   If .Recordset.RecordCount <> 1 Then
'       MsgShow "该账号不存在，请确认！"
'       Exit Sub
'   End If
     txtWkr = .Recordset!WkrName & ""
     txtCor = .Recordset!CorName & ""
     txtNowCount = .Recordset!NowMoney & ""
     End With
     ListView1.Visible = False
   End If
End If
End With
End With
End Sub
Private Sub txtCountNo_GotFocus()
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "账号"
ListView1.ColumnHeaders.Add , , "业务员"
ListView1.ListItems.Clear
blnManul = True
If Len(Trim(txtCountNo.Text)) <> 10 Then ListView1.Visible = True
End Sub
Private Sub txtCountNo_LostFocus()
ListView1.Visible = False
End Sub
Private Function txtCountNo_Validate_old() As Boolean
txtCountNo_Validate_old = False
With adcIPCount
'If Len(txtCountNo) <> 10 Then
'   MsgShow "账号应为10位!"
'   txtCountNo.SetFocus
'   'Cancel = False
'   txtCountNo_Validate_old = True
'   GoTo Exit_Sub
'End If
.RecordSource = "select * from cxIPCount where countno= " _
           & "'" & Trim(txtCountNo) & "'" '& "and WkrName = '" & gOpr.strOprName & "'"
'adccxIPCard.RecordSource = "select * from cxIPCard where countno= " _
'           & "'" & Trim(txtCountNo) & "'"
'adccxIPCard.Refresh
.Refresh
If .Recordset.RecordCount <> 1 Then
   MsgShow "该账号不存在，请确认！"
   txtCountNo.SetFocus
   'txtCountNo.SelText
   txtCountNo_Validate_old = True
   GoTo Exit_Sub
End If
txtWkr = .Recordset!WkrName & ""
txtCor = .Recordset!CorName & ""
txtNowCount = .Recordset!NowMoney & ""
End With
'With adcIPCount
' .RecordSource = "select * from cxIpCount where countno=" _
'              & "'" & Trim(txtCountNo) & "'"
' .Refresh
' If .Recordset.RecordCount <> 1 Then
'   .Recordset.AddNew "countno,PreMoney,NowMoney,UsedMoney", Trim(txtCountNo) & "0,0"
' End If
 'txtNowCount = .Recordset.Fields("NowMoney")
' txtNowCount = .Recordset.Fields("NowMoney") & ""
'End With
Exit_Sub:
End Function

