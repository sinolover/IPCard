VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmIPCardAppend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "账号信息"
   ClientHeight    =   4140
   ClientLeft      =   1590
   ClientTop       =   3120
   ClientWidth     =   6120
   Icon            =   "frmIpCardAppend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCardType 
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
      Left            =   3420
      TabIndex        =   28
      Top             =   735
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1710
      Left            =   1140
      TabIndex        =   25
      Top             =   1935
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
   Begin VB.TextBox txtCountPW 
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
      Left            =   3420
      MaxLength       =   6
      TabIndex        =   1
      Top             =   210
      Width           =   915
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
      Left            =   975
      MaxLength       =   10
      TabIndex        =   0
      Top             =   210
      Width           =   1455
   End
   Begin VB.TextBox txtRemark 
      Height          =   750
      Left            =   975
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2910
      Width           =   4965
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   400
      Left            =   4665
      TabIndex        =   10
      Top             =   1605
      Width           =   1200
   End
   Begin VB.TextBox txtInitMoney 
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
      Left            =   975
      TabIndex        =   3
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox txtAlert 
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
      Left            =   3420
      TabIndex        =   4
      Top             =   1260
      Width           =   915
   End
   Begin MSAdodcLib.Adodc adcCorporation 
      Height          =   435
      Left            =   5295
      Top             =   2505
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "Select * from Corporation"
      Caption         =   "adcCorporation"
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
   Begin MSAdodcLib.Adodc adcWorker 
      Height          =   435
      Left            =   5460
      Top             =   1650
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
      RecordSource    =   ""
      Caption         =   "adcWorker"
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
   Begin MSAdodcLib.Adodc adcIPCard 
      Height          =   435
      Left            =   5400
      Top             =   180
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from cxipcard"
      Caption         =   "adcIPCard"
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
   Begin VB.TextBox txtCorNo 
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
      Left            =   975
      TabIndex        =   6
      Top             =   2295
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&Q)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4665
      TabIndex        =   11
      Top             =   2250
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "修改(&C)"
      Height          =   400
      Left            =   4665
      TabIndex        =   9
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "增加(&A)"
      Height          =   400
      Left            =   4665
      TabIndex        =   8
      Top             =   315
      Width           =   1200
   End
   Begin VB.TextBox txtWkrName 
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
      Left            =   2340
      TabIndex        =   20
      Top             =   1770
      Width           =   1995
   End
   Begin VB.TextBox txtCorName 
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
      Left            =   2340
      TabIndex        =   12
      Top             =   2295
      Width           =   1995
   End
   Begin VB.TextBox txtWkrNo 
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
      Left            =   975
      TabIndex        =   5
      Top             =   1770
      Width           =   1245
   End
   Begin VB.TextBox txtWithTel 
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
      Left            =   975
      MaxLength       =   10
      TabIndex        =   2
      Top             =   735
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP 账号"
      Height          =   180
      Left            =   225
      TabIndex        =   27
      Top             =   300
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP 密码"
      Height          =   180
      Left            =   2580
      TabIndex        =   26
      Top             =   300
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "其它说明"
      Height          =   180
      Left            =   180
      TabIndex        =   24
      Top             =   2940
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "初始金额"
      Height          =   180
      Left            =   180
      TabIndex        =   23
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "警戒金额"
      Height          =   180
      Left            =   2535
      TabIndex        =   22
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblOprName 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1905
      TabIndex        =   21
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4770
      TabIndex        =   19
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "操作员："
      Height          =   180
      Left            =   1065
      TabIndex        =   18
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "IP 类型"
      Height          =   180
      Left            =   2580
      TabIndex        =   17
      Top             =   825
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "邦定电话"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   825
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "当前日期："
      Height          =   180
      Left            =   3690
      TabIndex        =   15
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "使用单位"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   2385
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "业 务 员"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1860
      Width           =   720
   End
End
Attribute VB_Name = "frmIPCardAppend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lst As ListItem
Dim blnManul As Boolean
Dim intLen As Integer

Private Sub adcCorporation_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True

End Sub

Private Sub adcIPCard_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True

End Sub

Private Sub adcWorker_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAppend_Click()
Dim rstVerify As New ADODB.Recordset
adcIPCard.Refresh
If Not IsNumeric(txtCountNo.Text & "0") Then
   MsgShow "账号必须为数字！"
'   txtCountNo.SetFocus
   Exit Sub
End If
With rstVerify
 If .State <> 0 Then .Close
 .Open "select count(*) as lngCount  from IPCount where rtrim(CountNo)='" & Trim(txtCountNo.Text) & "'", gCnn, adOpenForwardOnly, adLockReadOnly
 If !lngCount > 0 Then
  MsgShow "账号已经存在！"
  Exit Sub
 End If
End With
If Len(Trim(txtAlert.Text)) < 1 Then
 MsgShow "请输入警戒金额"
 txtAlert.SetFocus
 Exit Sub
End If
If Not IsNumeric(txtAlert.Text) Then
 MsgShow "警戒金额必须是数字"
 txtAlert.SetFocus
 Exit Sub
End If
If txtWkrNo_Validate_old Then Exit Sub
If txtCorNo_Validate_old Then Exit Sub
With adcIPCard.Recordset
.AddNew
.Fields("CountNo") = Trim(txtCountNo)
.Fields("CountPW") = IIf(IsNull(txtCountPW), "", Trim(txtCountPW))
.Fields("OprNo") = gOpr.strOprNo
.Fields("AddDate") = CDate(lblDate)
.Fields("CorNo") = Trim(txtCorNo) & ""
!CardType = cboCardType & ""
!AlertMoney = CCur(txtAlert.Text)
!WkrNo = Trim(txtWkrNo) & ""
!AddMoney = CCur(txtInitMoney)
!UsedMoney = 0
!WithTel = Trim(txtWithTel.Text) & ""
!Remark = Trim(txtRemark.Text) & ""
.Update
writeLog Me.Name, "Append", !CountNo
End With
'Dim strID As String
With frmIPCard.adcCard
' strID = Str(.Recordset!No)
 .Refresh
 .Refresh
' .Recordset.Find "ID=" & strID, , adSearchForward, 1
 If Not .Recordset.EOF Then .Recordset.MoveLast
End With
'txtCountNo = ""
'txtCountNo.SetFocus
'Unload Me
txtCountNo.Text = ""
End Sub

Private Sub cmdChange_Click()
'adcIPCard.Refresh
Dim strID As String
If Len(Trim(txtAlert.Text)) < 1 Then
 MsgShow "请输入警戒金额"
 txtAlert.SetFocus
 Exit Sub
End If
If Not IsNumeric(txtAlert.Text) Then
 MsgShow "警戒金额必须是数字"
 txtAlert.SetFocus
 Exit Sub
End If
If txtWkrNo_Validate_old Then Exit Sub
If txtCorNo_Validate_old Then Exit Sub
With frmIPCard.adcCard.Recordset
If Not .EOF And Not .BOF Then
'.Fields("CountNo") = Trim(txtCountNo) & ""
.Fields("CountPW") = Trim(txtCountPW) & ""
.Fields("OprNo") = gOpr.strOprNo
.Fields("CorNo") = Trim(txtCorNo) & ""
!CardType = cboCardType & ""
!AlertMoney = CCur(txtAlert)
!WkrNo = Trim(txtWkrNo) & ""
!WithTel = Trim(txtWithTel.Text) & ""
!Remark = Trim(txtRemark.Text) & ""
If blnRegstered Then
  .Update
 Else
  MsgShow "这是一个试用版本，您将不能够修改账号信息！"
End If
writeLog Me.Name, "Change", !CountNo
Else
MsgShow "请选定待修改记录！"
End If
End With
'frmIPCard.adcCard.Refresh
'If frmIPCard.adcCard.Recordset.RecordCount > 0 Then frmIPCard.adcCard.Recordset.MoveLast
With frmIPCard.adcCard
 strID = Str(.Recordset!No)
' '.MaxRecords
' .LockType
 .Refresh
 .Refresh
 .Recordset.Find "ID=" & strID, , adSearchForward, 1
 'If Not .Recordset.EOF Then .Recordset.MoveLast
End With
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim cmdIPCount As New ADODB.Command
Dim lngrow As Long
On Error Resume Next
lngrow = 1
With frmIPCard
If IsNumeric(.dgCard.Row) Then lngrow = .dgCard.Row
If blnRegstered Then
   cmdIPCount.ActiveConnection = gCnn
   cmdIPCount.CommandText = "delete * from IPCount where rtrim(countno)='" & Trim(txtCountNo.Text) & "'"
   cmdIPCount.Execute
   
   '.Delete adAffectCurrent
   .adcCard.Refresh
   .adcCard.Recordset.Move lngrow
   
 Else
   MsgShow "这是一个试用版本，您将不能够修改账号信息！"
End If
End With
With frmIPCard.adcCard
 .Refresh
 If Not .Recordset.EOF Then .Recordset.MoveLast
End With
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strSql As String
CenterForm Me
blnManul = False
Me.Icon = LoadResPicture(103, 1)
lblOprName = gOpr.strOprName
lblDate = gSysDate
setADODC Me.adcCorporation
setADODC Me.adcIPCard
setADODC Me.adcWorker
With cboCardType
    .AddItem "208"
    .AddItem "17908"
End With

With frmIPCard.adcCard.Recordset
If (Not .EOF) And (Not .BOF) Then
 txtCountNo.Text = .Fields("CountNo") & ""
 txtCountPW.Text = !CountPw & ""
 txtInitMoney = !AddMoney & ""
 txtAlert.Text = !AlertMoney & ""
 txtWkrNo = !WkrNo & ""
 txtWkrName = !WkrName & ""
 txtCorNo.Text = .Fields("CorNo") & ""
 txtCorName.Text = !CorName & ""
 txtWithTel = !WithTel & ""
 txtRemark.Text = !Remark & ""
 cboCardType = !CardType & ""
End If
End With
End Sub

Private Sub txtCorNo_Change()
If Not blnManul Then Exit Sub
intLen = Len(Trim(txtCorNo.Text))
If intLen < 1 Then
   ListView1.Visible = False
   Exit Sub
End If
With ListView1
 .ListItems.Clear
 .Top = txtCorNo.Top - ListView1.Height
 .Left = txtCorNo.Left + intLen * 120
 .Visible = True
End With
With adcWorker
'.RecordSource = "select top 20 corNo,corName from cxCorporation where Left(trim(corno), " & intLen & ") = '" & Trim(txtCorNo.Text) & "'"
.RecordSource = "select top 20 corNo,corName from cxCorporation where corno like '" & Trim(txtCorNo.Text) & "%' order by CorNo"
.Refresh
With .Recordset
While Not .EOF
 Set lst = ListView1.ListItems.Add(, , !CorNo)
 lst.SubItems(1) = !CorName
 .MoveNext
Wend
If .RecordCount = 1 Then
   If Trim(txtCorNo.Text) = Trim(lst.Text) Then
      txtCorName.Text = lst.SubItems(1)
      ListView1.Visible = False
   End If
End If
End With
End With
End Sub

Private Sub txtCorNo_GotFocus()
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "业务单位编号"
ListView1.ColumnHeaders.Add , , "业务单位名称"
ListView1.ListItems.Clear
blnManul = True

End Sub

Private Sub txtCorNo_LostFocus()
ListView1.Visible = False

End Sub

Private Function txtCorNo_Validate_old() As Boolean
txtCorNo_Validate_old = False
If Len(txtCorNo) = 0 Then
 MsgShow "业务单位编号不能为空"
 txtCorNo.SetFocus
 txtCorNo_Validate_old = True
 Exit Function
End If
adcCorporation.RecordSource = "select * from Corporation where CorNo = " & Chr(39) & Trim(txtCorNo.Text) & Chr(39)
adcCorporation.Refresh
If adcCorporation.Recordset.RecordCount < 1 Then
 MsgShow "此业务单位不存在"
 txtCorNo.SetFocus
 txtCorNo_Validate_old = True
 Exit Function
Else
 txtCorName = adcCorporation.Recordset!CorName
End If

End Function

'Private Sub txtCountNo_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode < 95 Or KeyCode > 105 Then KeyCode = 0
'End Sub

Private Sub txtCountNo_Validate(Cancel As Boolean)
'If Len(txtCountNo) <> 10 Then '  intCountNoLen Then
'  MsgShow "账号位数不对，请重新输入！"
'  txtCountNo.SetFocus
'  Cancel = True
'  Exit Sub
'End If
adcIPCard.RecordSource = "select * from IPCount where CountNo= '" & Trim(txtCountNo) & "'"
adcIPCard.Refresh
If adcIPCard.Recordset.RecordCount > 0 Then
 MsgShow "该账号已存在！"
 txtCountNo.SetFocus
 Cancel = True
 Exit Sub
End If
End Sub

Private Sub txtInitMoney_Validate(Cancel As Boolean)
If Not IsNumeric(txtInitMoney.Text) Then
    MsgBox "请输入初始金额！"
    Cancel = True
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListView1.Visible = False
txtWkrNo.Text = Item.Text
End Sub
Private Sub txtWkrNo_GotFocus()
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "业务员编号"
ListView1.ColumnHeaders.Add , , "业务员姓名"
ListView1.ListItems.Clear
blnManul = True
End Sub
Private Sub txtWkrNo_Change()
If Not blnManul Then Exit Sub
intLen = Len(Trim(txtWkrNo.Text))
If intLen < 1 Then
   ListView1.Visible = False
   Exit Sub
End If
With ListView1
 .ListItems.Clear
 .Top = txtWkrNo.Top - ListView1.Height
 .Left = txtWkrNo.Left + intLen * 120
 .Visible = True
End With
With adcWorker
'.RecordSource = "select top 20 WkrNo,WkrName from cxWorker where Left(trim(wkrno), " & intLen & ") = '" & Trim(txtWkrNo.Text) & "'"
.RecordSource = "select top 20 WkrNo,WkrName from cxWorker where wkrno like '" & Trim(txtWkrNo.Text) & "%'"
.Refresh
With .Recordset
While Not .EOF
 Set lst = ListView1.ListItems.Add(, , !WkrNo)
 lst.SubItems(1) = !WkrName
 .MoveNext
Wend
If .RecordCount = 1 Then
   If Trim(txtWkrNo.Text) = Trim(lst.Text) Then
      txtWkrName.Text = lst.SubItems(1)
      ListView1.Visible = False
   End If
End If
End With
End With
End Sub

Private Sub txtWkrNo_LostFocus()
ListView1.Visible = False
End Sub

Private Function txtWkrNo_Validate_old() As Boolean
txtWkrNo_Validate_old = False
If Len(txtWkrNo) = 0 Then
 MsgShow "业务员编号不能为空"
 txtWkrNo.SetFocus
 txtWkrNo_Validate_old = True
 Exit Function
End If
adcWorker.RecordSource = "select * from cxworker where wkrno= '" & Trim(txtWkrNo.Text) & "'"
adcWorker.Refresh
If adcWorker.Recordset.RecordCount < 1 Then
 MsgShow "此业务员不存在"
 txtWkrNo.SetFocus
 txtWkrNo_Validate_old = True
 Exit Function
Else
 txtWkrName = adcWorker.Recordset!WkrName
End If
End Function
Public Sub SetStatus(strStatus As String)
Select Case Trim(strStatus)
       Case "Append"
'        txtCorNo.Text = ""
        cmdAppend.Enabled = Enabled ' IIf(Trim(gOpr.chrAddOpr) = "1", True, False)
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
        Me.Caption = "增加新的账号信息"
       Case "Change"
        txtCountNo.Enabled = False
        txtInitMoney.Enabled = False
        cmdAppend.Enabled = False
        cmdChange.Enabled = Enabled ' IIf(Trim(gOpr.chrEditWkr) = "1", True, False)
        cmdDelete.Enabled = False
        Me.Caption = "对当前的账号信息进行修改"
       Case "Delete"
        txtCountNo.Enabled = False
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = Enabled ' IIf(Trim(gOpr.chrDelWkr) = "1", True, False)
        Me.Caption = "删除当前的账号信息"
       Case Else
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
End Select
End Sub

