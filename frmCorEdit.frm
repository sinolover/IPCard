VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCorEdit 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frmCorEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "全局通话数量调整"
      Height          =   270
      Left            =   795
      TabIndex        =   29
      Top             =   3555
      Width           =   1740
   End
   Begin VB.CheckBox Check1 
      Caption         =   "全局通话时间调整"
      Height          =   270
      Left            =   810
      TabIndex        =   28
      Top             =   3060
      Width           =   1755
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   825
      TabIndex        =   26
      Top             =   2550
      Width           =   4185
   End
   Begin VB.TextBox Text1 
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
      Left            =   3330
      TabIndex        =   23
      Top             =   210
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adcWorker 
      Height          =   465
      Left            =   5160
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   820
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=IPCard"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "IPCard"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   1710
      Left            =   4605
      TabIndex        =   16
      Top             =   2730
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
   Begin VB.CommandButton cmdAppend 
      Caption         =   "增加(&A)"
      Height          =   400
      Left            =   5205
      TabIndex        =   10
      Top             =   465
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(D)"
      Height          =   400
      Left            =   5205
      TabIndex        =   12
      Top             =   1775
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "修改(&C)"
      Height          =   400
      Left            =   5205
      TabIndex        =   11
      Top             =   1120
      Width           =   1200
   End
   Begin VB.TextBox txtCorPhone 
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
      Left            =   3330
      TabIndex        =   6
      Top             =   1770
      Width           =   1695
   End
   Begin VB.TextBox txtCorMan 
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
      Left            =   840
      TabIndex        =   5
      Top             =   1770
      Width           =   1830
   End
   Begin VB.TextBox txtCorFax 
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
      Left            =   3330
      TabIndex        =   4
      Top             =   1380
      Width           =   1695
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
      Height          =   885
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3990
      Width           =   5670
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
      Left            =   1950
      TabIndex        =   8
      Top             =   2160
      Width           =   1320
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
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(Q)"
      Height          =   400
      Left            =   5205
      TabIndex        =   13
      Top             =   2430
      Width           =   1200
   End
   Begin VB.TextBox txtCorTel 
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
      Left            =   840
      TabIndex        =   3
      Top             =   1380
      Width           =   1830
   End
   Begin VB.TextBox txtCorAddr 
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
      Left            =   840
      TabIndex        =   2
      Top             =   990
      Width           =   4185
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
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   4185
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
      Left            =   840
      TabIndex        =   0
      Top             =   210
      Width           =   1830
   End
   Begin MSAdodcLib.Adodc adcCor 
      Height          =   495
      Left            =   5280
      Top             =   2760
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
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
      RecordSource    =   "select * from cxCorporation"
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "EMail"
      Height          =   180
      Left            =   135
      TabIndex        =   27
      Top             =   2640
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "合同"
      Height          =   195
      Left            =   2775
      TabIndex        =   25
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "传真"
      Height          =   180
      Left            =   2775
      TabIndex        =   24
      Top             =   1470
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编  号"
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   300
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "名  称"
      Height          =   195
      Left            =   135
      TabIndex        =   21
      Top             =   690
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "地  址"
      Height          =   195
      Left            =   135
      TabIndex        =   20
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "电  话"
      Height          =   195
      Left            =   135
      TabIndex        =   19
      Top             =   1470
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "业务员"
      Height          =   180
      Left            =   135
      TabIndex        =   18
      Top             =   2235
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "联系人"
      Height          =   180
      Left            =   135
      TabIndex        =   17
      Top             =   1860
      Width           =   525
   End
   Begin VB.Label Label9 
      Caption         =   "电话"
      Height          =   210
      Left            =   2790
      TabIndex        =   15
      Top             =   1845
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "备  注"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   3990
      Width           =   450
   End
End
Attribute VB_Name = "frmCorEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCorporation As New ADODB.Recordset
Dim lst As ListItem
Dim blnManul As Boolean
Private Sub adcCor_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub adcWorker_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub cmdAppend_Click()
If Len(Trim(txtCorNo)) < 1 Then
   MsgShow "请输入公司编号！"
   txtCorNo.SetFocus
   Exit Sub
End If
With adcCor
.RecordSource = "select * from cxCorporation where CorNO='" & Trim(txtCorNo.Text) & "'"
.Refresh
If .Recordset.RecordCount > 0 Then
   MsgShow "该编号已经存在！"
   txtCorNo.SetFocus
   Exit Sub
End If
If Len(Trim(txtCorName)) < 1 Then
   MsgShow "请输入公司名称"
   txtCorName.SetFocus
   Exit Sub
End If
If txtWkrNo_Validate_old Then
   MsgShow "该业务员不存在！"
   txtWkrNo.SetFocus
   Exit Sub
End If
With .Recordset
.AddNew
!CorNo = Trim(txtCorNo.Text)
!CorName = Trim(txtCorName.Text)
!CorAddr = Trim(txtCorAddr.Text)
!CorTel = Trim(txtCorTel)
!CorFax = Trim(txtCorFax)
!CorMan = Trim(txtCorMan)
!CorPhone = Trim(txtCorPhone)
!WkrNo = Trim(txtWkrNo) & ""
!oprNo = gOpr.strOprNo
.Update
writeLog Me.Name, "Append", !CorName
End With
End With
With frmCorporation.adcCor
 .Refresh
 .Refresh
 If Not .Recordset.EOF Then .Recordset.MoveLast
End With
Unload Me
End Sub

Private Sub cmdChange_Click()
Dim strID As String
If Len(Trim(txtCorName.Text)) < 1 Then
   MsgShow "请输入公司名称"
   txtCorName.SetFocus
   Exit Sub
End If
If txtWkrNo_Validate_old Then
   MsgShow "该业务员不存在！"
   txtWkrNo.SetFocus
   Exit Sub
End If
With frmCorporation.adcCor
With .Recordset
If Not .EOF And Not .BOF Then
!CorName = Trim(txtCorName.Text) & ""
!CorAddr = Trim(txtCorAddr.Text) & ""
!CorTel = Trim(txtCorTel) & ""
!CorFax = Trim(txtCorFax) & ""
!CorMan = Trim(txtCorMan) & ""
!CorPhone = Trim(txtCorPhone) & ""
!WkrNo = Trim(txtWkrNo) & ""
!oprNo = gOpr.strOprNo & ""
!Remark = txtRemark.Text & ""
.Update
writeLog Me.Name, "Change", !CorName
Else
 MsgShow "当前记录无法修改!"
End If
End With
End With
Unload Me
With frmCorporation.adcCor
 strID = .Recordset!CorNo
 .Refresh
 .Refresh
' If Not .Recordset.EOF Then .Recordset.MoveLast
 .Recordset.Find "CorNo=" & Str(strID), , adSearchForward, 1
End With
'txtCorNo.SetFocus
'frmCorporation.adcCor.Refresh
End Sub
Private Sub cmdDelete_Click()
If MsgShow("确认删除？", vbYesNo) = vbNo Then Exit Sub
Unload Me
Exit Sub
On Error Resume Next
adcCor.RecordSource = "delete * from corporation where rtrim(corno)='" & Trim(txtCorNo.Text) & "'"
adcCor.Refresh
'With frmCorporation.adcCor
' .Refresh
' If Not .Recordset.EOF Then .Recordset.MoveLast
'End With
'frmCorporation.adcCor.Recordset.Delete adAffectCurrent
'writelog me.Name,"Change",
frmCorporation.adcCor.Refresh
frmCorporation.adcCor.Refresh
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
CenterForm Me
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "业务员账号"
ListView1.ColumnHeaders.Add , , "业务员姓名"
ListView1.ListItems.Clear
blnManul = False
setADODC Me.adcCor
setADODC Me.adcWorker
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListView1.Visible = False
txtWkrNo.Text = Item.Text
End Sub
Private Sub txtWkrNo_GotFocus()
blnManul = True
End Sub
Private Sub txtWkrNo_Change()
If Not blnManul Then Exit Sub
intLen = Len(Trim(txtWkrNo.Text))
If intLen = 0 Then Exit Sub
With ListView1
 .ListItems.Clear
 .Top = txtWkrNo.Top - ListView1.Height
 .Left = txtWkrNo.Left + intLen * 120
 .Visible = True
End With
With adcWorker
.RecordSource = "select top 20 WkrNo,WkrName from cxWorker where Left(trim(wkrno), " & intLen & ") = '" & Trim(txtWkrNo.Text) & "'"
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
'rstCorporation.Source = "select * from cxworker where wkrno= '" & Trim(txtWkrNo.Text) & "'"
'rstCorporation.Open , gCnn, adOpenStatic, adLockReadOnly
adcWorker.RecordSource = "select * from cxworker where wkrno= '" & Trim(txtWkrNo.Text) & "'"
adcWorker.Refresh
'If rstCorporation.RecordCount < 1 Then
If adcWorker.Recordset.RecordCount < 0 Then
 MsgShow "此业务员不存在"
 txtWkrNo.SetFocus
 txtWkrNo_Validate_old = True
 Exit Function
Else
' txtWkrName = rstCorporation!WkrName
 textwkrname = adcWorker.Recordset!WkrName
End If
End Function
Public Sub SetStatus(Optional strStatus As String = "Browse")
With frmCorporation.adcCor.Recordset
If Not .EOF And Not .BOF Then
txtCorNo.Text = !CorNo
txtCorName.Text = IIf(IsNull(!CorName), "", !CorName)
txtCorAddr.Text = IIf(IsNull(!CorAddr), "", !CorAddr)
txtCorTel.Text = IIf(IsNull(!CorTel), "", !CorTel)
txtCorFax.Text = IIf(IsNull(!CorFax), "", !CorFax)
txtCorMan.Text = IIf(IsNull(!CorMan), "", !CorMan)
txtCorPhone.Text = IIf(IsNull(!CorPhone), "", !CorPhone)
txtWkrNo.Text = IIf(IsNull(!WkrNo), "", !WkrNo)
txtRemark.Text = IIf(IsNull(!Remark), "", !Remark)
End If
End With
Select Case Trim(strStatus)
       Case "Append"
        txtCorNo.Text = ""
        txtCorName.Text = ""
        txtCorAddr.Text = ""
        txtCorTel.Text = ""
        txtCorFax.Text = ""
        txtCorMan.Text = ""
        txtCorPhone.Text = ""
        txtWkrNo.Text = ""
        txtRemark.Text = ""
        cmdAppend.Enabled = Enabled ' IIf(Trim(gOpr.chrAddOpr) = "1", True, False)
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
        Me.Caption = "增加新的业务单位信息"
       Case "Change"
        txtCorNo.Enabled = False
        cmdAppend.Enabled = False
        cmdChange.Enabled = Enabled ' IIf(Trim(gOpr.chrEditWkr) = "1", True, False)
        cmdDelete.Enabled = False
        Me.Caption = "对当前的业务单位信息进行修改"
       Case "Delete"
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = Enabled ' IIf(Trim(gOpr.chrDelWkr) = "1", True, False)
        Me.Caption = "删除当前的业务单位信息"
       Case Else
        cmdAppend.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
End Select
End Sub
