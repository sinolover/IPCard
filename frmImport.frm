VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据导入"
   ClientHeight    =   2370
   ClientLeft      =   3435
   ClientTop       =   3120
   ClientWidth     =   4680
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc adcTmpIpCount 
      Height          =   495
      Left            =   1260
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Connect         =   "DSN=IpCard"
      OLEDBString     =   "DSN=IpCard"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from ipcount"
      Caption         =   "adcTmpIPCount"
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
   Begin MSAdodcLib.Adodc adcTmpImport 
      Height          =   495
      Left            =   600
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "Select top 1 * from UsedMoney"
      Caption         =   "adcTmpImport"
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
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览..."
      Height          =   400
      Left            =   3285
      TabIndex        =   4
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   1725
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   165
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox txtImport 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   1200
      Width           =   2835
   End
   Begin MSComDlg.CommonDialog cdlImport 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPriImport 
      Caption         =   "上次导入文件："
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblImport 
      Caption         =   "导入文件:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1260
      Width           =   1095
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gstrFileName As String
Dim strFilePath As String
Private Sub rstTmpImport_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub adcTmpIpCount_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub
Private Sub Form_Load()
Dim rstImport As New ADODB.Recordset
writeLog Me.Name
CenterForm Me
Me.Icon = LoadResPicture(103, 1)
lblPriImport.Caption = lblPriImport.Caption & GetReg("LastFileName")
cdlImport.InitDir = GetReg("LastFileDir", "C:\")
setADODC Me.adcTmpImport
setADODC Me.adcTmpIpCount
End Sub
Private Sub cmdOK_Click()
If Trim(txtImport.Text) & "" = "" Then
   MsgShow "请输入文件名！", vbOKOnly
  Exit Sub
End If
If Len(Trim(txtImport.Text)) < 5 Then
   MsgShow "文件名格式错误，请检查！", vbOKOnly
   Exit Sub
End If
If Not Exists(Trim(txtImport.Text)) Then
   MsgShow "文件不存在！", vbOKOnly
   Exit Sub
End If
cmdOK.Enabled = False
Me.Hide
OpenXls (txtImport.Text)
'Unload frmImport
writeLog Me.Name, "Ok", txtImport.Text
End Sub
Private Sub cmdCancle_Click()
writeLog Me.Name, "Cancel", "未进行导入操作"
Unload Me
End Sub
Private Sub cmdBrowse_Click()
Dim intLen As Integer
With cdlImport
 .Filter = "Excel文档|*.xls"
 .InitDir = GetReg("ImportPath", "C:")
 .DefaultExt = "xls"
'cdlimport.InitDir=
 .CancelError = True
 On Error GoTo Error_Exit
 .ShowOpen
 intLen = Len(Trim(.FileName))
 txtImport = .FileName
 gstrFileName = .FileTitle
 strFilePath = Left(.FileName, intLen - Len(.FileTitle) - 1)
 SaveReg "ImportPath", strFilePath
End With
Error_Exit:
End Sub
Private Function OpenXls(ByVal strFileName As String) As Boolean
Dim intLine As Integer
Dim intCol As Integer
Dim intKH As Integer, intSC As Integer, intFY As Integer, intHTH As Integer, intOther As Integer
Dim blnDisplayAlerts As Boolean
Dim strFileDate As String
Dim xlsImport As New Excel.Application
Dim rstUsedMoney As New ADODB.Recordset
Dim rstIPCount As New ADODB.Recordset
Dim strSql As String

frmStatus.Show
frmStatus.ShowMsg "正在准备数据，请稍候..."
If IsDate(strFileName) Then
   strFileDate = Format(strFileName, "YYYY-MM-DD")
  Else
   strFileDate = Left(Right(strFileName, 8), 4)
   strFileDate = Format(Date, "YYYY") & "-" & Left(strFileDate, 2) & "-" & Right(strFileDate, 2)
End If
If Not IsDate(strFileDate) Then
   strFileDate = InputBox("系统无法解析文件名，请手工输入文件日期", "IP 卡帐务管理系统")
End If
If IsDate(strFileDate) Then
OpenXls = True
On Error GoTo Err_Exit
gCnn.BeginTrans
xlsImport.Workbooks.Open strFileName
On Error Resume Next
rstTmpImport.RecordSource = "Select top 1 * from UsedMoney"
rstTmpImport.Refresh
For intCol = 1 To 8
 'intLine = 2
 Select Case UCase(Trim(xlsImport.Cells(1, intCol)))
        Case "KH"
         intKH = intCol
        Case "SC"
         intSC = intCol
        Case "FY"
         intFY = intCol
        Case "HTH"
         intHTH = intCol
        Case Else
         intOther = 0
 End Select
Next intCol
If intKH = 0 Or intFY = 0 Then
   MsgShow "无法分析日志文件，请确认！"
   GoTo Error_Exit
End If
intLine = 2
While IsNumeric(xlsImport.Cells(intLine, intKH)) And Len(xlsImport.Cells(intLine, intKH)) > 0
 With rstTmpImport.Recordset
   .AddNew
   .Fields("CountNo") = xlsImport.Cells(intLine, intKH)
   .Fields("UsedDate") = CDate(strFileDate)
   .Fields("UsedMoney") = Format(xlsImport.Cells(intLine, intFY) / 100)
   .Fields("UsedID") = "2"
   .Fields("UsedTime") = Format(Val(.Fields("UsedTime") & "") + xlsImport.Cells(intLine, intSC))
   '.Fields("ID") = " "
   .Update
   frmStatus.ShowMsg "正在处理第" & Str(intLine) & " 行 " & !CountNo
End With
adcTmpIpCount.MaxRecords = 1
adcTmpIpCount.RecordSource = "Select * from IPCount where CountNO = '" & Trim(xlsImport.Cells(intLine, intKH)) & "'"
adcTmpIpCount.Refresh
Dim blnImport As Boolean
blnImport = False
With adcTmpIpCount.Recordset
 If .RecordCount < 1 Then
    Dim strOpr As String
    'frmStatus.Hide
    Select Case GetReg("NewCount", "0")
           Case "0"
            If MsgShow("账号:" & Trim(xlsImport.Cells(intLine, intKH)) _
               & "不存在是否导入", vbYesNo) = vbYes Then blnImport = True
           Case "1" 'Ingnore
            blnImport = False
           Case "2" 'Append
            blnImport = True
           Case Else
            blnImport = True
    End Select
    If blnImport Then
   .AddNew
   .Fields("CountNO") = xlsImport.Cells(intLine, intKH)
   .Fields("UsedMoney") = xlsImport.Cells(intLine, intFY) / 100
   '!OprNo = Trim(gOpr.strOprNo) & ""
   strOpr = InputBox("请输入业务员编号", "注意", "1001")
   strOpr = InputBox("请输入业务单位编号", "注意", "1001")
   !AddMoney = 0
   !AddDate = Date
   !AlertMoney = 0
   !WkrNo = Trim(strOpr) & ""
   !CorNo = strOpr
   .Update
   End If
  Else
   !UsedMoney = Val(!UsedMoney & "") + xlsImport.Cells(intLine, intFY) / 100
   If IsDate(!lastdate) Then
       If DateDiff("d", !lastdate, CDate(strFileDate)) > 0 Then !lastdate = CDate(strFileDate)
   Else
        !lastdate = CDate(strFileDate)
   End If
  .Update
 End If
End With
intLine = intLine + 1
Wend
With rstUsedMoney
 .Open "select Count(UsedTime) as lngUsedTime,count(UsedMoney) as lngUsedMoney from usedmoney where Useddate=# " & Format(strFileDate, "YYYY-MM-DD") & " # ", gCnn, adOpenStatic, adLockReadOnly
End With
rstTmpImport.RecordSource = "select * from ImportRZ"
rstTmpImport.Refresh
With rstTmpImport.Recordset
 .AddNew
 .Fields("FileName") = gstrFileName ' cdlImport.FileName
 .Fields("FilePath") = strFilePath ' Left(cdlImport.FileTitle, Len(cdlImport.FileName) - Len(cdlImport.FileTitle))
 .Fields("FileDate") = CDate(strFileDate)
 .Fields("ImportDate") = Date
 .Fields("OprNo") = gOpr.strOprNo
 .Fields("UsedID") = "2"
 !RecNum = intLine
 !CountTotal = IIf(IsNull(rstUsedMoney!lngUsedMoney), 0, rstUsedMoney!lngUsedMoney)
 !TimeTotal = IIf(IsNull(rstUsedMoney!lngUsedTime), 0, rstUsedMoney!lngUsedTime)
 .Update
End With
SaveReg "LastFileName", gstrFileName
SaveReg "LastFileDir", strFilePath
frmStatus.ShowMsg "正在准备数据..."
With rstUsedMoney
 If .State <> 0 Then .Close
 .Open "delete from UsedMoneyHistory where UsedDate < # " & Format(DateAdd("d", 31, CDate(strFileDate)), "YYYY-MM-DD") & " #"
End With
Else
 MsgShow "Sorry The file can't be reverse!"
 MsgShow strFileDate
 OpenXls = False
End If
Error_Exit:
blnDisplayAlerts = xlsImport.Workbooks.Application.DisplayAlerts
xlsImport.Workbooks.Application.DisplayAlerts = False
xlsImport.Workbooks.Close
xlsImport.Workbooks.Application.DisplayAlerts = blnDisplayAlerts
xlsImport.Quit
Err_Exit:
Unload frmStatus
If Err.Number = 0 Then
    MsgShow "导入顺利完成！"
Else
    On Error Resume Next
    MsgShow "ErrNumber:" & Err.Number & vbCrLf & "Err Description:" & Err.Description
End If
Unload Me
End Function

