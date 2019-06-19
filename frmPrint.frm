VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表输出"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optWkrMonthly 
      Caption         =   "月报表"
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
      Left            =   3315
      TabIndex        =   17
      Top             =   2100
      Width           =   1400
   End
   Begin VB.OptionButton optWkrDaily 
      Caption         =   "日报表"
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
      Left            =   1800
      TabIndex        =   16
      Top             =   2100
      Width           =   1400
   End
   Begin VB.OptionButton optCorMonthly 
      Caption         =   "月报表"
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
      Left            =   3315
      TabIndex        =   14
      Top             =   1710
      Width           =   1400
   End
   Begin VB.OptionButton optCorDaily 
      Caption         =   "日报表"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   1710
      Width           =   1400
   End
   Begin VB.OptionButton optAddMoney 
      Caption         =   "充值表"
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
      Left            =   3315
      TabIndex        =   11
      Top             =   1275
      Width           =   1400
   End
   Begin VB.OptionButton optAlertDate 
      Caption         =   "日期警戒"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1275
      Width           =   1400
   End
   Begin VB.OptionButton optAlertMoney 
      Caption         =   "金额警戒"
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
      Left            =   285
      TabIndex        =   9
      Top             =   1275
      Width           =   1400
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   1875
      TabIndex        =   8
      Top             =   2475
      Width           =   1200
   End
   Begin VB.OptionButton optOpr 
      Caption         =   "业务员"
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
      Left            =   3315
      TabIndex        =   7
      Top             =   825
      Width           =   1400
   End
   Begin VB.OptionButton optCount 
      Caption         =   "账号信息"
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
      Left            =   1792
      TabIndex        =   6
      Top             =   825
      Width           =   1400
   End
   Begin VB.OptionButton optCor 
      Caption         =   "业务单位"
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
      Left            =   270
      TabIndex        =   5
      Top             =   825
      Value           =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   400
      Left            =   3555
      TabIndex        =   4
      Top             =   2475
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "输出(&E)"
      Default         =   -1  'True
      Height          =   400
      Left            =   195
      TabIndex        =   3
      Top             =   2475
      Width           =   1200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览..."
      Height          =   400
      Left            =   3525
      TabIndex        =   2
      Top             =   195
      Width           =   1200
   End
   Begin VB.TextBox txtFileName 
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
      Left            =   1245
      TabIndex        =   1
      Top             =   210
      Width           =   2040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3615
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "业务员"
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
      Left            =   555
      TabIndex        =   15
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "公司"
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
      Left            =   555
      TabIndex        =   12
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件名称"
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
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   840
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
With CommonDialog1
 .CancelError = False
 '.DefaultExt = "mdb"
 .Filter = "Excel文件|*.xls"
 .ShowOpen
 txtFileName = .FileName
End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim strSql As String
Dim strKey As String
Dim strDBF As String
Dim strAlert As String
If Not Exists(txtFileName.Text) Then
   MsgShow "该文件不存在"
   Exit Sub
End If
If optCor.Value Then
 strKey = "frmCorporation"
 strDBF = " cxCorporation"
End If
If optOpr.Value Then
 strKey = "frmOperator"
 strDBF = " cxOperator"
End If
If optCount.Value Then
 strKey = "frmIpCard"
 strDBF = " cxIpCount"
End If
If optAlertDate.Value Then
 strKey = "frmAlertDate"
 strDBF = "cxAlertMoney"
End If
If optAlertMoney.Value Then
 strKey = "frmAlertMoney"
 strDBF = "cxAlertMoney"
End If
If optAddMoney.Value Then
 strKey = "frmAlertMoney"
 strDBF = "cxAlertMoney"
End If
If optCorDaily.Value Then
 strKey = "frmCorDaily"
 strDBF = "cxCorDaily"
End If
If optCorMonthly.Value Then
 strKey = "frmCorMonthly"
 strDBF = "cxCorMonthly"
End If
If optWkrDaily.Value Then
 strKey = "frmWkrDaily"
 strDBF = "cxWkrDaily"
End If
If optWkrMonthly.Value Then
 strKey = "frmWkrMonthly"
 strDBF = "cxWkrMonthly"
End If
strSql = GetSqlString(strKey)
If strSql = "" Then strSql = "*"
strSql = "select " & strSql & " from " & strDBF
If optAlertDate.Value Then
   strAlert = GetReg("AlertDate", "2")
   If Not IsNumeric(strAlert) Then strAlert = "2"
   strSql = strSql & " where AlertMoney > 0 and  (" _
         & " LastDate <= # " _
         & Format(DateAdd("d", Val(strAlert) * (-1), Date), "YYYY-MM-DD") _
         & " # or isnull(LastDate) ) order by LastDate DESC, WkrName"
End If
If optAlertMoney.Value Then
   strAlert = GetReg("AlertMoney", "0")
   If Not IsNumeric(strAlert) Then strAlert = "0"
   strSql = strSql & " where " _
         & " AlertMoney - nowMoney >=" & strAlert _
         & " order by alertMoney DESC,WkrName"
End If
sysPrint strSql, txtFileName.Text
Unload Me
End Sub
Private Sub sysPrint(strSql As String, strFileName As String)
Dim rstPrint As New ADODB.Recordset
Dim xlsOutPut As New Excel.Application
Dim i As Integer
Dim lngrow As Long
xlsOutPut.Workbooks.Open (strFileName)
If Len(strSql) < 10 Then Exit Sub
If optAddMoney Then
   With rstPrint
   .Open "select * from cxaddmoney where adddate = # " & _
    Format(gSysDate, "YYYY-MM-DD") & " #", gCnn, adOpenForwardOnly, adLockReadOnly
   End With
   With xlsOutPut
     .Cells(1, 1) = "序号"
     .Cells(1, 2) = "卡号"
     .Cells(1, 3) = "金额"
     .Cells(1, 4) = "大写"
     .Cells(1, 5) = "备注"
     lngrow = 2
    While Not rstPrint.EOF And Not rstPrint.BOF
     .Cells(lngrow, 1) = lngrow - 1
     .Cells(lngrow, 2) = rstPrint!CountNo
     .Cells(lngrow, 3) = rstPrint!AddMoney
     '.Cells(lngRow, 4) = ""
     .Cells(lngrow, 5) = rstPrint!FromCount
     lngrow = lngrow + 1
     rstPrint.MoveNext
    Wend
  End With
 Else
With rstPrint
 .Open strSql, gCnn, adOpenStatic, adLockReadOnly
 With xlsOutPut
  For i = 0 To rstPrint.Fields.Count - 1
   .Cells(1, i + 1) = rstPrint.Fields.Item(i).Name
  Next i
  lngrow = 2
  While Not rstPrint.EOF And Not rstPrint.BOF
   For i = 0 To rstPrint.Fields.Count - 1
    .Cells(lngrow, i + 1) = rstPrint.Fields(i).Value
   Next i
   rstPrint.MoveNext
   lngrow = lngrow + 1
  Wend
 End With
End With
End If
rstPrint.Close
xlsOutPut.Workbooks(1).Save
xlsOutPut.Quit
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

