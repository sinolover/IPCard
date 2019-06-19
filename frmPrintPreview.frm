VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "Comdlg32.OCX"
Begin VB.Form frmPrintPreview 
   Caption         =   "Visual Basic 实战网 - 预览列印"
   ClientHeight    =   6084
   ClientLeft      =   1836
   ClientTop       =   2496
   ClientWidth     =   5292
   LinkTopic       =   "Form2"
   ScaleHeight     =   6084
   ScaleWidth      =   5292
   WindowState     =   2  '最大化
   Begin VB.PictureBox picPrint 
      Height          =   1476
      Left            =   1872
      ScaleHeight     =   1428
      ScaleWidth      =   1548
      TabIndex        =   1
      Top             =   3672
      Width           =   1596
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   1296
      Top             =   4320
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327681
   End
   Begin VB.Frame frameA 
      Caption         =   "Frame1"
      Height          =   2940
      Left            =   1032
      TabIndex        =   0
      Top             =   504
      Width           =   2916
      Begin VB.Image imgView 
         Height          =   1932
         Index           =   1
         Left            =   288
         Top             =   624
         Width           =   1764
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "档案"
      Begin VB.Menu mnuSetup 
         Caption         =   "版面设定 ..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "列印 ..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "关闭"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "检视"
      Begin VB.Menu mnuView100 
         Caption         =   "100%"
      End
      Begin VB.Menu mnuViewFullPage 
         Caption         =   "整页"
      End
      Begin VB.Menu mnuViewCustomize 
         Caption         =   "自订百分比 ..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPage 
         Caption         =   "切换页码 ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "说明"
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1.程式名称：预览列印
'2.开发日期：09/02/1999
'3.开发环境：Visual Basic 5.0 中文专业版 + SP3
'4.作者姓名：宋世杰 (小翰,Jaric)
'5.作者信箱：jaric@tacocity.com.tw
'6.作者网址：http://fly.to/jaric 或 http://tacocity.com.tw/jaric
'7.网址名称：Visual Basic 实战网
'8.注意事项：您可以任意散布本程式，但是请勿将以上说明删除，谢谢！
'                     如果本程式经过您的修改，可以在下方加入您的个人资讯。
Option Explicit
Private Const dblHWRATIO As Double = 297 / 210 'A4纸张的长宽比
Private Const dblWHRATIO As Double = 210 / 297 'A4纸张的宽长比
Private Const lngVSPACE As Long = 100 '页与页之间的垂直距离   单位:twips
Private lngPages As Long '储存列印页数
Private lngViewRatio As Long '显示比例 ,介于 0 ~ 无限大的数值,通常输入0~100即可
Private gX As Long, gY As Long '移动图形之前储存的座标
'lngZoom是代表将资料列印到 PictureBox 时的比例,介于 0 ~ 无限大的数值
'通常输入0~100即可 ,愈大的数值将耗用较多的记忆体 , 同时缩小后易失真
'愈小的数值耗用的记忆体较少 ,但是放大后易失真,
'请不要将lngZoom与 lngViewRatio 搞混,lngViewRatio是图形绘制好之后在 imgView之内的显示比例
'若将 blnRePrint=True 则每次改变 lngViewRatio 都会重新呼叫 PrintResult 来绘图
'这样预览列印的结果将没有失真之虞 ,但是速度较慢
'若 blnRePrint=false , 则每次改变 lngViewRatio 并不会重新绘图 ,而是直接改变 imgView的大小以符合新的显示比例
'这样速度很快 ,但是预览列印的结果会失真
Public lngZoom As Long
Public blnRePrint As Boolean

Private Sub Form_Load()
     frameA.Caption = ""
     frameA.BorderStyle = vbBSNone
     imgView(1).BorderStyle = vbBSNone
     imgView(1).Width = glngPAPERW
     imgView(1).Height = glngPAPERH
     imgView(1).Stretch = True
     picPrint.BorderStyle = vbBSNone
     picPrint.BackColor = vbWhite
     picPrint.ScaleMode = vbTwips
     picPrint.AutoRedraw = True
     picPrint.Visible = False
     If lngZoom = 0 Then lngZoom = 100
     If Not blnRePrint Then
          Call gobjFormToPrint.PrintResult(picPrint, lngZoom)
          lngPages = imgView.Count
     End If
     lngViewRatio = 100
     Call ChangeRatio
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = vbRightButton Then PopupMenu mnuView
End Sub

Private Sub Form_Resize()
     Call FramePosition
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Dim i As Long
     For i = lngPages To 2 Step -1
          Set imgView(i).Picture = LoadPicture()
          Unload imgView(i)
     Next
     Set imgView(1).Picture = LoadPicture()
     picPrint.AutoRedraw = False
End Sub

Private Sub frameA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     gX = X
     gY = Y
     If Button = vbRightButton Then PopupMenu mnuView
End Sub

'frameA 比表单小时一定要位于表单的中央,以下的程式码才能work
Private Sub frameA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not Button = vbLeftButton Then Exit Sub
     Dim dx As Long, dy As Long, ax As Long, ay As Long, t As Long, l As Long, tt As Long, ll As Long
     With frameA
          dy = Y - gY
          dx = X - gX
          ll = .Left
          tt = .Top
          l = Abs(ll)
          t = Abs(tt)
          ax = (.Width - l - ScaleWidth)
          ay = (.Height - t - ScaleHeight)
          If ll > 0 Then
               dx = 0
          Else
               If dx < 0 Then
                    If Abs(dx) > ax Then dx = -ax
               Else
                    If dx > l Then dx = l
               End If
          End If
          If tt > 0 Then
               dy = 0
          Else
               If dy < 0 Then
                    If Abs(dy) > ay Then dy = -ay
               Else
                    If dy > t Then dy = t
               End If
           End If
          .Move ll + dx, tt + dy
     End With
End Sub

Private Sub mnuClose_Click()
     Unload Me
End Sub

Private Sub mnuHelp_Click()
     Dim s As String
     s = "1. 虚线是代表列印的边界，真正列印时不会印出来。" & vbNewLine
     s = s & "2. 这个程式没有卷轴，但是用滑鼠拖曳图片就可以看到所有的列印资料。" & vbNewLine
     s = s & "3.如果要在每次改变显示比例时都重新绘图，请至功能表的 /档案/版面设定/ 内设定。" & vbNewLine
     s = s & "4. 在表单上按滑鼠右键亦可显示 ""检视"" 功能表。"
     MsgBox s
End Sub

Private Sub mnuPrint_Click()
     On Error GoTo ErrorTrap
     Dim i As Long
     dlgPrint.CancelError = True
     dlgPrint.PrinterDefault = True
     dlgPrint.Flags = cdlPDDisablePrintToFile + cdlPDNoSelection  '+ cdlPDUseDevModeCopies
     dlgPrint.ShowPrinter
     For i = 1 To dlgPrint.Copies
          Call gobjFormToPrint.PrintResult(Printer, lngZoom)
     Next
ErrorTrap:
End Sub

Private Sub mnuSetup_Click()
     Dim lngTM As Long, lngBM As Long, lngLM As Long, lngRM As Long
     Dim i As Long, plngZoom As Long
     lngTM = glngTopMargin
     lngBM = glngBottomMargin
     lngLM = glngLeftMargin
     lngRM = glngRightMargin
     plngZoom = lngZoom
     frmPrintSetup.Show vbModal, Me
     ' 检查边界值是否被更改过 ,若是则重新列印资料以符合新的边界值
     If lngTM <> glngTopMargin Or lngBM <> glngBottomMargin Or lngLM <> glngLeftMargin _
     Or lngRM <> glngRightMargin Or plngZoom <> lngZoom Then
          If Not blnRePrint Then Call gobjFormToPrint.PrintResult(picPrint, lngZoom)
          Call ChangeRatio
     End If
End Sub

Private Sub mnuView100_Click()
     If lngViewRatio = 100 Then Exit Sub '如果目前显示的比例与期望的比例相同 , 则不要重新绘图
     lngViewRatio = 100
     Call ChangeRatio
End Sub

Private Sub mnuViewCustomize_Click()
     Dim X As String
     X = InputBox("请输入欲显示的百分比，", "Visual Basic 实战网 http://fly.to/jaric", CLng(lngViewRatio))
     If Trim(X) = "" Then Exit Sub
     If Not IsNumeric(X) Or InStr(X, ",") > 0 Or InStr(X, "-") > 0 Or Val(X) = 0 Then
          MsgBox "您输入的数值有误，请重新输入"
     Else
          If lngViewRatio = CLng(X) Then Exit Sub '如果目前显示的比例与期望的比例相同 , 则不要重新绘图
          lngViewRatio = CLng(X)
          Call ChangeRatio
     End If
End Sub

Private Sub mnuViewFullPage_Click()
     Call FullPage
End Sub

Private Sub mnuViewPage_Click()
     Dim X As String, n As Long
     X = InputBox("请输入欲显示的页码，", "Visual Basic 实战网 http://fly.to/jaric", "1")
     If Trim(X) = "" Then Exit Sub
     If Not IsNumeric(X) Or InStr(X, ",") > 0 Or InStr(X, ".") > 0 Then
          MsgBox "请输入大于 0 并且不大于 " & lngPages & " 之整数"
     Else
          n = CLng(X)
          If n <= 0 Or n > lngPages Then
               MsgBox "请输入大于 0 并且不大于 " & lngPages & " 之整数"
               Exit Sub
          Else
               Call ChangePage(n)
          End If
     End If
End Sub

Private Sub imgView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     gX = X
     gY = Y
     Call changeCaption(Index)
     If Button = vbRightButton Then PopupMenu mnuView
End Sub

Private Sub imgView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call frameA_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub ChangeRatio()
     Dim i As Long, w As Long
     If blnRePrint Then
          lngZoom = lngViewRatio
          Call gobjFormToPrint.PrintResult(picPrint, lngZoom)
          lngPages = imgView.Count
     End If
     w = glngPAPERW * (lngViewRatio / 100)
     For i = 1 To lngPages
          imgView(i).Move 0, (i - 1) * (w * dblHWRATIO + lngVSPACE), w, w * dblHWRATIO
     Next
     frameA.Move 0, 0, imgView(1).Width, (imgView(1).Height + lngVSPACE) * lngPages
     Call FramePosition
End Sub

Private Sub FullPage()
     Dim i As Long, w As Long, h As Long, wBase As Boolean, ratio As Long
     w = ScaleWidth
     h = ScaleHeight
     If CDbl(w / h) >= dblWHRATIO Then
          ratio = h / glngPAPERH * 100
          '如果目前显示的比例与期望的比例相同 , 则不要重新绘图
          If lngViewRatio = ratio Then Exit Sub Else lngViewRatio = ratio
     Else
          ratio = w / glngPAPERW * 100
          If lngViewRatio = ratio Then Exit Sub Else lngViewRatio = ratio: wBase = True
     End If
     If blnRePrint Then
          lngZoom = lngViewRatio
          Call gobjFormToPrint.PrintResult(picPrint, lngZoom)
          lngPages = imgView.Count
     End If
     For i = 1 To lngPages
          If wBase Then
               imgView(i).Move 0, (i - 1) * (w * dblHWRATIO + lngVSPACE), w, w * dblHWRATIO
          Else
               imgView(i).Move 0, (i - 1) * (h + lngVSPACE), h * dblWHRATIO, h
          End If
     Next
     frameA.Move 0, 0, imgView(1).Width, (imgView(1).Height + lngVSPACE) * lngPages
     Call FramePosition
End Sub

Private Sub ChangePage(n As Long)
     frameA.Move frameA.Left, -(imgView(1).Height + lngVSPACE) * (n - 1)
     Call changeCaption(n)
End Sub

Private Sub FramePosition()
     Dim w As Long, h As Long, fw As Long, fh As Long
     fw = frameA.Width
     fh = frameA.Height
     w = ScaleWidth
     h = ScaleHeight
     If fh < h And fw < w Then
          frameA.Move (w - fw) / 2, (h - fh) / 2
     ElseIf fh < h Then
          frameA.Move 0, (h - fh) / 2
     ElseIf fw < w Then
          frameA.Move (w - fw) / 2, 0
     Else
          frameA.Move 0, 0
     End If
     Call changeCaption(1)
End Sub

Public Sub changeCaption(ByVal n As Long)
     Caption = "Visual Basic 实战网 - 预览列印" & " ( 共有 " & lngPages & " 页，这是第 " & n & " 页，显示比例：" & CLng(lngViewRatio) & "%)"
End Sub
