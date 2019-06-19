VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "Comdlg32.OCX"
Begin VB.Form frmPrintPreview 
   Caption         =   "Visual Basic ʵս�� - Ԥ����ӡ"
   ClientHeight    =   6084
   ClientLeft      =   1836
   ClientTop       =   2496
   ClientWidth     =   5292
   LinkTopic       =   "Form2"
   ScaleHeight     =   6084
   ScaleWidth      =   5292
   WindowState     =   2  '���
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
      Caption         =   "����"
      Begin VB.Menu mnuSetup 
         Caption         =   "�����趨 ..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "��ӡ ..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "�ر�"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "����"
      Begin VB.Menu mnuView100 
         Caption         =   "100%"
      End
      Begin VB.Menu mnuViewFullPage 
         Caption         =   "��ҳ"
      End
      Begin VB.Menu mnuViewCustomize 
         Caption         =   "�Զ��ٷֱ� ..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPage 
         Caption         =   "�л�ҳ�� ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "˵��"
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1.��ʽ���ƣ�Ԥ����ӡ
'2.�������ڣ�09/02/1999
'3.����������Visual Basic 5.0 ����רҵ�� + SP3
'4.���������������� (С��,Jaric)
'5.�������䣺jaric@tacocity.com.tw
'6.������ַ��http://fly.to/jaric �� http://tacocity.com.tw/jaric
'7.��ַ���ƣ�Visual Basic ʵս��
'8.ע���������������ɢ������ʽ��������������˵��ɾ����лл��
'                     �������ʽ���������޸ģ��������·��������ĸ�����Ѷ��
Option Explicit
Private Const dblHWRATIO As Double = 297 / 210 'A4ֽ�ŵĳ����
Private Const dblWHRATIO As Double = 210 / 297 'A4ֽ�ŵĿ���
Private Const lngVSPACE As Long = 100 'ҳ��ҳ֮��Ĵ�ֱ����   ��λ:twips
Private lngPages As Long '������ӡҳ��
Private lngViewRatio As Long '��ʾ���� ,���� 0 ~ ���޴����ֵ,ͨ������0~100����
Private gX As Long, gY As Long '�ƶ�ͼ��֮ǰ���������
'lngZoom�Ǵ���������ӡ�� PictureBox ʱ�ı���,���� 0 ~ ���޴����ֵ
'ͨ������0~100���� ,�������ֵ�����ý϶�ļ����� , ͬʱ��С����ʧ��
'��С����ֵ���õļ�������� ,���ǷŴ����ʧ��,
'�벻Ҫ��lngZoom�� lngViewRatio ���,lngViewRatio��ͼ�λ��ƺ�֮���� imgView֮�ڵ���ʾ����
'���� blnRePrint=True ��ÿ�θı� lngViewRatio �������º��� PrintResult ����ͼ
'����Ԥ����ӡ�Ľ����û��ʧ��֮�� ,�����ٶȽ���
'�� blnRePrint=false , ��ÿ�θı� lngViewRatio ���������»�ͼ ,����ֱ�Ӹı� imgView�Ĵ�С�Է����µ���ʾ����
'�����ٶȺܿ� ,����Ԥ����ӡ�Ľ����ʧ��
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

'frameA �ȱ�Сʱһ��Ҫλ�ڱ�������,���µĳ�ʽ�����work
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
     s = "1. �����Ǵ�����ӡ�ı߽磬������ӡʱ����ӡ������" & vbNewLine
     s = s & "2. �����ʽû�о��ᣬ�����û�����ҷͼƬ�Ϳ��Կ������е���ӡ���ϡ�" & vbNewLine
     s = s & "3.���Ҫ��ÿ�θı���ʾ����ʱ�����»�ͼ���������ܱ�� /����/�����趨/ ���趨��" & vbNewLine
     s = s & "4. �ڱ��ϰ������Ҽ������ʾ ""����"" ���ܱ�"
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
     ' ���߽�ֵ�Ƿ񱻸��Ĺ� ,������������ӡ�����Է����µı߽�ֵ
     If lngTM <> glngTopMargin Or lngBM <> glngBottomMargin Or lngLM <> glngLeftMargin _
     Or lngRM <> glngRightMargin Or plngZoom <> lngZoom Then
          If Not blnRePrint Then Call gobjFormToPrint.PrintResult(picPrint, lngZoom)
          Call ChangeRatio
     End If
End Sub

Private Sub mnuView100_Click()
     If lngViewRatio = 100 Then Exit Sub '���Ŀǰ��ʾ�ı����������ı�����ͬ , ��Ҫ���»�ͼ
     lngViewRatio = 100
     Call ChangeRatio
End Sub

Private Sub mnuViewCustomize_Click()
     Dim X As String
     X = InputBox("����������ʾ�İٷֱȣ�", "Visual Basic ʵս�� http://fly.to/jaric", CLng(lngViewRatio))
     If Trim(X) = "" Then Exit Sub
     If Not IsNumeric(X) Or InStr(X, ",") > 0 Or InStr(X, "-") > 0 Or Val(X) = 0 Then
          MsgBox "���������ֵ��������������"
     Else
          If lngViewRatio = CLng(X) Then Exit Sub '���Ŀǰ��ʾ�ı����������ı�����ͬ , ��Ҫ���»�ͼ
          lngViewRatio = CLng(X)
          Call ChangeRatio
     End If
End Sub

Private Sub mnuViewFullPage_Click()
     Call FullPage
End Sub

Private Sub mnuViewPage_Click()
     Dim X As String, n As Long
     X = InputBox("����������ʾ��ҳ�룬", "Visual Basic ʵս�� http://fly.to/jaric", "1")
     If Trim(X) = "" Then Exit Sub
     If Not IsNumeric(X) Or InStr(X, ",") > 0 Or InStr(X, ".") > 0 Then
          MsgBox "��������� 0 ���Ҳ����� " & lngPages & " ֮����"
     Else
          n = CLng(X)
          If n <= 0 Or n > lngPages Then
               MsgBox "��������� 0 ���Ҳ����� " & lngPages & " ֮����"
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
          '���Ŀǰ��ʾ�ı����������ı�����ͬ , ��Ҫ���»�ͼ
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
     Caption = "Visual Basic ʵս�� - Ԥ����ӡ" & " ( ���� " & lngPages & " ҳ�����ǵ� " & n & " ҳ����ʾ������" & CLng(lngViewRatio) & "%)"
End Sub
