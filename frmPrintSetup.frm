VERSION 5.00
Begin VB.Form frmPrintSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "版面设定"
   ClientHeight    =   4050
   ClientLeft      =   2850
   ClientTop       =   2700
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4950
   Begin VB.Frame Frame2 
      Caption         =   "预览条件"
      Height          =   1332
      Left            =   336
      TabIndex        =   16
      Top             =   1992
      Width           =   4308
      Begin VB.TextBox txtZoom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3072
         TabIndex        =   19
         Top             =   912
         Width           =   564
      End
      Begin VB.OptionButton optRePrint 
         Caption         =   "不重新绘图，但是重新设定原始图片的比例"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   648
         Width           =   3732
      End
      Begin VB.OptionButton optRePrint 
         Caption         =   "每次改变显示比例时都重新绘图"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   264
         Width           =   3300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   180
         Left            =   3720
         TabIndex        =   20
         Top             =   960
         Width           =   144
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "预设值"
      Height          =   372
      Left            =   3648
      TabIndex        =   15
      Top             =   3504
      Width           =   972
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "取消"
      Height          =   372
      Left            =   2568
      TabIndex        =   5
      Top             =   3504
      Width           =   972
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "确定"
      Height          =   372
      Left            =   1512
      TabIndex        =   4
      Top             =   3504
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "边界值"
      Height          =   1596
      Left            =   312
      TabIndex        =   6
      Top             =   216
      Width           =   4308
      Begin VB.TextBox txtMargin 
         Alignment       =   2  'Center
         Height          =   348
         Index           =   4
         Left            =   3048
         TabIndex        =   3
         Top             =   936
         Width           =   804
      End
      Begin VB.TextBox txtMargin 
         Alignment       =   2  'Center
         Height          =   348
         Index           =   3
         Left            =   1008
         TabIndex        =   2
         Top             =   936
         Width           =   804
      End
      Begin VB.TextBox txtMargin 
         Alignment       =   2  'Center
         Height          =   348
         Index           =   2
         Left            =   3048
         TabIndex        =   1
         Top             =   384
         Width           =   804
      End
      Begin VB.TextBox txtMargin 
         Alignment       =   2  'Center
         Height          =   348
         Index           =   1
         Left            =   1008
         TabIndex        =   0
         Top             =   384
         Width           =   804
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   180
         Index           =   7
         Left            =   3888
         TabIndex        =   14
         Top             =   1008
         Width           =   204
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   180
         Index           =   6
         Left            =   1872
         TabIndex        =   13
         Top             =   984
         Width           =   204
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "右边界："
         Height          =   180
         Index           =   3
         Left            =   2304
         TabIndex        =   12
         Top             =   1008
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "左边界："
         Height          =   180
         Index           =   2
         Left            =   264
         TabIndex        =   11
         Top             =   1008
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   180
         Index           =   5
         Left            =   3888
         TabIndex        =   10
         Top             =   456
         Width           =   204
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   180
         Index           =   4
         Left            =   1872
         TabIndex        =   9
         Top             =   432
         Width           =   204
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下边界："
         Height          =   180
         Index           =   1
         Left            =   2304
         TabIndex        =   8
         Top             =   456
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上边界："
         Height          =   180
         Index           =   0
         Left            =   264
         TabIndex        =   7
         Top             =   456
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPrintSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDefault_Click()
     Dim i As Long
     For i = 1 To 4
          txtMargin(i).Text = "1.0"
     Next
     optRePrint(1).Value = True
     txtZoom.Text = 100
End Sub

Private Sub cmdNo_Click()
     Unload Me
End Sub

Private Sub cmdYes_Click()
     Dim i As Long, s As String
     For i = 1 To 4
          s = txtMargin(i).Text
          If Not IsNumeric(s) Or InStr(s, ",") > 0 Or Val(s) < 0 Or Val(s) > 10 Then
               MsgBox "边界值请输入大于 0 并且不大于 10 的数值。"
               txtMargin(i).SetFocus
               Exit Sub
          End If
     Next
     s = txtZoom.Text
     If Not IsNumeric(s) Or InStr(s, ",") > 0 Or InStr(s, "-") > 0 Or Val(s) = 0 Or InStr(s, ".") Then
          MsgBox "您输入的图片比例有误，请输入一个正整数。"
          txtZoom.SelStart = 0
          txtZoom.SelLength = Len(s)
          txtZoom.SetFocus
          Exit Sub
     End If
     frmPrintPreview.lngZoom = txtZoom.Text
     frmPrintPreview.blnRePrint = optRePrint(0).Value
     glngTopMargin = txtMargin(1).Text * 567
     glngBottomMargin = txtMargin(2).Text * 567
     glngLeftMargin = txtMargin(3).Text * 567
     glngRightMargin = txtMargin(4).Text * 567
     Unload Me
End Sub

Private Sub Form_Load()
     txtMargin(1) = Format(glngTopMargin / 567, "0.0")
     txtMargin(2) = Format(glngBottomMargin / 567, "0.0")
     txtMargin(3) = Format(glngLeftMargin / 567, "0.0")
     txtMargin(4) = Format(glngRightMargin / 567, "0.0")
     txtZoom.Text = frmPrintPreview.lngZoom
     If frmPrintPreview.blnRePrint Then optRePrint(0).Value = True Else optRePrint(1).Value = True
End Sub

Private Sub optRePrint_Click(Index As Integer)
     txtZoom.Enabled = (Index = 1)
     If Index = 0 Then txtZoom.Text = frmPrintPreview.lngZoom
End Sub

Private Sub txtMargin_GotFocus(Index As Integer)
     txtMargin(Index).SelStart = 0
     txtMargin(Index).SelLength = Len(txtMargin(Index).Text)
End Sub
