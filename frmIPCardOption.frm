VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIPCardOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ת"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmIPCardOption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   3090
      TabIndex        =   7
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��ʼ(&B)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3090
      TabIndex        =   6
      Top             =   300
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   360
      Left            =   2790
      TabIndex        =   5
      Top             =   1575
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59506689
      CurrentDate     =   37424
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   1155
      TabIndex        =   4
      Top             =   1575
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59506689
      CurrentDate     =   37424
   End
   Begin VB.OptionButton optBetween 
      Caption         =   "�˼�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1605
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpBefore 
      Height          =   360
      Left            =   1155
      TabIndex        =   2
      Top             =   930
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59506689
      CurrentDate     =   37424
   End
   Begin VB.OptionButton optBefore 
      Caption         =   "��ǰ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   810
   End
   Begin VB.OptionButton optAll 
      Caption         =   "��תȫ��������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   315
      Value           =   -1  'True
      Width           =   2265
   End
End
Attribute VB_Name = "frmIPCardOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim rstIPCount As New ADODB.Recordset
Dim rstIPCountHistory As New ADODB.Recordset
Dim i As Integer
Dim lngrow As Long
Dim strSql As String
writeLog Me.Name, "Begin"
If optAll Then
   strSql = " * from IPCount"
End If
If optBefore Then
   strSql = " * from IPCount where  AddDate <=#" & dtpBefore.Value & "#"
End If
If optBetween Then
   MsgShow "�Բ��������ܲ��ô˷�ʽת����"
   Exit Sub
End If
Me.Hide
frmStatus.Show
rstIPCount.Open "select " & strSql, gCnn, adOpenStatic, adLockOptimistic
rstIPCountHistory.Open "select top 1 * from IPCountHistory", gCnn, adOpenStatic, adLockOptimistic
With rstIPCountHistory
While Not rstIPCount.EOF
 .AddNew
 For i = 1 To .Fields.Count - 2
  .Fields(i) = rstIPCount.Fields(i)
 Next i
 !ZZDate = Date
 frmStatus.ShowMsg "����ת��ʹ��ʹ����־���˺ţ�" & !CountNo
 .Update
 With rstIPCount
  !AddMoney = !AddMoney - !UsedMoney
  !UsedMoney = 0
  .Update
 End With
 rstIPCount.MoveNext
 lngrow = lngrow + 1
Wend
.Close
rstIPCount.Close
Unload frmStatus
MsgShow "ת����ϣ�"
Unload Me
End With
writeLog Me.Name, "End", "���ƣ�" & Str(lngrow)
End Sub
Private Sub Form_Load()
CenterForm Me
End Sub


