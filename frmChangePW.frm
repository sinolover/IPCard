VERSION 5.00
Begin VB.Form frmChangePW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����޸�"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmChangePW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(C)"
      Height          =   400
      Left            =   3330
      TabIndex        =   4
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3330
      TabIndex        =   3
      Top             =   330
      Width           =   1200
   End
   Begin VB.TextBox txtAgain 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1305
      Width           =   1620
   End
   Begin VB.TextBox txtNewPW 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   750
      Width           =   1620
   End
   Begin VB.TextBox txtOldPW 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   195
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   7
      Top             =   1380
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   6
      Top             =   825
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ԭ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   5
      Top             =   270
      Width           =   1050
   End
End
Attribute VB_Name = "frmChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim rstOpr As New ADODB.Recordset
Dim strErrDescription As String

On Error GoTo Err_Exit
gCnn.BeginTrans
If Trim(txtNewPW.Text) <> Trim(txtAgain.Text) Then
   txtNewPW.SetFocus
   strErrDescription = "���벻ƥ�䣬���������룡"
   Err.Raise lngErrNumberCommon, "me"
   Exit Sub
End If
With rstOpr
 If .State <> 0 Then .Close
 .Open "select * from operator where oprno='" & Trim(gOpr.strOprNo) & "'", gCnn, adOpenStatic, adLockOptimistic
 If Trim(txtOldPW.Text) <> Trim(!OprPW) Then
    txtOldPW.SetFocus
    strErrDescription = "���벻��ȷ�����������룡"
    Err.Raise lngErrNumberCommon, "ChangePassWord"
 End If
 !OprPW = Trim(txtNewPW.Text) & ""
 .Update
 .Close
End With
gCnn.CommitTrans
Me.Hide
MsgShow "�����޸ĳɹ����ס�����룡"
Unload Me
Exit Sub
Err_Exit:
On Error Resume Next
gCnn.RollbackTrans
Me.Hide
MsgShow "�����޸�ʧ�ܣ�" & vbCrLf & "������Ϣ��" & Err.Description & strErrDescription
Me.Show vbModal
End Sub
Private Sub Form_Load()
Me.Left = 1200
Me.Top = 1200
Me.Caption = "�����޸�--" & gOpr.strOprName
End Sub
