VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������ע����Ϣ"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "����(&H)"
      Height          =   400
      Left            =   3315
      TabIndex        =   10
      Top             =   1530
      Width           =   1200
   End
   Begin VB.TextBox txtRegOld 
      Height          =   360
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1195
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   3315
      TabIndex        =   5
      Top             =   900
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ע��(&R)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3315
      TabIndex        =   4
      Top             =   270
      Width           =   1200
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1140
      TabIndex        =   3
      Top             =   1710
      Width           =   1815
   End
   Begin VB.TextBox txtRegID 
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
      Left            =   1140
      TabIndex        =   1
      Top             =   680
      Width           =   1815
   End
   Begin VB.TextBox txtRegName 
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
      Left            =   1140
      TabIndex        =   0
      Top             =   165
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ע����"
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
      Left            =   285
      TabIndex        =   9
      Top             =   1305
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
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
      Left            =   285
      TabIndex        =   8
      Top             =   1815
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ע���"
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
      Left            =   285
      TabIndex        =   7
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ע����"
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
      Left            =   285
      TabIndex        =   6
      Top             =   270
      Width           =   630
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim rstReg As New ADODB.Recordset
Dim strRegID As String
If Len(Trim(txtRegName.Text)) < 1 Then
   MsgShow "������ע�����ƣ�"
   txtRegName.SetFocus
   Exit Sub
End If
If Len(Trim(txtRegID.Text)) < 1 Then
   MsgShow "������ע��ţ�"
   txtRegID.SetFocus
   Exit Sub
End If
If Len(Trim(txtUserName.Text)) < 1 Then
   MsgShow "������ע���û����ƣ�"
   txtUserName.SetFocus
   Exit Sub
End If
With rstReg
 If .State <> 0 Then .Close
 .Open "select * from RegID", gCnn, adOpenStatic, adLockOptimistic
 If Not .EOF Then
    !RegCorName = txtUserName.Text
    !RegUser = txtRegName.Text
    !RegNumber = IIf(Len(Trim(txtRegID.Text)) < 13, Trim(txtRegID.Text), Left(Trim(txtRegID.Text), 12))
   .Update
 End If
End With
MsgShow "ע����ɣ����ڽ��˳�ϵͳ�������½��룡"
Unload Me
Unload frmAbout
End
'SubCipher strRegID, txtRegName, "WinWay200206"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
txtRegOld.Text = Trim(SubCipher(gstrKeyTMP, IDVerify_new))
End Sub
