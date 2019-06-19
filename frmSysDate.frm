VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统日期"
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker txtSysDate 
      Height          =   360
      Left            =   1755
      TabIndex        =   4
      Top             =   705
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59703297
      CurrentDate     =   37424
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消(C)"
      Height          =   400
      Left            =   3885
      TabIndex        =   1
      Top             =   615
      Width           =   1200
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3885
      TabIndex        =   0
      Top             =   135
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "当前系统日期："
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
      Left            =   225
      TabIndex        =   5
      Top             =   225
      Width           =   1470
   End
   Begin VB.Label lblDate 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1755
      TabIndex        =   3
      Top             =   195
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "输入系统日期："
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
      Left            =   225
      TabIndex        =   2
      Top             =   795
      Width           =   1470
   End
End
Attribute VB_Name = "frmSysDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
writeLog Me.Name, "Cancel", "未修改系统日期"
Unload Me
End Sub
Private Sub Form_Activate()
txtSysDate.SetFocus
End Sub
Private Sub Form_Load()
writeLog Me.Name
CenterForm Me
Me.Icon = LoadResPicture(103, 1)
lblDate = CStr(gSysDate)
txtSysDate = CStr(gSysDate)
End Sub
Private Sub OKButton_Click()
If IsDate(txtSysDate) Then gSysDate = CDate(txtSysDate)
mfrmMain.sta.Panels(2).Text = "系统日期：" & CStr(gSysDate)
writeLog Me.Name, "Edit", mfrmMain.sta.Panels(2).Text
Unload Me
End Sub
