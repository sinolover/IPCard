VERSION 5.00
Begin VB.Form frmLogo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎使用"
      BeginProperty Font 
         Name            =   "方正舒体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   930
      Left            =   270
      TabIndex        =   2
      Top             =   225
      Width           =   4590
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP 卡帐务管理系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   285
      TabIndex        =   1
      Top             =   1455
      Width           =   4650
   End
   Begin VB.Label lblShowMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注  意：此程序为测试版 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   360
      Left            =   705
      TabIndex        =   0
      Top             =   2775
      Width           =   4140
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'Dim sngBegin As Single
'BorderStyle = 3
'ControlBox = False
With lblShowMessage
 .Caption = GetReg("RegName", "用户")
 .Left = Int((Me.Width - .Width) / 2)
End With
'lblUserName.Left = Int((Me.Width - lblUserName.Width) / 2)
Picture = LoadResPicture(101, vbResBitmap)
Width = ScaleX(Picture.Width, vbHimetric, vbTwips) + (Width - ScaleWidth)
'Height = ScaleY(Picture.Height, vbHimetric, vbTwips) + (Height - ScaleHeight)
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
Show
'sngBegin = Timer
'Do While (Timer - sngBegin + 86400) Mod 86400 < 2
DoEvents
'Loop
'Unload Me
End Sub

