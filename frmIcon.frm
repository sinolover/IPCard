VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   0  'None
   ClientHeight    =   1530
   ClientLeft      =   1515
   ClientTop       =   1710
   ClientWidth     =   2835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1530
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuPopup 
      Caption         =   "SysTray Popup Menu"
      Begin VB.Menu mnuAbout 
         Caption         =   "&A 关于..."
      End
      Begin VB.Menu mnuAlertMoney 
         Caption         =   "&M 余额警戒"
      End
      Begin VB.Menu mnuAlertDate 
         Caption         =   "&D 使用警戒"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&H 帮助"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&E 退出"
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'Add the icon to the system tray...
ConnectDB
With nfIconData
 .hWnd = Me.hWnd
 .uID = Me.Icon
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.Handle
 .szTip = "IP卡管理系统 一级警员为您服务" & Chr$(0)
 .cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If frmAbout.Visible And X = 7740 Then frmAbout.Hide
If IsNull(frmAlertDate) Then MsgBox " asd"
If blnFrmOpened And X = 7740 Then frmAlertMoney.Hide
If blnFrmOpened And X = 7740 Then frmAlertDate.Hide
Select Case X
Case 7680 'MouseMove

Case 7695 'LeftMouseDown
 'frmAbout.Show
 'frmAlertMoney.Show
  PopupMenu mnuPopup, 0, , , mnuClose


Case 7710 'LeftMouseUp
 
Case 7725 'LeftDblClick
' Load frmAlertMoney
' frmAlertMoney.Show
 PopupMenu mnuPopup, 0, , , mnuClose

Case 7740 'RightMouseDown
 PopupMenu mnuPopup, 0, , , mnuClose

Case 7755 'RightMouseUp

Case 7770 'RightDblClick

End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub mnuAbout_Click()
MsgBox "IP 卡帐务管理系统警戒员，忠诚为您服务！"
End Sub
Private Sub mnuAlertDate_Click()
strOprName = ""
Load frmLoginTray
frmLoginTray.Show vbModal
If Len(Trim(strOprName)) > 0 Then frmAlertDate.Show
End Sub

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub

Private Sub mnuAlertMoney_click()
strOprName = ""
Load frmLoginTray
frmLoginTray.Show vbModal
If Len(Trim(strOprName)) > 0 Then frmAlertMoney.Show
End Sub
Private Sub mnuClose_Click()
If MsgBox("您撤销警戒员吗？", vbOKCancel) = vbOK Then
   Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   Unload Me
End If
End Sub

Private Sub ConnectDB()
Dim strConnectionString As String
On Error Resume Next
strConnectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" _
                       & App.Path & "\IPCard.mdb" & _
                      ";uid=Admin;pwd=iamchinese"
'strConnectionString = "DSN=IPCard;uid=Admin;pwd=iamchinese"
With gCnn
 If .State <> 0 Then .Close
 .Open strConnectionString
 If .State <> adStateOpen Then
  MsgBox "无法打开数据库，请重新安装系统"
  Unload Me
  End
 End If
End With
End Sub

