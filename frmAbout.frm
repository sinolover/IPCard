VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 此程序为测试版 不承担连带责任 并请勿分发"
   ClientHeight    =   3645
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2515.844
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   210
      ScaleHeight     =   1065
      ScaleWidth      =   5190
      TabIndex        =   5
      Top             =   105
      Width           =   5250
      Begin VB.CommandButton cmdReg 
         DisabledPicture =   "frmAbout.frx":0000
         Height          =   615
         Left            =   135
         Picture         =   "frmAbout.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   780
      End
      Begin VB.Label lblNoReg 
         Caption         =   "请单击左边的加密锁注册"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2220
         TabIndex        =   11
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label lblUser 
         Caption         =   "感谢："
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2760
         TabIndex        =   9
         Top             =   690
         Width           =   2400
      End
      Begin VB.Label lblReg 
         Caption         =   "注册码："
         Height          =   240
         Left            =   2745
         TabIndex        =   8
         Top             =   345
         Width           =   2475
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "版本：200212061218"
         Height          =   180
         Left            =   1005
         TabIndex        =   7
         Top             =   675
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Caption         =   "IP卡管理系统"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1005
         TabIndex        =   6
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   4140
      TabIndex        =   0
      Top             =   2670
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   4140
      TabIndex        =   1
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Label lblMail 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmAbout.frx":0D0C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3660
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1749.702
      Y2              =   1749.702
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":15D6
      ForeColor       =   &H00FF0000&
      Height          =   1140
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   5505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1760.055
      Y2              =   1760.055
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":1711
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   255
      TabIndex        =   3
      Top             =   2670
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'      Height = 2580
'      Width = 3630
'   End
'End
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 独立的空的终结字符串
Const REG_DWORD = 4                      ' 32位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdReg_Click()
Load frmReg
frmReg.Show vbModal
End Sub
Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
    CenterForm Me
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    If blnRegstered Then
       cmdReg.Enabled = False
       lblUser.Caption = "本软件注册到 " & strRegUser
       lblUser.Visible = True
       lblReg.Visible = True
       lblReg.Caption = GetReg("RegName", "感谢您的支持") ' strRegCorName
       lblNoReg.Visible = False
       Me.Caption = "IP卡管理系统 " ' & App.Title"
      Else
       cmdReg.Enabled = True
       lblUser.Visible = False
       lblReg.Visible = False
       lblNoReg.Visible = True
       Me.Caption = "IP卡管理系统 试用版" ' & App.Title"
       MsgShow "此程序为测试版 不承担连带责任 并请勿分发！"
    End If
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表中获得系统信息程序的路径及名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 试图仅从注册表中获得系统信息程序的路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 已知32位文件版本的有效位置
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件不能被找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册表相应条目不能被找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgShow "此时系统信息不可用"
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册表关键字句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字值的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键自变量的尺寸
    '------------------------------------------------------------
    ' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 外接程序空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 被找到,从字符串中分离出来
    Else                                                    ' WinNT 没有空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 没有被找到, 分离字符串
    End If
    '------------------------------------------------------------
    ' 决定转换的关键字的值类型...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串注册关键字数据类型
        KeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节的注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 将每位进行转换
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 生成值字符。 By Char。
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换四字节的字符为字符串
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:      ' 错误发生后将其清除...
    KeyVal = ""                                             ' 设置返回值到空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Private Sub lblMail_Click()
  Call ShellExecute(Me.hWnd, "open", "mailto:sllin@cmmail.com", "", "", 5)
End Sub

