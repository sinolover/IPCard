VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� �˳���Ϊ���԰� ���е��������� ������ַ�"
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
         Caption         =   "�뵥����ߵļ�����ע��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��л��"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2760
         TabIndex        =   9
         Top             =   690
         Width           =   2400
      End
      Begin VB.Label lblReg 
         Caption         =   "ע���룺"
         Height          =   240
         Left            =   2745
         TabIndex        =   8
         Top             =   345
         Width           =   2475
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "�汾��200212061218"
         Height          =   180
         Left            =   1005
         TabIndex        =   7
         Top             =   675
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Caption         =   "IP������ϵͳ"
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
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   4140
      TabIndex        =   0
      Top             =   2670
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
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

' ע���ؼ��ְ�ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��� ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' �����Ŀյ��ս��ַ���
Const REG_DWORD = 4                      ' 32λ����

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
    lblVersion.Caption = "�汾 " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    If blnRegstered Then
       cmdReg.Enabled = False
       lblUser.Caption = "�����ע�ᵽ " & strRegUser
       lblUser.Visible = True
       lblReg.Visible = True
       lblReg.Caption = GetReg("RegName", "��л����֧��") ' strRegCorName
       lblNoReg.Visible = False
       Me.Caption = "IP������ϵͳ " ' & App.Title"
      Else
       cmdReg.Enabled = True
       lblUser.Visible = False
       lblReg.Visible = False
       lblNoReg.Visible = True
       Me.Caption = "IP������ϵͳ ���ð�" ' & App.Title"
       MsgShow "�˳���Ϊ���԰� ���е��������� ������ַ���"
    End If
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע����л��ϵͳ��Ϣ�����·��������...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ��ͼ����ע����л��ϵͳ��Ϣ�����·��...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' ��֪32λ�ļ��汾����Чλ��
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���� - �ļ����ܱ��ҵ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ע�����Ӧ��Ŀ���ܱ��ҵ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgShow "��ʱϵͳ��Ϣ������"
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �򿪵�ע���ؼ��־��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ���ֵ����ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��Ա����ĳߴ�
    '------------------------------------------------------------
    ' �� {HKEY_LOCAL_MACHINE...} �µ� RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ���ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ��ӳ�����ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null ���ҵ�,���ַ����з������
    Else                                                    ' WinNT û�п��ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null û�б��ҵ�, �����ַ���
    End If
    '------------------------------------------------------------
    ' ����ת���Ĺؼ��ֵ�ֵ����...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������������...
    Case REG_SZ                                             ' �ַ���ע��ؼ�����������
        KeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽڵ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ��ÿλ����ת��
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' ����ֵ�ַ��� By Char��
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ת�����ֽڵ��ַ�Ϊ�ַ���
    End Select
    
    GetKeyValue = True                                      ' ���سɹ�
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:      ' �������������...
    KeyVal = ""                                             ' ���÷���ֵ�����ַ���
    GetKeyValue = False                                     ' ����ʧ��
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function

Private Sub lblMail_Click()
  Call ShellExecute(Me.hWnd, "open", "mailto:sllin@cmmail.com", "", "", 5)
End Sub

