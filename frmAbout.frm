VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "정보"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Tag             =   "정보 일정표"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   480
      ScaleMode       =   0  '사용자
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "확인"
      Top             =   2625
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "시스템 정보..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "시스템 정보..."
      Top             =   3075
      Width           =   1452
   End
   Begin VB.Label lblDescription 
      Caption         =   "응용 프로그램 설명"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   6
      Tag             =   "응용 프로그램 설명"
      Top             =   1125
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "응용 프로그램 제목"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "응용 프로그램 제목"
      Top             =   240
      Width           =   4092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '내부 단색
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label lblVersion 
      Caption         =   "버전"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Tag             =   "버전"
      Top             =   780
      Width           =   4092
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "경고: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Tag             =   "경고: ..."
      Top             =   2625
      Visible         =   0   'False
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 레지스트리 키 보안 옵션...
Const KEY_ALL_ACCESS = &H2003F
                                          

' 레지스트리 키 ROOT 형식...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode null 종료 문자열
Const REG_DWORD = 4                      ' 32비트 숫자


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    'lblVersion.Caption = "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = "버전 3.0.0 베타 " & App.Revision
    lblTitle.Caption = App.Title ' & " " & App.Major
    Me.Caption = App.Title & " 정보"
    lblDescription.Caption = App.FileDescription
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr

        Dim rc As Long
        Dim SysInfoPath As String
        

        ' 레지스트리에서 시스템 정보 프로그램 경로\이름을 가져오기를 시도합니다...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' 레지스트리에서 시스템 정보 프로그램 경로만을 가져오기를 시도합니다...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' 알고 있는 32비트 파일 버전이 있는 것을 확인합니다.
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' 오류 - 파일을 찾을 수 없습니다...
                Else
                        GoTo SysInfoErr
                End If
        ' 오류 - 레지스트리 항목을 찾을 수 없습니다...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MessageBox "지금은 시스템 정보를 사용할 수 없습니다. 시스템에 MS Info 구성요소가 설치됐는지 확인하십시오. 없으면 다시 설치하십시오.", "시스템 정보", Me, 16
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' 루프 카운터
        Dim rc As Long                                          ' 반환 코드
        Dim hKey As Long                                        ' 열려 있는 레지스트리 키를 처리합니다.
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' 레지스트리 키의 데이터 형식
        Dim tmpVal As String                                    ' 레지스트리 키 값을 임시로 저장합니다.
        Dim KeyValSize As Long                                  ' 레지스트리 키 변수의 크기
        '------------------------------------------------------------
        ' KeyRoot 아래의 RegKey를 엽니다.{HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 레지스트리 키를 엽니다.
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다...
        

        tmpVal = String$(1024, 0)                             ' 변수 공간을 할당합니다.
        KeyValSize = 1024                                       ' 변수 크기를 표시합니다.
        

        '------------------------------------------------------------
        ' 레지스트리 키 값을 검색합니다...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' 키 값을 가져 오고 작성합니다.
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다.
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' 변환할 키 값 형식을 결정합니다...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' 데이터 형식을 검색합니다...
        Case REG_SZ                                             ' 문자열 레지스트리 키 데이터 형식
                KeyVal = tmpVal                                     ' 문자열 값을 복사합니다.
        Case REG_DWORD                                          ' 이진 워드 레지스트리 키 데이터 형식
                For i = Len(tmpVal) To 1 Step -1                    ' 각각의 비트를 변환합니다.
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 변수 문자를 문자별로 작성합니다.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' 이진 워드를 문자열로 변환합니다.
        End Select
        

        GetKeyValue = True                                      ' 성공을 반환합니다.
        rc = RegCloseKey(hKey)                                  ' 레지스트리 값을 닫습니다.
        Exit Function                                           ' 끝냅니다.
        

GetKeyError:    ' 발생한 오류를 처리합니다...
        KeyVal = ""                                             ' 반환값을 빈 문자열로 설정합니다.
        GetKeyValue = False                                     '실패를 반환합니다.
        rc = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
End Function

