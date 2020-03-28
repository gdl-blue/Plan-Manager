VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "알고 계십니까"
   ClientHeight    =   3390
   ClientLeft      =   2295
   ClientTop       =   2325
   ClientWidth     =   5415
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "시작 시 표시(&S)"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "다음 팁(&N)"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":000C
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "알고 계십니까.."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 메모리에 있는 팁 데이터베이스
Dim Tips As New Collection

' 팁 파일의 이름
Const TIP_FILE = "TIPOFDAY.TXT"

' 현재 표시되어 있는 팁 컬렉션의 인덱스
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' 팁을 임의로 선택합니다.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' 또는, 팁을 순서대로 표시할 수 있습니다.

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' 팁을 표시합니다.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' 각 팁은 파일에서 읽어 옵니다.
    Dim InFile As Integer   ' 파일에 대한 설명자
    
    ' 사용할 수 있는 다음 파일 설명자를 가져옵니다.
    InFile = FreeFile
    
    ' 파일을 지정하였는지 확인합니다.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 파일을 열기 전에 파일이 있는지 확인합니다.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 텍스트 파일에서 컬렉션을 읽습니다.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' 팁을 임의의 순서대로 표시합니다.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' 시작 시 이 폼의 표시 여부를 저장합니다.
    SaveSetting App.EXEName, "Options", "시작 시 이 화면 표시", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    ' 시작 시 표시할 것인지를 확인합니다.
    ShowAtStartup = GetSetting(App.EXEName, "Options", "시작 시 이 화면 표시", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    ' 확인란을 설정합니다. 설정하면 값을 레지스트리에 다시 기록하게 됩니다.
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Rnd를 시작합니다.
    Randomize
    
    ' 팁 파일을 읽어서 임의의 팁을 표시합니다.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = TIP_FILE & " 파일을 찾지 못했습니까? " & vbCrLf & vbCrLf & _
           TIP_FILE & " 파일을 [메모장]을 사용하여 한 줄에 한 팁씩 작성한 후 " & _
           "해당 응용 프로그램이 있는 디렉터리에 복사하십시오."
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
