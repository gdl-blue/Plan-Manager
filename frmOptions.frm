VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환경설정"
   ClientHeight    =   4635
   ClientLeft      =   -75
   ClientTop       =   -75
   ClientWidth     =   8250
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOptionHelp 
      Caption         =   "도움말(&H)..."
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      TabCaption(0)   =   "화면 표시"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "사용자 데이터"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "표준"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "검사"
      TabPicture(3)   =   "frmOptions.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Label9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "사용자 분류"
      TabPicture(4)   =   "frmOptions.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdClearCates"
      Tab(4).Control(1)=   "cmdDelSelCate"
      Tab(4).Control(2)=   "cmdAddNewCate"
      Tab(4).Control(3)=   "txtCategory"
      Tab(4).Control(4)=   "Frame5"
      Tab(4).Control(5)=   "Label8"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "테마"
      TabPicture(5)   =   "frmOptions.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "초기화"
      TabPicture(6)   =   "frmOptions.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame3"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "달력"
      TabPicture(7)   =   "frmOptions.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame10 
         Caption         =   "테마"
         Height          =   735
         Left            =   120
         TabIndex        =   44
         Top             =   3240
         Width           =   7455
         Begin VB.CommandButton cmdTheSet 
            Caption         =   "테마(&T)..."
            Height          =   375
            Left            =   6120
            TabIndex        =   45
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "테마를 적용하려면 다음 단추를 누르십시오."
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   4215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "달력"
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   7455
         Begin VB.CommandButton cmdCalSet 
            Caption         =   "달력(&A)..."
            Height          =   375
            Left            =   6120
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "달력을 사용자 정의하려면 다음 단추를 누르십시오."
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   300
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "달력"
         Height          =   975
         Left            =   -74880
         TabIndex        =   38
         Top             =   720
         Width           =   6015
         Begin VB.ComboBox cmbWSD 
            Height          =   300
            Left            =   120
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label5 
            Caption         =   "주의 시작 요일:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "레이아웃"
         Height          =   1575
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   5895
         Begin VB.CheckBox chkNoRibbon 
            Caption         =   "리본 메뉴 비활성(&N)"
            Height          =   255
            Left            =   3360
            TabIndex        =   48
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkTP 
            Caption         =   "오늘의일정 숨기기(&O)"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "초기화"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   6015
         Begin VB.CommandButton cmdPrgReset 
            Caption         =   "초기화(&R)"
            Height          =   375
            Left            =   4560
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblResetN2 
            Caption         =   "단계 전입니다."
            Height          =   255
            Left            =   1440
            TabIndex        =   32
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "프로그램 전체 데이터를 초기화합니다."
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblResetCount 
            Caption         =   "7"
            Height          =   255
            Left            =   1320
            TabIndex        =   34
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblResetN1 
            Caption         =   "데이터 초기화"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "색 테마"
         Height          =   975
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   5775
         Begin VB.ComboBox cmbBGColor 
            Height          =   300
            Left            =   120
            Style           =   2  '드롭다운 목록
            TabIndex        =   29
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label Label10 
            Caption         =   "프로그램 배경 테마:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.CommandButton cmdClearCates 
         Caption         =   "모두 삭제(&C)"
         Height          =   375
         Left            =   -68760
         TabIndex        =   26
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "값 검사 설정"
         Height          =   615
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   6015
         Begin VB.CheckBox chkNoTimeCHeck 
            Caption         =   "일정 추가 시 시간이 올바르지 검사 안함(&T)"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.CommandButton cmdDelSelCate 
         Caption         =   "선택 분류 삭제"
         Height          =   375
         Left            =   -68760
         TabIndex        =   22
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddNewCate 
         Caption         =   "입력 분류 추가"
         Height          =   375
         Left            =   -68760
         TabIndex        =   21
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtCategory 
         Height          =   270
         Left            =   -72960
         TabIndex        =   20
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Caption         =   "분류 목록"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   1815
         Begin VB.FileListBox lvCustomCates 
            Height          =   2970
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "시작"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Width           =   6015
         Begin VB.OptionButton radCFQ 
            Caption         =   "이전에 끝낸 지점에서 시작(&Q)"
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   1200
            Width           =   5295
         End
         Begin VB.OptionButton radSelST 
            Caption         =   "화면 선택(&T)"
            Height          =   255
            Left            =   360
            TabIndex        =   49
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox cmbStartPage 
            Height          =   300
            Left            =   600
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label Label7 
            Caption         =   "시작 화면:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "내 데이터"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   6015
         Begin VB.FileListBox lvTaskFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   9
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.FileListBox lvContactFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   8
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.FileListBox lvPlanFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelTasks 
            Caption         =   "모두 삭제(&L)"
            Height          =   375
            Left            =   4560
            TabIndex        =   6
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelContacts 
            Caption         =   "모두 삭제(&E)"
            Height          =   375
            Left            =   4560
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelPlans 
            Caption         =   "모두 삭제(&D)"
            Height          =   375
            Left            =   4560
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "내 작업목록:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "내 주소록:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "내 일정:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Label Label9 
         Caption         =   "[*] 이 설정을 변경하면 프로그램이 올바로 작동하지 않을 수 있습니다."
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   3720
         Width           =   7335
      End
      Begin VB.Label Label8 
         Caption         =   "새 분류 추가:"
         Height          =   255
         Left            =   -72960
         TabIndex        =   19
         Top             =   3480
         Width           =   2055
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "이 설정을 완전히 적용하려면 프로그램을 다시 시작해야 합니다."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   4320
      Width           =   5295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ResetCount As Integer

Private Sub cmdAddNewCate_Click()
    If txtCategory.Text <> "업무" And txtCategory.Text <> "여가생활" And txtCategory.Text <> "약속" And txtCategory.Text <> "취미" And txtCategory.Text <> "(지정되지 않음)" Then
        On Error Resume Next
        If Len(txtCategory.Text) < 1 Then
            MessageBox "분류의 제목을 입력해주십시오.", "오류", Me, 16
            Exit Sub
        End If
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        Dim iFileNo As Integer
        iFileNo = FreeFile
        '파일을 연다.
        
        Open "C:\CALPLANS\CTGORIES\" & txtCategory.Text For Output As #iFileNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFileNo, ""
        
        '파일을 닫는다.
        Close #iFileNo
        
        lvCustomCates.Refresh
        
        MessageBox "추가되었습니다.", "성공", Me
    Else
        MessageBox "이미 존재하거나 올바르지 않습니다.", "오류", Me, 16
    End If
End Sub

Private Sub cmdCalSet_Click()
    SSTab1.Tab = 7
End Sub

Private Sub cmdClearCates_Click()
    If Confirm("정말로 " & lvCustomCates.ListCount & "개의 분류를 *모두* 삭제하시겠습니까?", "삭제", Me) Then
        On Error Resume Next
        Dim i As Integer
        For i = 0 To lvCustomCates.ListCount - 1
            Kill "C:\CALPLANS\CTGORIES\" & lvCustomCates.List(i)
        Next i
        
        lvCustomCates.Refresh
        MessageBox "모두 삭제되었습니다.", "성공", Me, 48
    End If
End Sub

Sub cmdDelContacts_Click()
    If Confirm("정말로 삭제하시겠습니까?", "삭제", Me) Then
        If Confirm("복구 *불가능*합니다. 정말로 모든 주소록을 삭제하시겠습니까?", "삭제", Me, , True) Then
            On Error Resume Next
            lvContactFiles.Path = "C:\CALPLANS\CONTACTS"
            
            Dim Contact As Integer
            Dim ContactName As String
            For Contact = 0 To lvContactFiles.ListCount - 1
                ContactName = lvContactFiles.List(Contact)
                Kill "C:\CALPLANS\CONTACTS\" & ContactName
                DeleteSetting "Calendar", "Contacts", ContactName & "OtherNum"
                DeleteSetting "Calendar", "Contacts", ContactName & "Postal"
                DeleteSetting "Calendar", "Contacts", ContactName & "Home"
                DeleteSetting "Calendar", "Contacts", ContactName & "Fax"
                DeleteSetting "Calendar", "Contacts", ContactName & "Email"
                DeleteSetting "Calendar", "Contacts", ContactName & "Content"
                DeleteSetting "Calendar", "Contacts", ContactName & "Company"
                DeleteSetting "Calendar", "Contacts", ContactName & "CellPhone"
                DeleteSetting "Calendar", "Contacts", ContactName & "Addr"
            Next Contact
            
            frmMain.LoadContacts
            
            MessageBox "주소록 데이타가 모두 삭제됐습니다.", "성공", Me, 64
        End If
    End If
End Sub

Private Sub cmdDelPlans_Click()
    Dim DelYear As String
    DelYear = InputBox("삭제할 연도를 입력하십시오.", "일정 모두 지우기")
    If DelYear <> "" Then
        If IsNumeric(DelYear) = False Then
            MessageBox "연도의 값이 올바르지 않습니다.", "연도", Me, 16
            Exit Sub
        End If
    
        On Error Resume Next
        If Confirm("정말로 삭제하시겠습니까?", "삭제", Me) Then
            If Confirm("복구 *불가능*합니다. 정말로 " & DelYear & "년의 모든 일정을 삭제하시겠습니까?", "삭제", Me, , True) Then
                On Error Resume Next
                Shell "CMD /C RD /S /Q " & ChrW$(34) & "C:\CALPLANS\" & DelYear & ChrW$(34)
                Shell "COMMAND /C DELTREE /Y " & ChrW$(34) & "C:\CALPLANS\" & DelYear & ChrW$(34)
            End If
        End If
    End If
End Sub

Private Sub cmdDelSelCate_Click()
    If Confirm("한번만 경고합니다. 정말로 삭제하시겠습니까?", "삭제", Me) Then
        On Error Resume Next
        Kill "C:\CALPLANS\CTGORIES\" & lvCustomCates.List(lvCustomCates.ListIndex)
        
        lvCustomCates.Refresh
    End If
End Sub

Sub cmdDelTasks_Click()
    If Confirm("정말로 삭제하시겠습니까?", "삭제", Me) Then
        If Confirm("복구 *불가능*합니다. 정말로 모든 작업을 삭제하시겠습니까?", "삭제", Me, , True) Then
            On Error Resume Next
            lvTaskFiles.Path = "C:\CALPLANS\TASKS"
            
            Dim Plan As Integer
            Dim TaskName As String
            For Plan = 0 To lvTaskFiles.ListCount - 1
                TaskName = lvTaskFiles.List(Plan)
                Kill "C:\CALPLANS\TASKS\" & TaskName
                DeleteSetting "Calendar", "Tasks", TaskName & "Perc"
                DeleteSetting "Calendar", "Tasks", TaskName & "Memo"
            Next Plan
            
            frmMain.LoadContacts
            
            MessageBox "작업목록 데이타가 모두 삭제됐습니다.", "성공", Me
        End If
    End If
    
    frmMain.LoadTasks
End Sub

Private Sub cmdOptionHelp_Click()
    MessageBox "설정과 관련된 도움말이 없습니다.", "오류", Me, 16
End Sub

Private Sub cmdTheSet_Click()
    SSTab1.Tab = 5
End Sub

Private Sub Command1_Click()
    'SaveSetting "Calendar", "Options", "NoResize", chkNoResize.Value
    SaveSetting "Calendar", "Options", "WSD", cmbWSD.ListIndex
    
    SaveSetting "Calendar", "Options", "StartPage", cmbStartPage.ListIndex
    
    SaveSetting "Calendar", "Options", "NoTimeCheck", chkNoTimeCHeck.Value
    
    SaveSetting "Calendar", "Options", "BGColor", cmbBGColor.ListIndex
    
    SaveSetting "Calendar", "Options", "TP", chkTP.Value
    
    SaveSetting "Calendar", "Options", "NoRibbon", chkNoRibbon.Value
    
    If radSelST.Value = False Then
        SaveSetting "Calendar", "Options", "SST", False
    Else
        SaveSetting "Calendar", "Options", "SST", True
    End If
    
    If GetSetting("Calendar", "Options", "TP", 0) = 1 Then
        frmMain.Width = 8715
    Else
        frmMain.Width = 11040
    End If
    frmMain.SetColor
    
    frmMain.MonthView1.StartOfWeek = cmbWSD.ListIndex + 1
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub cmdPrgReset_Click()
    ResetCount = ResetCount - 1
    lblResetCount.Caption = ResetCount
    If ResetCount = 6 Then
        lblResetN1.Visible = True
        lblResetN2.Visible = True
        lblResetCount.Visible = True
    End If
    
    If ResetCount = 0 Then
        cmdPrgReset.Enabled = False
        If Confirm("마지막 경고. 정말로 프로그램 전체를 초기화하시겠습니까?", "초기화", Me) Then
            Shell "CMD /C RD /S /Q C:\CALPLANS"
            Shell "CMD /C RD /S /Q C:\CALPLANS"
            MessageBox "초기화 완료. 프로그램을 종료합니다...", "초기화", Me
            End
        Else
            cmdPrgReset.Enabled = True
            lblResetN1.Visible = False
            lblResetN2.Visible = False
            lblResetCount.Visible = False
            
            ResetCount = 7
        End If
    End If
End Sub

Private Sub Form_Load()
    ResetCount = 7
    'chkNoResize.Value = GetSetting("Calendar", "Options", "NoResize", "0")
    
    chkNoTimeCHeck.Value = GetSetting("Calendar", "Options", "NoTimeCheck", 0)
    
    chkTP.Value = GetSetting("Calendar", "Options", "TP", 0)
    
    If GetSetting("Calendar", "Options", "SST", True) = True Then
        radSelST.Value = True
    Else
        radCFQ.Value = True
    End If
    
    chkNoRibbon.Value = GetSetting("Calendar", "Options", "NoRibbon", 0)
    
    
    On Error Resume Next
    cmbWSD.AddItem "일요일"
    cmbWSD.AddItem "월요일"
    cmbWSD.AddItem "화요일"
    cmbWSD.AddItem "수요일"
    cmbWSD.AddItem "목요일"
    cmbWSD.AddItem "금요일"
    cmbWSD.AddItem "토요일"
    
    cmbStartPage.AddItem "일정"
    cmbStartPage.AddItem "주소록"
    cmbStartPage.AddItem "할 일"
    cmbStartPage.AddItem "일과표"
    
    cmbBGColor.AddItem "시스템: 응용프로그램 작업영역"
    cmbBGColor.AddItem "시스템: 단추 표면색"
    cmbBGColor.AddItem "빨강"
    cmbBGColor.AddItem "노랑"
    cmbBGColor.AddItem "초록"
    cmbBGColor.AddItem "옥색"
    cmbBGColor.AddItem "청록"
    cmbBGColor.AddItem "파랑"
    cmbBGColor.AddItem "검정"
    
    cmbBGColor.ListIndex = GetSetting("Calendar", "Options", "BGColor", 0)
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CTGORIES"
    
    lvCustomCates.Path = "C:\CALPLANS\CTGORIES"
    
    cmbStartPage.ListIndex = GetSetting("Calendar", "Options", "StartPage", 0)
    
    cmbWSD.ListIndex = GetSetting("Calendar", "Options", "WSD", 0)
End Sub

Private Sub radCFQ_Click()
    cmbStartPage.Enabled = False
End Sub

Private Sub radSelST_Click()
    cmbStartPage.Enabled = True
End Sub
