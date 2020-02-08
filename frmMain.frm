VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  '단일 고정
   Caption         =   "일정관리자"
   ClientHeight    =   6645
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H8000000C&
      Caption         =   "도움말 ▼"
      Height          =   360
      Left            =   7440
      TabIndex        =   52
      Top             =   75
      Width           =   1095
   End
   Begin TabDlg.SSTab ssRibbonMenu 
      Height          =   1335
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   1940
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      MouseIcon       =   "frmMain.frx":0442
      TabCaption(0)   =   " 홈"
      TabPicture(0)   =   "frmMain.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPlanList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPlanIndex"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEndPrg"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " 보기"
      TabPicture(1)   =   "frmMain.frx":08B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tglStatusBar"
      Tab(1).Control(1)=   "cmdOptions"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdEndPrg 
         Caption         =   "끝내기"
         Height          =   855
         Left            =   2760
         Picture         =   "frmMain.frx":0D02
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlanIndex 
         Caption         =   "데이터 색인"
         Height          =   855
         Left            =   1200
         Picture         =   "frmMain.frx":1144
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlanList 
         Caption         =   "일정 목록"
         Height          =   855
         Left            =   120
         Picture         =   "frmMain.frx":1586
         Style           =   1  '그래픽
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "옵션"
         Height          =   855
         Left            =   -73680
         Picture         =   "frmMain.frx":19C8
         Style           =   1  '그래픽
         TabIndex        =   48
         Top             =   360
         Width           =   1095
      End
      Begin MSForms.ToggleButton tglStatusBar 
         Height          =   855
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1931;1508"
         Value           =   "1"
         Caption         =   "상태표시줄"
         Picture         =   "frmMain.frx":1E0A
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      TabCaption(0)   =   "일정"
      TabPicture(0)   =   "frmMain.frx":225C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MonthView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Dir1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "주소록"
      TabPicture(1)   =   "frmMain.frx":26AE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvContacts"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdSaveContact"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "lvContactFiles"
      Tab(1).Control(6)=   "cmdDelContact"
      Tab(1).Control(7)=   "cmdDeleteAllContacts"
      Tab(1).Control(8)=   "cmdResetFields"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "할 일"
      TabPicture(2)   =   "frmMain.frx":2B00
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvTasks"
      Tab(2).Control(1)=   "cmdSaveTask"
      Tab(2).Control(2)=   "cmdDelTask"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "lvTaskFiles"
      Tab(2).Control(5)=   "cmdDeleteAllTasks"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdDeleteAllTasks 
         Caption         =   "모두 삭제(&L)"
         Height          =   495
         Left            =   -67920
         TabIndex        =   53
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdResetFields 
         Caption         =   "내용 초기화(&R)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   46
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeleteAllContacts 
         Caption         =   "모두 삭제(&E)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   45
         Top             =   2520
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   300
         Left            =   7440
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.FileListBox lvTaskFiles 
         Height          =   270
         Left            =   -72960
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "할 일 정보"
         Height          =   4095
         Left            =   -72480
         TabIndex        =   28
         Top             =   120
         Width           =   4455
         Begin ComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   3840
            TabIndex        =   37
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtPercentage"
            BuddyDispid     =   196622
            OrigLeft        =   3850
            OrigTop         =   1200
            OrigRight       =   4105
            OrigBottom      =   1455
            Increment       =   10
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMemo 
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   36
            Top             =   1920
            Width           =   4215
         End
         Begin VB.TextBox txtTaskTitle 
            Height          =   270
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtPercentage 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   3450
            TabIndex        =   32
            Text            =   "0"
            Top             =   1200
            Width           =   420
         End
         Begin ComctlLib.ProgressBar TaskProgress 
            Height          =   300
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label Label11 
            Caption         =   "메모:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "제목:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "%"
            Height          =   255
            Left            =   4155
            TabIndex        =   31
            Top             =   1245
            Width           =   135
         End
         Begin VB.Label Label8 
            Caption         =   "완료율:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdDelTask 
         Caption         =   "삭제(&D)"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -67920
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveTask 
         Caption         =   "저장(&S)"
         Height          =   495
         Left            =   -67920
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.ListBox lvTasks 
         Height          =   4050
         ItemData        =   "frmMain.frx":2F52
         Left            =   -74880
         List            =   "frmMain.frx":2F59
         Style           =   1  '확인란
         TabIndex        =   25
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelContact 
         Caption         =   "삭제(&D)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.FileListBox lvContactFiles 
         Height          =   270
         Left            =   -69240
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "메모"
         Height          =   1575
         Left            =   -73080
         TabIndex        =   10
         Top             =   2640
         Width           =   4935
         Begin VB.TextBox txtContent 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   22
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.CommandButton cmdSaveContact 
         Caption         =   "저장(&S)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "전화번호"
         Height          =   975
         Left            =   -73080
         TabIndex        =   8
         Top             =   1560
         Width           =   4935
         Begin VB.TextBox txtOtherNumber 
            Height          =   270
            Left            =   2880
            TabIndex        =   21
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtFax 
            Height          =   270
            Left            =   600
            TabIndex        =   19
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtHome 
            Height          =   270
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtCompany 
            Height          =   270
            Left            =   3000
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "기타:"
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "팩스:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "회사(학교):"
            Height          =   255
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "집:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "기본 정보"
         Height          =   1335
         Left            =   -73080
         TabIndex        =   4
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txtAddress 
            Height          =   270
            Left            =   2520
            TabIndex        =   42
            Top             =   900
            Width           =   2295
         End
         Begin VB.TextBox txtPostalCode 
            Height          =   270
            Left            =   1080
            TabIndex        =   39
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtCellPhone 
            Height          =   270
            Left            =   3360
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEmail 
            Height          =   270
            Left            =   1080
            TabIndex        =   15
            Top             =   550
            Width           =   3735
         End
         Begin VB.Label Label13 
            Caption         =   "주소:"
            Height          =   255
            Left            =   2040
            TabIndex        =   41
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "우편번호:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   950
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "전자우편:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "휴대전화:"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "이름:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox lvContacts 
         Height          =   4020
         ItemData        =   "frmMain.frx":2F6E
         Left            =   -74880
         List            =   "frmMain.frx":2F75
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4170
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   7355
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         MonthColumns    =   3
         MonthRows       =   2
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   48037889
         CurrentDate     =   43858
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6375
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10001
            Text            =   "날짜를 누르십시오."
            TextSave        =   "날짜를 누르십시오."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2020-02-08"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "오후 11:45"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuFileProperties 
         Caption         =   "일정 목록(&I)..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFilePlanBrowser 
         Caption         =   "모든 일정/데이터 색인(&B)..."
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "저장(&S)"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "보기(&V)"
      Visible         =   0   'False
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "상태 표시줄(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu erfaefewrfrfwe5r 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "옵션(&O)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "도움말(&C)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "색인(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "정보(&A)"
      End
   End
   Begin VB.Menu mnuDateMenu 
      Caption         =   "일정(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuTodaysPlan 
         Caption         =   "이날의 일정(&T)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim Contact As Integer
Dim iFIleNo As Integer
Dim Task As Integer

Sub LoadContacts()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CONTACTS"
    
    lvContacts.Clear
    lvContacts.AddItem "새 연락처 추가..."
    
    lvContactFiles.Refresh
    
    lvContacts.ListIndex = 0
    lvContactFiles.Path = "C:\CALPLANS\CONTACTS"
    
    For Contact = 0 To lvContactFiles.ListCount - 1
        lvContacts.AddItem lvContactFiles.List(Contact)
    Next Contact
End Sub

Private Sub cmdDelContact_Click()
    If MsgBox(lvContacts.List(lvContacts.ListIndex) & " 연락처를 삭제하시겠습니까?", vbQuestion + vbOKCancel, "주소록 삭제") = vbOK Then
        Kill "C:\CALPLANS\CONTACTS\" & lvContacts.List(lvContacts.ListIndex)
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "CellPhone"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Email"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Home"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Fax"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Company"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "OtherNum"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Content"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Addr"
        DeleteSetting "Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Postal"
        LoadContacts
    End If
End Sub

Private Sub cmdDeleteAllContacts_Click()
    frmOptions.cmdDelContacts_Click
End Sub

Private Sub cmdDeleteAllTasks_Click()
    frmOptions.cmdDelTasks_Click
End Sub

Private Sub cmdDelTask_Click()
    On Error Resume Next
    If MsgBox(txtTaskTitle.Text & " 작업을 삭제하시겠습니까?", vbOKCancel + vbQuestion, "작업 삭제") = vbOK Then
        DeleteSetting "Calendar", "Tasks", txtTaskTitle.Text & "Perc"
        DeleteSetting "Calendar", "Tasks", txtTaskTitle.Text & "Memo"
        Kill "C:\CALPLANS\TASKS\" & txtTaskTitle.Text
    End If
    
    LoadTasks
End Sub

Private Sub cmdEndPrg_Click()
    mnuFileExit_Click
End Sub

Private Sub cmdHelp_Click()
    PopupMenu mnuHelp, , Me.Width - 2350, 400
End Sub

Private Sub cmdOptions_Click()
    mnuViewOptions_Click
End Sub

Private Sub cmdPlanIndex_Click()
    mnuFilePlanBrowser_Click
End Sub

Private Sub cmdPlanList_Click()
    mnuFileProperties_Click
End Sub

Private Sub cmdResetFields_Click()
    If MsgBox("계속하시겠습니까?", vbOKCancel + vbQuestion, "초기화") = vbOK Then
        txtCellPhone.Text = ""
        txtEmail.Text = ""
        txtPostalCode.Text = ""
        txtAddress.Text = ""
        txtHome.Text = ""
        txtCompany.Text = ""
        txtFax.Text = ""
        txtOtherNumber.Text = ""
    End If
End Sub

Private Sub cmdSaveContact_Click()
    On Error Resume Next
    If InStr(1, txtName.Text, "?") > 0 Or InStr(1, txtName.Text, "\") > 0 Or InStr(1, txtName.Text, "|") > 0 Or InStr(1, txtName.Text, "/") > 0 Or InStr(1, txtName.Text, "*") > 0 Or InStr(1, txtName.Text, ":") > 0 Or InStr(1, txtName.Text, ".") > 0 Or InStr(1, txtName.Text, ChrW$(34)) > 0 Then
        MsgBox "이름의 값이 올바르지 않습니다.", 16, "입력 값 오류:"
    End If
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "CellPhone", txtCellPhone.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "Email", txtEmail.Text
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "Home", txtHome.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "Company", txtCompany.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "Fax", txtFax.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "OtherNum", txtOtherNumber.Text
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "Content", txtContent.Text
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "Postal", txtPostalCode.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "Addr", txtAddress.Text
    
    If lvContacts.List(lvContacts.ListIndex) = "새 연락처 추가..." Then
        '해당 연락처가 존재함을 알리는 파일을 만든다.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFIleNo = FreeFile
        '파일을 연다.
        Open "C:\CALPLANS\CONTACTS\" & txtName.Text For Output As #iFIleNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFIleNo, ""
        
        '파일을 닫는다.
        Close #iFIleNo
        
        txtName.Text = ""
        
        txtCellPhone.Text = ""
        txtEmail.Text = ""
        
        txtHome.Text = ""
        txtCompany.Text = ""
        txtFax.Text = ""
        txtOtherNumber.Text = ""
        
        txtContent.Text = ""
        
        txtPostalCode.Text = ""
    End If
    
    LoadContacts
End Sub

Sub LoadTasks()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\TASKS"
    
    lvTaskFiles.Path = "C:\CALPLANS\TASKS"
    lvTaskFiles.Refresh
    lvTasks.Clear
    
    lvTasks.AddItem "새 작업 추가..."
    
    For Task = 0 To lvTaskFiles.ListCount - 1
        lvTasks.AddItem lvTaskFiles.List(Task)
        If GetSetting("Calendar", "Tasks", lvTaskFiles.List(Task) & "Perc", "0") = "100" Then
            lvTasks.Selected(Task + 1) = True
        End If
    Next Task
    
    lvTasks.ListIndex = 0
    
    txtTaskTitle.Text = ""
    txtPercentage.Text = ""
    txtMemo.Text = ""
End Sub

Private Sub cmdSaveTask_Click()
    If InStr(1, txtTaskTitle.Text, "?") > 0 Or InStr(1, txtTaskTitle.Text, "\") > 0 Or InStr(1, txtTaskTitle.Text, "|") > 0 Or InStr(1, txtTaskTitle.Text, "/") > 0 Or InStr(1, txtTaskTitle.Text, "*") > 0 Or InStr(1, txtTaskTitle.Text, ":") > 0 Or InStr(1, txtTaskTitle.Text, ".") > 0 Or InStr(1, txtTaskTitle.Text, ChrW$(34)) > 0 Then
        MsgBox "제목의 값이 올바르지 않습니다.", 16, "입력 값 오류:"
    End If
    
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Perc", txtPercentage.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Memo", txtMemo.Text
    
    If lvTasks.List(lvTasks.ListIndex) = "새 작업 추가..." Then
        '해당 작업이 존재함을 알리는 파일을 만든다.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFIleNo = FreeFile
        '파일을 연다.
        Open "C:\CALPLANS\TASKS\" & txtTaskTitle.Text For Output As #iFIleNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFIleNo, ""
        
        '파일을 닫는다.
        Close #iFIleNo
        
        txtTaskTitle.Text = ""
        txtPercentage.Text = ""
        txtMemo.Text = ""
    End If
    
    LoadTasks
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CONTACTS"
    MkDir "C:\CALPLANS\TASKS"

    Select Case Command
        Case "/?"
            MessageBox "일정관리자 풀그림을 시작합니다." & vbCrLf & vbCrLf & _
                   "    PLANMGR.EXE [/R]" & vbCrLf & vbCrLf & _
                   "    /R  알리미 프로그램만 시작합니다.", _
                   "일정관리자 도움말", Me
            End
        Case "/R"
            frmNotifyMgr.Show
            Exit Sub
    End Select
    
    'mnuHelpAbout.Caption = App.Title & " 정보(&A)"
    
    frmNotifyMgr.Show

    Me.Left = GetSetting("Calendar", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("Calendar", "Settings", "MainTop", 1000)
    
    Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab)
    Me.Caption = Me.Caption & " (" & MonthView1.Year & "년 " & MonthView1.Month & "월)"
    
    If GetSetting("Calendar", "Config", "FirstRun", "0") = "0" Then
        SaveSetting "Calendar", "Config", "FirstRun", "1"
        MessageBox "이 풀그림이 종료된 상태에서도 알림을 받으려면 " & ChrW$(34) & Dir1.Path & "\PLNMGR32.EXE /R" & ChrW$(34) & _
               "(경로 복사됨) 바로가기를 시작프로그램에 추가하십시오.", "알리미 활성화", Me
        Clipboard.SetText ChrW$(34) & Dir1.Path & "\PLANMGR.EXE /R" & ChrW$(34)
    End If
    
    LoadContacts
    LoadTasks
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting "Calendar", "Settings", "MainLeft", Me.Left
        SaveSetting "Calendar", "Settings", "MainTop", Me.Top
    End If
    
    Cancel = 1
    Me.Hide
    frmNotifyMgr.Show
End Sub

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvContacts_Click()
    If lvContacts.List(lvContacts.ListIndex) = "새 연락처 추가..." Then
        txtName.BackColor = &H80000005
        txtName.Locked = False
        
        txtName.Text = ""
        
        txtCellPhone.Text = ""
        txtEmail.Text = ""
        
        txtHome.Text = ""
        txtCompany.Text = ""
        txtFax.Text = ""
        txtOtherNumber.Text = ""
        
        txtAddress.Text = ""
        txtPostalCode.Text = ""
        
        txtContent.Text = ""
        
        cmdDelContact.Enabled = False
        
        If SSTab1.Tab = 1 Then Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (새 주소록 추가)"
    Else
        txtName.BackColor = &H8000000F
        txtName.Locked = True
        
        txtName.Text = lvContacts.List(lvContacts.ListIndex)
        
        txtCellPhone.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "CellPhone", "")
        txtEmail.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Email", "")
        
        txtHome.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Home", "")
        txtCompany.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Company", "")
        txtFax.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Fax", "")
        txtOtherNumber.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "OtherNum", "")
        
        txtPostalCode.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Postal", "")
        txtAddress.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Addr", "")
        
        txtContent.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Content", "")
        
        cmdDelContact.Enabled = True
        
        Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (" & txtName.Text & ")"
    End If
End Sub

Private Sub lvTasks_Click()
    If lvTasks.List(lvTasks.ListIndex) = "새 작업 추가..." Then
        cmdDelTask.Enabled = False
    Else
        cmdDelTask.Enabled = True
    End If
    
    txtTaskTitle.Text = lvTasks.List(lvTasks.ListIndex)
    txtPercentage.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Perc", "")
    txtMemo.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Memo", "")
    
    If SSTab1.Tab = 2 Then
        Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (" & lvTasks.List(lvTasks.ListIndex) & ")"
        If lvTasks.List(lvTasks.ListIndex) = "새 작업 추가..." Then
            Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (새 작업 추가)"
        End If
    End If
End Sub

Private Sub lvTasks_ItemCheck(Item As Integer)
    If lvTasks.List(Item) <> "새 작업 추가..." Then
        If lvTasks.Selected(Item) = True Then
            SaveSetting "Calendar", "Tasks", lvTasks.List(Item) & "Perc", "100"
        Else
            SaveSetting "Calendar", "Tasks", lvTasks.List(Item) & "Perc", "0"
        End If
    End If

    If lvTasks.Selected(Item) = True Then
        txtPercentage.Text = 100
    Else
        txtPercentage.Text = 0
    End If
    
    lvTasks.ListIndex = Item
End Sub

Private Sub mnuFilePlanBrowser_Click()
    frmDataBrowser.Show vbModal, Me
End Sub

Private Sub mnuTodaysPlan_Click()
    MonthView1_DateClick MonthView1.SelStart
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    frmPlans.CurrentDate = DateClicked
    frmPlans.Show vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    '이 프로젝트에 대한 도움말 파일이 없으면 사용자에게 메시지를 표시합니다.
    '사용자는 [프로젝트 속성] 대화 상자에서 응용 프로그램에 대한
    '도움말 파일을 설정할 수 있습니다.
    If Len(App.HelpFile) = 0 Then
        MessageBox "도움말 목차를 표시할 수 없습니다. 이 프로그램과 연관된 도움말이 없습니다.", App.Title, Me, 16
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub



Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    '이 프로젝트에 대한 도움말 파일이 없으면 사용자에게 메시지를 표시합니다.
    '사용자는 [프로젝트 속성] 대화 상자에서 응용 프로그램에 대한
    '도움말 파일을 설정할 수 있습니다.
    If Len(App.HelpFile) = 0 Then
        MessageBox "도움말 목차를 표시할 수 없습니다. 이 프로그램과 연관된 도움말이 없습니다.", App.Title, Me, 16
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuToolsOptions_Click()
    '작업: 'mnuToolsOptions_Click' 코드를 추가하십시오.
    MsgBox "'mnuToolsOptions_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewWebBrowser_Click()
    '작업: 'mnuViewWebBrowser_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewWebBrowser_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewOptions_Click()
    '작업: 'mnuViewOptions_Click' 코드를 추가하십시오.
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    '작업: 'mnuViewRefresh_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewRefresh_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    '작업: 'mnuEditPasteSpecial_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditPasteSpecial_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditPaste_Click()
    '작업: 'mnuEditPaste_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditPaste_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditCopy_Click()
    '작업: 'mnuEditCopy_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditCopy_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditCut_Click()
    '작업: 'mnuEditCut_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditCut_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditUndo_Click()
    '작업: 'mnuEditUndo_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditUndo_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileExit_Click()
    '알리미는 남아야 하므로 폼을 숨기기만 한다.
    'Unload Me
    Me.Hide
End Sub

Private Sub mnuFileSend_Click()
    '작업: 'mnuFileSend_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSend_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePrint_Click()
    '작업: 'mnuFilePrint_Click' 코드를 추가하십시오.
    MsgBox "'mnuFilePrint_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePrintPreview_Click()
    '작업: 'mnuFilePrintPreview_Click' 코드를 추가하십시오.
    MsgBox "'mnuFilePrintPreview_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "페이지 설정"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    '작업: 'mnuFileProperties_Click' 코드를 추가하십시오.
    MonthView1_DateClick MonthView1.SelStart
End Sub

Private Sub mnuFileSaveAll_Click()
    '작업: 'mnuFileSaveAll_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSaveAll_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSaveAs_Click()
    '작업: 'mnuFileSaveAs_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSaveAs_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSave_Click()
    '작업: 'mnuFileSave_Click' 코드를 추가하십시오.
    If SSTab1.Tab = 1 Then
        cmdSaveContact_Click
    Else
        cmdSaveTask_Click
    End If
End Sub

Private Sub mnuFileClose_Click()
    '작업: 'mnuFileClose_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileClose_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "열기"
        .CancelError = False
        '작업: Common Dialog 컨트롤의 플래그와 특성을 설정합니다.
        .Filter = "모든 파일(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    '작업: 코드를 추가하여 열려 있는 파일을 처리합니다.

End Sub

Private Sub mnuFileNew_Click()
    '작업: 'mnuFileNew_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileNew_Click' 코드를 추가하십시오."
End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuTodaysPlan.Caption = MonthView1.SelStart & "의 일정"
        PopupMenu mnuDateMenu
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab)
    If SSTab1.Tab = 0 Then
        Me.Caption = Me.Caption & " (" & MonthView1.Year & "년 " & MonthView1.Month & "월)"
    ElseIf SSTab1.Tab = 1 Then
        Me.Caption = Me.Caption & " (새 주소록 추가)"
    ElseIf SSTab1.Tab = 2 Then
        Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (새 작업 추가)"
    End If
    
    
    
    If SSTab1.Tab > 0 Then
        mnuFileBar0.Visible = True
        mnuFileSave.Visible = True
    Else
        mnuFileBar0.Visible = False
        mnuFileSave.Visible = False
    End If
End Sub

Private Sub tglStatusBar_Click()
    mnuViewStatusBar_Click
End Sub

Private Sub txtPercentage_Change()
    On Error Resume Next
    TaskProgress.Value = txtPercentage.Text
    
    If TaskProgress.Value = 100 Then
        lvTasks.Selected(lvTasks.ListIndex) = True
    Else
        lvTasks.Selected(lvTasks.ListIndex) = False
    End If
End Sub
