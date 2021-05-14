VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMainOld 
   BorderStyle     =   0  '없음
   Caption         =   "frmMain"
   ClientHeight    =   8325
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11700
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleMode       =   0  '사용자
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   8760
      TabIndex        =   162
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
      Caption         =   "_"
      Height          =   375
      Left            =   9600
      TabIndex        =   159
      ToolTipText     =   "프로그램의 도움말과 관련된 항목입니다."
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "×"
      Height          =   375
      Left            =   10440
      TabIndex        =   158
      ToolTipText     =   "프로그램의 도움말과 관련된 항목입니다."
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Visible         =   0   'False
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15452
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2021-05-14"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "오후 5:11"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRibbonFile 
      BackColor       =   &H8000000C&
      Caption         =   "cmdRibbonFile"
      Height          =   330
      Left            =   240
      TabIndex        =   143
      Top             =   960
      Width           =   1125
   End
   Begin VB.CommandButton cmdMnuAbout 
      BackColor       =   &H8000000C&
      Caption         =   "☎"
      Height          =   375
      Left            =   7080
      TabIndex        =   140
      ToolTipText     =   "프로그램 정보를 보여줍니다."
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdMnuOptions 
      BackColor       =   &H8000000C&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   139
      ToolTipText     =   "환경 설정"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox pbxTodaysPlanTab 
      Height          =   1335
      Left            =   11760
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   138
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox pbxRibbonBackground 
      Height          =   975
      Left            =   11400
      Picture         =   "frmMain.frx":27604
      ScaleHeight     =   915
      ScaleWidth      =   4635
      TabIndex        =   137
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.FileListBox lvAlarmList 
      Height          =   270
      Left            =   3480
      TabIndex        =   133
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timAlarmChecker 
      Interval        =   10000
      Left            =   12960
      Top             =   480
   End
   Begin VB.FileListBox lvGroupList 
      Height          =   270
      Left            =   10800
      TabIndex        =   112
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdTltRef 
      Caption         =   "갱신(&R)"
      Height          =   300
      Left            =   8880
      TabIndex        =   52
      ToolTipText     =   "오늘의 일정목록을 갱신합니다."
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   12720
      Top             =   240
   End
   Begin TabDlg.SSTab ssTodaysPlan 
      Height          =   5535
      Left            =   8760
      TabIndex        =   48
      Top             =   960
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      TabCaption(0)   =   "오늘 일정"
      TabPicture(0)   =   "frmMain.frx":5C116
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvTodaysPlan"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvTodaysPlans"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "내일 일정"
      TabPicture(1)   =   "frmMain.frx":5C132
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvTmrPlans"
      Tab(1).ControlCount=   1
      Begin VB.FileListBox lvTmrPlans 
         Height          =   5130
         Left            =   -74880
         TabIndex        =   51
         Top             =   360
         Width           =   1935
      End
      Begin VB.FileListBox lvTodaysPlans 
         Height          =   270
         Left            =   240
         TabIndex        =   50
         Top             =   140
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.FileListBox lvTodaysPlan 
         Height          =   5130
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H8000000C&
      Caption         =   "?"
      Height          =   375
      Left            =   8760
      TabIndex        =   46
      ToolTipText     =   "프로그램의 도움말과 관련된 항목입니다."
      Top             =   120
      Width           =   375
   End
   Begin TabDlg.SSTab ssRibbonMenu 
      Height          =   1335
      Left            =   240
      TabIndex        =   43
      Top             =   960
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   548
      TabMaxWidth     =   1940
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      MouseIcon       =   "frmMain.frx":5C14E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain.frx":5C16A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "home"
      TabPicture(1)   =   "frmMain.frx":5C186
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "timHidemenu"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "view"
      TabPicture(2)   =   "frmMain.frx":5C5D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tglStatusBar"
      Tab(2).Control(1)=   "tglCalWeekNum"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "plans"
      TabPicture(3)   =   "frmMain.frx":5CA2A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).ControlCount=   1
      Begin VB.Timer timHidemenu 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   12360
         Top             =   240
      End
      Begin VB.Frame Frame8 
         Caption         =   "-"
         Height          =   855
         Left            =   -74880
         TabIndex        =   149
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdDelAllTodaysPlan 
            Caption         =   "이날의   일정 삭제"
            Height          =   735
            Left            =   1320
            Picture         =   "frmMain.frx":5CE7C
            Style           =   1  '그래픽
            TabIndex        =   151
            ToolTipText     =   "선택한 날의 일정을 모두 삭제합니다."
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdTodaysPlan 
            Caption         =   "이날의 일정"
            Height          =   735
            Left            =   120
            Picture         =   "frmMain.frx":5D2BE
            Style           =   1  '그래픽
            TabIndex        =   150
            ToolTipText     =   "표시한 날짜의 일정 목록을 표시합니다."
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "-"
         Height          =   855
         Left            =   2640
         TabIndex        =   147
         Top             =   360
         Width           =   1335
         Begin VB.CommandButton cmdEndPrg 
            Caption         =   "끝내기"
            Height          =   735
            Left            =   120
            Picture         =   "frmMain.frx":5D700
            Style           =   1  '그래픽
            TabIndex        =   148
            ToolTipText     =   "프로그램을 끝냅니다."
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   144
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdPlanIndex 
            Caption         =   "데이터 색인"
            Height          =   735
            Left            =   1200
            Picture         =   "frmMain.frx":5DB42
            Style           =   1  '그래픽
            TabIndex        =   146
            ToolTipText     =   "주소록, 일정 전체목록입니다."
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdPlanList 
            Caption         =   "일정 목록"
            Height          =   720
            Left            =   120
            Picture         =   "frmMain.frx":5DF84
            Style           =   1  '그래픽
            TabIndex        =   145
            ToolTipText     =   "표시한 날짜의 일정 목록을 표시합니다."
            Top             =   120
            Width           =   975
         End
      End
      Begin MSForms.ToggleButton tglStatusBar 
         Height          =   840
         Left            =   -74880
         TabIndex        =   142
         ToolTipText     =   "상태표시줄을 표시하거나 숨깁니다."
         Top             =   375
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1931;1482"
         Value           =   "1"
         Caption         =   "상태표시줄"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tglCalWeekNum 
         Height          =   855
         Left            =   -73680
         TabIndex        =   141
         ToolTipText     =   "달력에서 주의 번호를 표시하거나 숨깁니다."
         Top             =   360
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1931;1508"
         Value           =   "1"
         Caption         =   "주 번호"
         Picture         =   "frmMain.frx":5E3C6
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   582
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      MouseIcon       =   "frmMain.frx":5E6E0
      TabCaption(0)   =   "일정"
      TabPicture(0)   =   "frmMain.frx":5E6FC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Dir1"
      Tab(0).Control(1)=   "MonthView1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "주소록"
      TabPicture(1)   =   "frmMain.frx":5EB4E
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
      TabPicture(2)   =   "frmMain.frx":5EFA0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvTasks"
      Tab(2).Control(1)=   "cmdSaveTask"
      Tab(2).Control(2)=   "cmdDelTask"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "lvTaskFiles"
      Tab(2).Control(5)=   "cmdDeleteAllTasks"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "일과표"
      TabPicture(3)   =   "frmMain.frx":5F3F2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblDOW"
      Tab(3).Control(1)=   "Label15"
      Tab(3).Control(2)=   "txtPlannerTF(0)"
      Tab(3).Control(3)=   "txtPlannerTF(1)"
      Tab(3).Control(4)=   "txtPlannerTF(2)"
      Tab(3).Control(5)=   "txtPlannerTF(3)"
      Tab(3).Control(6)=   "txtPlannerTF(4)"
      Tab(3).Control(7)=   "txtPlannerTF(5)"
      Tab(3).Control(8)=   "txtPlannerTF(6)"
      Tab(3).Control(9)=   "txtPlannerTF(7)"
      Tab(3).Control(10)=   "txtPlannerTF(8)"
      Tab(3).Control(11)=   "txtPlannerTF(9)"
      Tab(3).Control(12)=   "txtPlannerTF(10)"
      Tab(3).Control(13)=   "txtPlannerTF(11)"
      Tab(3).Control(14)=   "txtPlannerTF(12)"
      Tab(3).Control(15)=   "txtPlannerTF(13)"
      Tab(3).Control(16)=   "txtPlannerTF(14)"
      Tab(3).Control(17)=   "txtPlannerTF(15)"
      Tab(3).Control(18)=   "txtPlannerTF(16)"
      Tab(3).Control(19)=   "txtPlannerTF(17)"
      Tab(3).Control(20)=   "txtPlannerTF(18)"
      Tab(3).Control(21)=   "txtPlannerTF(19)"
      Tab(3).Control(22)=   "txtPlannerTF(20)"
      Tab(3).Control(23)=   "txtPlannerTF(21)"
      Tab(3).Control(24)=   "txtPlannerTF(22)"
      Tab(3).Control(25)=   "txtPlannerTF(23)"
      Tab(3).Control(26)=   "txtPlannerTF(24)"
      Tab(3).Control(27)=   "txtPlannerTF(25)"
      Tab(3).Control(28)=   "txtPlannerTF(26)"
      Tab(3).Control(29)=   "txtPlannerTF(27)"
      Tab(3).Control(30)=   "txtPlannerTF(28)"
      Tab(3).Control(31)=   "txtPlannerTF(29)"
      Tab(3).Control(32)=   "txtPlannerTF(30)"
      Tab(3).Control(33)=   "txtPlannerTF(31)"
      Tab(3).Control(34)=   "txtPlannerTF(32)"
      Tab(3).Control(35)=   "txtPlannerTF(33)"
      Tab(3).Control(36)=   "txtPlannerTF(34)"
      Tab(3).Control(37)=   "txtPlannerTF(35)"
      Tab(3).Control(38)=   "txtPlannerTF(36)"
      Tab(3).Control(39)=   "txtPlannerTF(37)"
      Tab(3).Control(40)=   "txtPlannerTF(38)"
      Tab(3).Control(41)=   "txtPlannerTF(39)"
      Tab(3).Control(42)=   "txtPlannerTF(40)"
      Tab(3).Control(43)=   "txtPlannerTF(41)"
      Tab(3).Control(44)=   "txtPlannerTF(42)"
      Tab(3).Control(45)=   "txtPlannerTF(43)"
      Tab(3).Control(46)=   "txtPlannerTF(44)"
      Tab(3).Control(47)=   "txtPlannerTF(45)"
      Tab(3).Control(48)=   "txtPlannerTF(46)"
      Tab(3).Control(49)=   "txtPlannerTF(47)"
      Tab(3).Control(50)=   "txtPlannerTF(48)"
      Tab(3).Control(51)=   "sdcmdSavePlanner"
      Tab(3).ControlCount=   52
      TabCaption(4)   =   "알람"
      TabPicture(4)   =   "frmMain.frx":5F70C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label18"
      Tab(4).Control(1)=   "Label19"
      Tab(4).Control(2)=   "Label20"
      Tab(4).Control(3)=   "lvAlarms"
      Tab(4).Control(4)=   "txtAlarmTitle"
      Tab(4).Control(5)=   "txtTimeHrs"
      Tab(4).Control(6)=   "txtTimeMin"
      Tab(4).Control(7)=   "Frame5"
      Tab(4).Control(8)=   "cmdResetAF"
      Tab(4).Control(9)=   "cmdSaveAlarm"
      Tab(4).Control(10)=   "cmdDeleteAlarm"
      Tab(4).Control(11)=   "txtAlarmMemo"
      Tab(4).Control(12)=   "lvAlarmFiles"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   " 메모"
      TabPicture(5)   =   "frmMain.frx":5FB5E
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Text1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   4095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   161
         Top             =   120
         Width           =   8175
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4230
         Left            =   -74880
         TabIndex        =   157
         Top             =   98
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   7461
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483636
         Appearance      =   1
         MonthColumns    =   3
         MonthRows       =   2
         StartOfWeek     =   65798145
         CurrentDate     =   44330
      End
      Begin VB.FileListBox lvAlarmFiles 
         Height          =   270
         Left            =   -67560
         TabIndex        =   132
         Top             =   -22
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAlarmMemo 
         Height          =   1215
         Left            =   -72360
         MultiLine       =   -1  'True
         TabIndex        =   131
         Top             =   2498
         Width           =   5535
      End
      Begin VB.CommandButton cmdDeleteAlarm 
         Caption         =   "cmdDeleteAlarm"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69600
         TabIndex        =   129
         Top             =   3818
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveAlarm 
         Caption         =   "cmdSaveAlarm"
         Height          =   375
         Left            =   -68160
         TabIndex        =   128
         Top             =   3818
         Width           =   1335
      End
      Begin VB.CommandButton cmdResetAF 
         Caption         =   "cmdResetAF"
         Height          =   375
         Left            =   -72360
         TabIndex        =   127
         Top             =   3818
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1335
         Left            =   -72360
         TabIndex        =   119
         Top             =   858
         Width           =   5535
         Begin VB.CommandButton cmdSelectAllDW 
            Caption         =   "cmdSelectAllDW"
            Height          =   320
            Left            =   1560
            TabIndex        =   136
            Top             =   940
            Width           =   1215
         End
         Begin VB.CommandButton cmdUnselectAllDW 
            Caption         =   "cmdUnselectAllDW"
            Height          =   320
            Left            =   2880
            TabIndex        =   135
            Top             =   940
            Width           =   1215
         End
         Begin VB.CommandButton cmdRelectAllDW 
            Caption         =   "cmdRelectAllDW"
            Height          =   320
            Left            =   4200
            TabIndex        =   134
            Top             =   940
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "토요일"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   126
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "금요일"
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   125
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "목요일"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   124
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "수요일"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   123
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "화요일"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   122
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "월요일"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   121
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkDayOfWeeks 
            Caption         =   "일요일"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtTimeMin 
         Height          =   270
         Left            =   -71400
         TabIndex        =   118
         Top             =   578
         Width           =   375
      End
      Begin VB.TextBox txtTimeHrs 
         Height          =   270
         Left            =   -71760
         TabIndex        =   117
         Top             =   578
         Width           =   375
      End
      Begin VB.TextBox txtAlarmTitle 
         Height          =   270
         Left            =   -71760
         TabIndex        =   115
         Top             =   218
         Width           =   4935
      End
      Begin ComctlLib.ListView lvAlarms 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   113
         Top             =   98
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton sdcmdSavePlanner 
         Caption         =   "sdcmdSavePlanner"
         Height          =   375
         Left            =   -68040
         TabIndex        =   104
         Top             =   98
         Width           =   1215
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   48
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   103
         Top             =   3578
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   47
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   102
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   46
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   101
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   45
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   44
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   99
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   43
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   98
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   42
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   41
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   3098
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   40
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   39
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   94
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   38
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   93
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   37
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   36
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   91
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   35
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   34
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   2618
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   33
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   32
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   87
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   31
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   30
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   29
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   28
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   27
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   2018
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   26
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   25
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   24
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   79
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   23
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   22
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   21
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   20
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   1418
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   19
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   18
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   17
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   16
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   15
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   14
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   13
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   938
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   12
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   11
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   10
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   9
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   8
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   7
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   6
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   458
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   5
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   60
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   4
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   3
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   2
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   1
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   0
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   458
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteAllTasks 
         Caption         =   "cmdDeleteAllTasks"
         Height          =   495
         Left            =   -67920
         TabIndex        =   47
         Top             =   3676
         Width           =   1215
      End
      Begin VB.CommandButton cmdResetFields 
         Caption         =   "내용 초기화(&R)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   45
         Top             =   3676
         Width           =   1350
      End
      Begin VB.CommandButton cmdDeleteAllContacts 
         Caption         =   "clear"
         Height          =   495
         Left            =   -68040
         TabIndex        =   44
         Top             =   2476
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   300
         Left            =   -66960
         TabIndex        =   42
         Top             =   -44
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.FileListBox lvTaskFiles 
         Height          =   270
         Left            =   -67920
         TabIndex        =   39
         Top             =   676
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4095
         Left            =   -72480
         TabIndex        =   27
         Top             =   98
         Width           =   4455
         Begin VB.TextBox txtPart 
            Height          =   270
            Left            =   1080
            TabIndex        =   109
            Top             =   1920
            Width           =   3255
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   270
            Left            =   600
            TabIndex        =   108
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtImpt"
            BuddyDispid     =   196658
            OrigLeft        =   600
            OrigTop         =   1920
            OrigRight       =   855
            OrigBottom      =   2175
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtImpt 
            Height          =   270
            Left            =   120
            MaxLength       =   2
            TabIndex        =   107
            Text            =   "1"
            Top             =   1920
            Width           =   480
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   3840
            TabIndex        =   36
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtPercentage"
            BuddyDispid     =   196661
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
            Height          =   1335
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   35
            Top             =   2640
            Width           =   4215
         End
         Begin VB.TextBox txtTaskTitle 
            Height          =   270
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtPercentage 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   3450
            TabIndex        =   31
            Text            =   "0"
            Top             =   1200
            Width           =   420
         End
         Begin ComctlLib.ProgressBar TaskProgress 
            Height          =   300
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label Label16 
            Caption         =   "Label16"
            Height          =   255
            Left            =   1080
            TabIndex        =   106
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "%"
            Height          =   255
            Left            =   4155
            TabIndex        =   30
            Top             =   1245
            Width           =   135
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdDelTask 
         Caption         =   "cmdDelTask"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -67920
         TabIndex        =   26
         Top             =   3076
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveTask 
         Caption         =   "cmdSaveTask"
         Height          =   495
         Left            =   -67920
         TabIndex        =   25
         Top             =   76
         Width           =   1215
      End
      Begin VB.ListBox lvTasks 
         Height          =   4050
         ItemData        =   "frmMain.frx":5FFB0
         Left            =   -74880
         List            =   "frmMain.frx":5FFB7
         Style           =   1  '확인란
         TabIndex        =   24
         Top             =   76
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelContact 
         Caption         =   "delete"
         Height          =   495
         Left            =   -68040
         TabIndex        =   23
         Top             =   1396
         Width           =   1335
      End
      Begin VB.FileListBox lvContactFiles 
         Height          =   270
         Left            =   -69240
         TabIndex        =   22
         Top             =   76
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1575
         Left            =   -73080
         TabIndex        =   9
         Top             =   2596
         Width           =   4935
         Begin VB.TextBox txtBDay 
            Height          =   270
            Left            =   1320
            TabIndex        =   155
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtBMonth 
            Height          =   270
            Left            =   840
            TabIndex        =   154
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtBYear 
            Height          =   270
            Left            =   120
            TabIndex        =   153
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtContent 
            Height          =   975
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   21
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label22 
            Caption         =   "Label22"
            Height          =   255
            Left            =   1920
            TabIndex        =   156
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdSaveContact 
         Caption         =   "저장(&S)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   8
         Top             =   196
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   975
         Left            =   -73080
         TabIndex        =   7
         Top             =   1516
         Width           =   4935
         Begin VB.TextBox txtOtherNumber 
            Height          =   270
            Left            =   2880
            TabIndex        =   20
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtFax 
            Height          =   270
            Left            =   600
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtHome 
            Height          =   270
            Left            =   720
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCompany 
            Height          =   270
            Left            =   3000
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "기타:"
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "팩스:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "회사(학교):"
            Height          =   255
            Left            =   2040
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "집:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   -73080
         TabIndex        =   3
         Top             =   98
         Width           =   4935
         Begin VB.ComboBox cmbGroup 
            Height          =   300
            Left            =   3360
            Style           =   2  '드롭다운 목록
            TabIndex        =   111
            Top             =   560
            Width           =   1455
         End
         Begin VB.TextBox txtAddress 
            Height          =   270
            Left            =   2520
            TabIndex        =   41
            Top             =   900
            Width           =   2295
         End
         Begin VB.TextBox txtPostalCode 
            Height          =   270
            Left            =   1080
            TabIndex        =   38
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtCellPhone 
            Height          =   270
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   600
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEmail 
            Height          =   270
            Left            =   1080
            TabIndex        =   14
            Top             =   550
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "그룹:"
            Height          =   255
            Left            =   2880
            TabIndex        =   110
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "주소:"
            Height          =   255
            Left            =   2040
            TabIndex        =   40
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "우편번호:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   950
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "전자우편:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "휴대전화:"
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "이름:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox lvContacts 
         Height          =   4020
         ItemData        =   "frmMain.frx":5FFCC
         Left            =   -74880
         List            =   "frmMain.frx":5FFD3
         TabIndex        =   2
         Top             =   76
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   -72360
         TabIndex        =   130
         Top             =   2258
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   -72360
         TabIndex        =   116
         Top             =   578
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   -72360
         TabIndex        =   114
         Top             =   218
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "7             9             12              15             18            21               22-"
         Height          =   225
         Left            =   -74280
         TabIndex        =   54
         Top             =   218
         Width           =   6135
      End
      Begin VB.Label lblDOW 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   53
         Top             =   578
         Width           =   255
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   600
      TabIndex        =   160
      Top             =   7320
      Width           =   9255
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   10080
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   7740
      Left            =   0
      Picture         =   "frmMain.frx":5FFEA
      Top             =   0
      Width           =   11385
   End
   Begin VB.Image menuhover 
      Height          =   1455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
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
   Begin VB.Menu mnuDateMenu 
      Caption         =   "일정(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuTodaysPlan 
         Caption         =   "이날의 일정(&T)..."
      End
      Begin VB.Menu mnuPlansClear 
         Caption         =   "선택한 날짜의 일정 모두 삭제(&D)"
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
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuQuit 
         Caption         =   "종료(&Q)"
      End
   End
End
Attribute VB_Name = "frmMainOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://www.vbforums.com/showthread.php?396385-Making-A-Form-Transparent-(But-with-visible-controls)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
 
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2


Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim Contact As Integer
Dim iFileNo As Integer
Dim Task As Integer

'퍼온곳: http://www.vbforums.com/showthread.php?595990-VB6-System-tray-icon-systray
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' The following code is required:
Option Explicit



Sub ClearAlarmFields()
    cmdDeleteAlarm.Enabled = False
    
    txtAlarmTitle.Text = ""
    txtTimeHrs.Text = ""
    txtTimeMin.Text = ""
    
    Dim i As Integer
    For i = 0 To 6
        chkDayOfWeeks(i).Value = 0
    Next i
    
    txtAlarmMemo.Text = ""
    
    txtAlarmTitle.Enabled = True
End Sub



Private Sub cmdDeleteAlarm_Click()
    On Error Resume Next
    If Confirm("한 번만 경고하는데 선택한 알람을 삭제할까요?", "경고", Me) Then
        Kill "C:\CALPLANS\ALARMS\" & lvAlarms.SelectedItem.SubItems(1)
        
        ClearAlarmFields
        
        LoadAlarms
    End If
End Sub

Private Sub cmdMnuAbout_Click()
    mnuHelpAbout_Click
End Sub

Private Sub cmdMnuOptions_Click()
    cmdOptions_Click
End Sub

Private Sub cmdRelectAllDW_Click()
    Dim i As Integer
    For i = 0 To 6
        If chkDayOfWeeks(i).Value = 1 Then
            chkDayOfWeeks(i).Value = 0
        Else
            chkDayOfWeeks(i).Value = 1
        End If
    Next i
End Sub

Private Sub cmdSaveAlarm_Click()
    '입력값을 검사한다.
    If Mid$(txtTimeMin.Text, 1, 1) = "0" Then
        txtTimeMin.Text = Mid$(txtTimeMin.Text, 2, 1)
    End If
    If InStr(1, txtAlarmTitle.Text, "?") > 0 Or InStr(1, txtAlarmTitle.Text, "\") > 0 Or InStr(1, txtAlarmTitle.Text, "|") > 0 Or InStr(1, txtAlarmTitle.Text, ".") > 0 Or InStr(1, txtAlarmTitle.Text, "/") > 0 Or InStr(1, txtAlarmTitle.Text, "*") > 0 Or InStr(1, txtAlarmTitle.Text, ":") > 0 Or InStr(1, txtAlarmTitle.Text, ChrW$(34)) > 0 Or txtAlarmTitle.Text = LoadLang("새 알람 추가...", "New...") Then
        MessageBox "제목의 값이 올바르지 않습니다.", "입력 값 오류", Me, 16
    End If
    If IsNumeric(txtTimeHrs.Text) = False Or IsNumeric(txtTimeMin.Text) = False Then
        MessageBox "시간의 값이 올바르지 않습니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    If GetSetting("Calendar", "Options", "NoTimeCheck", 0) = 0 Then
        If txtTimeHrs.Text > 23 Or txtTimeMin.Text > 59 Or txtTimeHrs.Text < 0 Or txtTimeMin.Text < 0 Then
            MessageBox "시간에서 시는 0부터 23, 분은 0부터 59까지의 정수이여야 합니다.", "입력 값 오류", Me, 16
            Exit Sub
    End If
    End If
    If txtAlarmTitle.Text = "" Then
        MessageBox "제목의 값은 필수입니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    
    '일정을 추가하기 전에 해당 제목의 일정이 존재하는지 확인한다.
    If FileExists("C:\CALPLANS\ALARMS\" & txtAlarmTitle.Text) = True And lvAlarms.SelectedItem.SubItems(1) = LoadLang("새 알람 추가...", "New...") Then
        MessageBox "해당 이름의 알람이 이미 존재합니다.", "처리 중 오류", Me, 16
    End If
    
    '해당 알람이 존재함을 알리는 파일을 만든다.
    'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
    Dim iFileNo As Integer
    iFileNo = FreeFile
    '파일을 연다.
    Open "C:\CALPLANS\ALARMS\" & txtAlarmTitle.Text For Output As #iFileNo
    
    '파일의 내용은 보지 않으므로 빈 칸으로...
    Print #iFileNo, ""
    
    '파일을 닫는다.
    Close #iFileNo
    
    Dim txtTime As String
    
    '레지스트리에 일정의 기타 정보를 저장한다.
    If txtTimeHrs.Text < 9 Then
        If txtTimeMin.Text < 9 Then
            txtTime = "0" & txtTimeHrs.Text & ":0" & txtTimeMin.Text
        Else
            txtTime = "0" & txtTimeHrs.Text & ":" & txtTimeMin.Text
        End If
    Else
        If txtTimeMin.Text < 9 Then
            txtTime = txtTimeHrs.Text & ":0" & txtTimeMin.Text
        Else
            txtTime = txtTimeHrs.Text & ":" & txtTimeMin.Text
        End If
    End If
    
    SaveSetting "Calendar", "Alarms", txtAlarmTitle.Text & "Time", txtTime
    SaveSetting "Calendar", "Alarms", txtAlarmTitle.Text & "Memo", txtAlarmMemo.Text
    
    Dim i As Integer
    For i = 0 To 6
        SaveSetting "Calendar", "Alarms", txtAlarmTitle.Text & "W" & CStr(i), chkDayOfWeeks(i).Value
    Next i
    
    ClearAlarmFields
    
    LoadAlarms
End Sub

Private Sub cmdSelectAllDW_Click()
    Dim i As Integer
    For i = 0 To 6
        chkDayOfWeeks(i).Value = 1
    Next i
End Sub

Private Sub cmdTltRef_Click()
    lvTodaysPlan.Refresh
    lvTodaysPlans.Refresh
    lvTmrPlans.Refresh
End Sub

' End required code
' /////////////////////////////////////////////

Sub LoadContacts()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CONTACTS"
    
    lvContacts.Clear
    lvContacts.AddItem LoadLang("새 연락처 추가...", "New...")
    
    lvContactFiles.Refresh
    
    lvContacts.ListIndex = 0
    lvContactFiles.Path = "C:\CALPLANS\CONTACTS"
    
    For Contact = 0 To lvContactFiles.ListCount - 1
        lvContacts.AddItem lvContactFiles.List(Contact)
    Next Contact
End Sub

Private Sub cmdDelContact_Click()
    If Confirm(lvContacts.List(lvContacts.ListIndex) & " 연락처를 삭제하시겠습니까?", "주소록 삭제", Me) Then
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
        
        If lvContacts.List(lvContacts.ListIndex) = "둘리" Then
            SaveSetting "Calendar", "Config", "EggEnabled", 0
            DeleteSetting "Calendar", "Config", "EggEnabled"
            
            If GetSetting("Calendar", "Options", "Ringtone", 0) = 2 Then
                SaveSetting "Calendar", "Options", "Ringtone", 0
            End If
        End If
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
    If Confirm(txtTaskTitle.Text & " 작업을 삭제하시겠습니까?", "작업 삭제", Me) Then
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
    If GetSetting("Calendar", "Options", "TP", 0) = 0 Then
        PopupMenu mnuHelp, , Me.Width - 2350 - ssTodaysPlan.Width + 100, 400
    Else
        PopupMenu mnuHelp, , Me.Width - 2350, 400
    End If
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
    If Confirm("한번만 경고합니다. 모든 입력상자의 값을 초기화하시겠습니까?", "초기화", Me) Then
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
    If InStr(1, txtName.Text, "?") > 0 Or InStr(1, txtName.Text, "\") > 0 Or InStr(1, txtName.Text, "|") > 0 Or InStr(1, txtName.Text, "/") > 0 Or InStr(1, txtName.Text, "*") > 0 Or InStr(1, txtName.Text, ":") > 0 Or InStr(1, txtName.Text, ".") > 0 Or InStr(1, txtName.Text, ChrW$(34)) > 0 Or txtName.Text = "" Then
        MessageBox "이름의 값이 올바르지 않습니다.", "입력 값 오류", Me, 16
        Exit Sub
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
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "Group", cmbGroup.Text
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "BY", txtBYear.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "BM", txtBMonth.Text
    SaveSetting "Calendar", "Contacts", txtName.Text & "BD", txtBDay.Text
    
    If txtName.Text = "둘리" And txtBYear.Text = "1983" And (txtBMonth.Text = "4" Or txtBMonth.Text = "04") And txtBDay.Text = "22" Then
        SaveSetting "Calendar", "Config", "EggEnabled", "1"
    End If
    
    If lvContacts.List(lvContacts.ListIndex) = LoadLang("새 연락처 추가...", "New...") Then
        '해당 연락처가 존재함을 알리는 파일을 만든다.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFileNo = FreeFile
        '파일을 연다.
        Open "C:\CALPLANS\CONTACTS\" & txtName.Text For Output As #iFileNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFileNo, ""
        
        '파일을 닫는다.
        Close #iFileNo
        
        txtName.Text = ""
        
        txtCellPhone.Text = ""
        txtEmail.Text = ""
        
        txtHome.Text = ""
        txtCompany.Text = ""
        txtFax.Text = ""
        txtOtherNumber.Text = ""
        
        txtContent.Text = ""
        
        txtPostalCode.Text = ""
        
        txtBMonth.Text = ""
        txtBYear.Text = ""
        txtBDay.Text = ""
        
        cmbGroup.ListIndex = 0
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
    
    lvTasks.AddItem LoadLang("새 작업 추가...", "New...")
    
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
    If InStr(1, txtTaskTitle.Text, "?") > 0 Or InStr(1, txtTaskTitle.Text, "\") > 0 Or InStr(1, txtTaskTitle.Text, "|") > 0 Or InStr(1, txtTaskTitle.Text, "/") > 0 Or InStr(1, txtTaskTitle.Text, "*") > 0 Or InStr(1, txtTaskTitle.Text, ":") > 0 Or InStr(1, txtTaskTitle.Text, ".") > 0 Or InStr(1, txtTaskTitle.Text, ChrW$(34)) > 0 Or txtTaskTitle.Text = "" Then
        MessageBox "제목의 값이 올바르지 않습니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    
    
    If IsNumeric(txtImpt.Text) = False Or txtImpt.Text < 1 Or txtImpt.Text > 10 Then
        MessageBox "중요도는 1(낮음)부터 10(높음)까지여야 합니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Perc", txtPercentage.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Memo", txtMemo.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Impt", txtImpt.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Part", txtPart.Text
    
    If lvTasks.List(lvTasks.ListIndex) = LoadLang("새 작업 추가...", "New...") Then
        '해당 작업이 존재함을 알리는 파일을 만든다.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFileNo = FreeFile
        '파일을 연다.
        Open "C:\CALPLANS\TASKS\" & txtTaskTitle.Text For Output As #iFileNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFileNo, ""
        
        '파일을 닫는다.
        Close #iFileNo
        
        txtTaskTitle.Text = ""
        txtPercentage.Text = ""
        txtMemo.Text = ""
    End If
    
    LoadTasks
End Sub

Private Sub cmdTodaysPlan_Click()
    cmdPlanList_Click
End Sub

Private Sub cmdDelAllTodaysPlan_Click()
    On Error Resume Next
    If Confirm("삭제하시겠습니까?", "삭제", Me) Then
        If Confirm("복구 *불가능*합니다. 정말로 " & MonthView1.SelStart & "의 모든 일정을 삭제하시겠습니까?", "삭제", Me, , True) Then
            On Error Resume Next
            Shell "CMD /C RD /S /Q " & ChrW$(34) & "C:\CALPLANS\" & Split(MonthView1.SelStart, "-")(0) & "\" & Split(MonthView1.SelStart, "-")(1) & "\" & Split(MonthView1.SelStart, "-")(2) & ChrW$(34)
            Shell "COMMAND /C DELTREE /Y " & ChrW$(34) & "C:\CALPLANS\" & Split(MonthView1.SelStart, "-")(0) & "\" & Split(MonthView1.SelStart, "-")(1) & "\" & Split(MonthView1.SelStart, "-")(2) & ChrW$(34)
            
            MessageBox "삭제되었습니다.", "성공", Me
        End If
    End If
End Sub

Sub SetColor()
    Select Case GetSetting("Calendar", "Options", "BGColor", 0)
        Case 0
            Me.BackColor = &H8000000C
            ssRibbonMenu.BackColor = &H8000000C
            SSTab1.BackColor = &H8000000C
        Case 1
            Me.BackColor = &H8000000F
            ssRibbonMenu.BackColor = &H8000000F
            SSTab1.BackColor = &H8000000F
        Case 2
            Me.BackColor = &HFF&
            ssRibbonMenu.BackColor = &HFF&
            SSTab1.BackColor = &HFF&
        Case 3
            Me.BackColor = &HFFFF&
            ssRibbonMenu.BackColor = &HFFFF&
            SSTab1.BackColor = &HFFFF&
        Case 4
            Me.BackColor = &HC000&
            ssRibbonMenu.BackColor = &HC000&
            SSTab1.BackColor = &HC000&
        Case 5
            Me.BackColor = &HFFFF00
            ssRibbonMenu.BackColor = &HFFFF00
            SSTab1.BackColor = &HFFFF00
        Case 6
            Me.BackColor = &H808000
            ssRibbonMenu.BackColor = &H808000
            SSTab1.BackColor = &H808000
        Case 7
            Me.BackColor = &HC00000
            ssRibbonMenu.BackColor = &HC00000
            SSTab1.BackColor = &HC00000
        Case 8
            Me.BackColor = &H0&
            ssRibbonMenu.BackColor = &H0&
            SSTab1.BackColor = &H0&
    End Select
    
    ssRibbonMenu.BackColor = RGB(248, 164, 24)
    cmdRibbonFile.BackColor = RGB(248, 164, 24)
    cmdHelp.BackColor = RGB(248, 164, 24)
    cmdMnuOptions.BackColor = RGB(248, 164, 24)
    cmdMnuAbout.BackColor = RGB(248, 164, 24)
    Me.BackColor = RGB(255, 0, 255)
    
    ssTodaysPlan.BackColor = Me.BackColor
    cmdHelp.BackColor = RGB(248, 164, 24)
End Sub

Private Sub cmdUnselectAllDW_Click()
    Dim i As Integer
    For i = 0 To 6
        chkDayOfWeeks(i).Value = 0
    Next i
End Sub

Private Sub cmdRibbonFile_Click()
    PopupMenu mnuFile, , cmdRibbonFile.Left, cmdRibbonFile.Top + cmdRibbonFile.Height
End Sub

Private Sub Command1_Click()
    Form_Unload 0
End Sub

Private Sub Command2_Click()
    Me.WindowState = 1
End Sub

Private Sub Form_Load()
    'MsgBox DayOfWeek()
    'MessageBox PlayFair("dlfjs qkqhrkxdms sdfhuj", "ultra"), "3", Me
    
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbMagenta, 0&, LWA_COLORKEY

    If GetSetting("Calendar", "Options", "TP", 0) = 1 Then
        Me.Width = 8715
    End If
    
    If GetSetting("Calendar", "Options", "NoRibbon", 0) = 1 Then
        SSTab1.Top = 120
        ssTodaysPlan.Height = 4695
        lvTodaysPlan.Height = 3870
        cmdTltRef.Top = 4440
        Me.Height = 5900
        
        ssRibbonMenu.Visible = False
        cmdHelp.Visible = False
        cmdRibbonFile.Visible = False
        cmdMnuAbout.Visible = False
        cmdMnuOptions.Visible = False
        
        mnuDateMenu.Visible = True
        mnuFile.Visible = True
        mnuView.Visible = True
        mnuHelp.Visible = True
    End If
    
    tglCalWeekNum.Value = GetSetting("Calendar", "Options", "SWN", True)
    If GetSetting("Calendar", "Options", "SWN", "True") = "False" Then
        MonthView1.ShowWeekNumbers = "False"
    Else
        MonthView1.ShowWeekNumbers = "True"
    End If
    
    MonthView1.StartOfWeek = GetSetting("Calendar", "Options", "WSD", 0) + 1
    
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CONTACTS"
    MkDir "C:\CALPLANS\TASKS"
    
    cmbGroup.AddItem LoadLang("지정 안 함", "None")
    
    cmbGroup.ListIndex = 0
    
    Dim ty As Integer
    ty = Split(DateAdd("d", 1, Date), "-")(0)
    Dim tm As Integer
    tm = Split(DateAdd("d", 1, Date), "-")(1)
    Dim td As Integer
    td = Split(DateAdd("d", 1, Date), "-")(2)
    
    MkDir "C:\CALPLANS\" & ty
    MkDir "C:\CALPLANS\" & ty & "\" & tm
    MkDir "C:\CALPLANS\" & ty & "\" & tm & "\" & td

    Select Case UCase(Command)
        Case "/?"
            Select Case LoadLang(1, 2, 3)
                Case 1
                    MessageBox "일정관리자 풀그림을 시작합니다." & vbCrLf & vbCrLf & _
                           "    PLNMGR32.EXE [/R]" & vbCrLf & vbCrLf & _
                           "    /R  최소화된 상태로 시작합니다.", _
                           "스위치 도움말", Me

                Case 2
                    MessageBox "Starts the program." & vbCrLf & vbCrLf & _
                           "    PLNMGR32.EXE [/R]" & vbCrLf & vbCrLf & _
                           "    /R  Application window is minimized.", _
                           "Switch Guide", Me

                Case 3
                    MessageBox "Inicia el programa." & vbCrLf & vbCrLf & _
                           "    PLNMGR32.EXE [/R]" & vbCrLf & vbCrLf & _
                           "    /R  Haz que el programa sea transparente.", _
                           "Guia de comando", Me
            End Select
            End
        Case "/R"
            Me.WindowState = 1
        Case ""
        Case Else
            MessageBox LoadLang("스위치가 틀립니다", "Switch is wrong", "El comando no es valido.") & " - " & Command, LoadLang("오류", "Error", "Error"), Me, 16
            End
    End Select
    
    'mnuHelpAbout.Caption = App.Title & " 정보(&A)"
    
    'frmNotifyMgr.Show

    Me.Left = GetSetting("Calendar", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("Calendar", "Settings", "MainTop", 1000)
    
    Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab)
    Me.Caption = Me.Caption & " (" & MonthView1.Year & "년 " & MonthView1.Month & "월)"
    
    If GetSetting("Calendar", "Config", "FirstRun", "0") = "0" Then
        SaveSetting "Calendar", "Config", "FirstRun", "1"
        
        frmWizard.Show vbModal, Me
        
        If GetWinver(1) >= 6 And GetWinver(2) >= 1 Then
        Else
            'MessageBox LoadLang("컴퓨터가 Windows Vista 혹은 Windows XP 이하의 운영 체제를 실행하고 있습니다. 달력이 올바로 표시되지 않을 수 있습니다.", "Your PC is running Windows VIsta or earlier. The calendar may display incorrectly.", "La computadora esta ejecutando un sistema operativo de Windows Vista o Windows XP o inferior. Es posible que el calendario no se muestre correctamente."), LoadLang("경고", "Warning", "Advertencia"), Me, 48
        End If
        
        MessageBox LoadLang("컴퓨터를 시작할 때부터 알림을 받으려면 ", "Add ", "Agregue ") & "[" & ChrW$(34) & Dir1.Path & "\PLNMGR32.EXE" & ChrW$(34) & " /R]" & _
               LoadLang("(경로 복사됨) 바로가기를 시작프로그램에 추가하십시오.", "(Path Copied) to your startup program to be notified when you start your computer.", "(Ruta copiada) a su programa de inicio para recibir una notificacion cuando inicie su computadora."), LoadLang("알리미 활성화", "Tip", "Propina"), Me
        Clipboard.SetText ChrW$(34) & Dir1.Path & "\PLNMGR32.EXE" & ChrW$(34) & " /R"
    End If
    
    LoadContacts
    LoadTasks
    LoadAlarms
    
    SSTab1.Tab = GetSetting("Calendar", "Options", "StartPage", 0)
    
    If GetSetting("Calendar", "Options", "SST", True) = False Then
        SSTab1.Tab = GetSetting("Calendar", "Config", "LTB", GetSetting("Calendar", "Options", "StartPage", 0))
    End If
    
    SetColor
    
    MkDir "C:\CALPLANS\" & Format(Now, "YYYY") & "\" & Format(Now, "M") & "\" & Format(Now, "D")
    
    MkDir "C:\CALPLANS\" & ty & "\" & tm & "\" & td
    
    lvTmrPlans.Path = "C:\CALPLANS\" & ty & "\" & tm & "\" & td
    
    lvTodaysPlan.Path = "C:\CALPLANS\" & Format(Now, "YYYY") & "\" & Format(Now, "M") & "\" & Format(Now, "D")
    
    Dim DOWLS(6) As String
    DOWLS(0) = LoadLang("일", "S")
    DOWLS(1) = LoadLang("월", "M")
    DOWLS(2) = LoadLang("화", "T")
    DOWLS(3) = LoadLang("수", "W")
    DOWLS(4) = LoadLang("목", "T")
    DOWLS(5) = LoadLang("금", "F")
    DOWLS(6) = LoadLang("토", "S")
    
    Dim i As Variant
    For Each i In DOWLS
        lblDOW.Caption = lblDOW.Caption & i & vbNewLine & vbNewLine & vbNewLine
    Next i
    
    Dim j As Integer
    For j = 0 To txtPlannerTF.Count - 1
        txtPlannerTF(j).Text = GetSetting("Calendar", "Planner", CStr(j), "")
    Next j
    
    MkDir "C:\CALPLANS\CTGROUPS"
    
    lvGroupList.Path = "C:\CALPLANS\CTGROUPS"
    
    For i = 0 To lvGroupList.ListCount - 1
        cmbGroup.AddItem lvGroupList.List(i)
    Next i
    
    mnuFile.Caption = LoadLang("파일(&F)", "&File", "Archivo(&F)")
    mnuView.Caption = LoadLang("보기(&V)", "&View", "&Ver")
    mnuHelp.Caption = LoadLang("도움말(&H)", "&Help", "Ayuda(&H)")
    
    Me.mnuFileExit.Caption = LoadLang("비상문(&X)", "E&xit", "Salida(&X)")
    mnuFileProperties.Caption = LoadLang("일정 목록(&I)", "L&ist of Plans", "L&ista de horarios") & "..."
    mnuFilePlanBrowser.Caption = LoadLang("모든 일정/데이터 색인(&B)", "&Browse the Data", "Indice de datos(&B)") & "..."
    mnuFileSave.Caption = LoadLang("저장(&S)", "&Save", "Tienda(&S)")
    
    mnuViewStatusBar.Caption = LoadLang("상태 표시줄(&S)", "&Status Bar", "Barra de e&stado")
    mnuViewOptions.Caption = LoadLang("옵션(&O)", "&Options", "Ambientaci&on")
    
    mnuDateMenu.Caption = LoadLang("일정(&P)", "&Plans", "&Planes")
    mnuTodaysPlan.Caption = LoadLang("이날의 일정(&T)", "Selec&ted Date's Plans", "&Planes de la fecha seleccionada")
    mnuPlansClear.Caption = LoadLang("선택한 날짜의 일정 모두 삭제(&D)", "Clear selected &Date's Plans", "Borrar los planes &de la fecha seleccionada")
    
    ssRibbonMenu.TabCaption(1) = LoadLang("홈", "Home", "Inicio")
    ssRibbonMenu.TabCaption(2) = LoadLang("보기", "View", "Ver")
    ssRibbonMenu.TabCaption(3) = LoadLang("일정", "Plan", "Planes")
    
    cmdPlanList.Caption = LoadLang("일정 목록", "Plan List", "Lista de planes")
    cmdPlanIndex.Caption = LoadLang("데이터 색인", "Data Index", "Indice de datos")
    cmdEndPrg.Caption = LoadLang("끝내기", "Exit", "Salida")
    
    tglStatusBar.Caption = LoadLang("상태표시줄", "Status Bar", "Barra de estado")
    tglCalWeekNum.Caption = LoadLang("주 번호", "Week Number", "Numero de la semana")
    
    cmdTodaysPlan.Caption = LoadLang("이날의 일정", "Selected Day's Plans", "Planes del dia seleccionado")
    cmdDelAllTodaysPlan.Caption = LoadLang("이날의   일정 삭제", "Delete Plans", "Eliminar planes")
    
    cmdMnuAbout.ToolTipText = LoadLang("프로그램 정보", "About this application...")
    cmdMnuOptions.ToolTipText = LoadLang("환경 설정", "Settings...")
    cmdHelp.ToolTipText = LoadLang("도움말", "Help")
    
    cmdRibbonFile.Caption = LoadLang("파일", "File", "Archivo")
    
    ssTodaysPlan.TabCaption(0) = LoadLang("오늘 일정", "Today's Plans", "Los planes de hoy")
    ssTodaysPlan.TabCaption(1) = LoadLang("내일 일정", "Tomorrow's Plans", "Los planes de manana")
    
    SSTab1.TabCaption(0) = LoadLang("일정", "Plans", "Planes")
    SSTab1.TabCaption(1) = LoadLang("주소록", "Contacts", "Contactos")
    SSTab1.TabCaption(2) = LoadLang("할 일", "Tasks", "Tareas")
    SSTab1.TabCaption(3) = LoadLang("일과표", "Schedule", "Calendario")
    SSTab1.TabCaption(4) = LoadLang("알람", "Alarms", "Alarmas")
    
    cmdTltRef.Caption = LoadLang("갱신(&R)", "&Refresh", "Actualiza&r")
    
    Frame1.Caption = LoadLang("기본 정보", "Basic Information", "Informacion basica")
    Frame2.Caption = LoadLang("전화번호", "Phone Numbers", "Numeros de telefono")
    Frame3.Caption = LoadLang("기타 정보", "Other Informations", "Otra informacion")

    Label22.Caption = LoadLang("메모", "Note", "Nota") & ":"
    Label21.Caption = LoadLang("생일", "Birthday", "Cumpleanos") & ":"
    
     Label1.Caption = LoadLang("이름", "Name", "Nombre") & ":"
     Label2.Caption = LoadLang("휴대전화", "Cell-phone", "Celular") & ":"
     Label3.Caption = LoadLang("전자우편", "E-mail", "Correo electronico") & ":"
    Label17.Caption = LoadLang("그룹", "Group", "Grupo") & ":"
    Label12.Caption = LoadLang("우편번호", "Postal", "Postal") & ":"
    Label12.Caption = LoadLang("주소", "Address", "Direccion") & ":"
     Label4.Caption = LoadLang("집", "Home", "Casa") & ":"
     Label5.Caption = LoadLang("회사", "Company", "Empresa") & ":"
     Label6.Caption = LoadLang("팩스", "Fax", "Fax") & ":"
     Label7.Caption = LoadLang("기타", "Other", "Otros") & ":"
    
    cmdSaveContact.Caption = LoadLang("저장(&S)", "&Save", "Tienda(&S)")
    cmdDelContact.Caption = LoadLang("삭제(&D)", "&Delete", "Eliminar(&D)")
    cmdDeleteAllContacts.Caption = LoadLang("모두 삭제(&E)", "Cl&ear contatcs", "Eliminar todo(&E)")
    cmdResetFields.Caption = LoadLang("내용 초기화(&R)", "&Reset Fields", "&Agregar")
    
    Frame4.Caption = LoadLang("할 일 정보", "Task Information")
    Label10.Caption = LoadLang("제목", "Title") & ":"
    Label8.Caption = LoadLang("완료율", "Percent Complete") & ":"
    Label14.Caption = LoadLang("중요도", "Importance") & ":"
    Label16.Caption = LoadLang("참여자", "Participants") & ":"
    Label11.Caption = LoadLang("메모", "Note") & ":"
    
    cmdSaveTask.Caption = LoadLang("저장(&S)", "&Save", "Tienda(&S)")
    cmdDelTask.Caption = LoadLang("삭제", "&Delete", "Eliminar(&D)")
    cmdDeleteAllTasks.Caption = LoadLang("모두 삭제(&E)", "Cl&ear Tasks", "Eliminar todo(&E)")
    
    sdcmdSavePlanner.Caption = LoadLang("저장(&S)", "&Save", "Tienda(&S)")
    
    Label18.Caption = LoadLang("이름", "Name", "Nombre") & ":"
    Label19.Caption = LoadLang("시간", "Time", "Tiempo") & ":"
    Label20.Caption = LoadLang("메모", "Note", "Nota") & ":"
    Frame5.Caption = LoadLang("요일", "-", "-")
    
    chkDayOfWeeks(0).Caption = LoadLang("일요일", "Sunday", "Domingo")
    chkDayOfWeeks(1).Caption = LoadLang("월요일", "Monday", "Lunes")
    chkDayOfWeeks(2).Caption = LoadLang("화요일", "Tuesday", "Martes")
    chkDayOfWeeks(3).Caption = LoadLang("수요일", "Wednesday", "Miercoles")
    chkDayOfWeeks(4).Caption = LoadLang("목요일", "Thursday", "Jueves")
    chkDayOfWeeks(5).Caption = LoadLang("금요일", "Friday", "Viernes")
    chkDayOfWeeks(6).Caption = LoadLang("토요일", "Saturday", "Sabado")
    
    cmdSelectAllDW.Caption = LoadLang("모두 선택(&A)", "Select &All", "Seleccion&ar todo")
    cmdUnselectAllDW.Caption = LoadLang("선택 해제(&L)", "Dese&lect All", "Dese&leccionar todo")
    cmdRelectAllDW.Caption = LoadLang("선택 반전(&I)", "&Invert", "&Invertir seleccion")
    
    cmdResetAF.Caption = LoadLang("초기화(&R)", "&Reset Fields", "&Restablecer")
    cmdDeleteAlarm.Caption = LoadLang("삭제(&D)", "&Delete", "Eliminar(&D)")
    cmdSaveAlarm.Caption = LoadLang("추가(&A)", "&Add", "&Agregar")
    
    Me.Caption = LoadLang(App.Title, "Plan Manager 3")
    
    MonthView1.Value = Split(Format(Now, "YYYY-M-D"), "-")(0) & "-" & Split(Format(Now, "YYYY-M-D"), "-")(1) & "-" & Split(Format(Now, "YYYY-M-D"), "-")(2)

    Me.Show
    frmTip.Show vbModal, Me
    
    Me.BorderStyle = 0
    Me.Caption = Me.Caption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "1"
End Sub

Private Sub Image1_Click()
    cmdTltRef_Click
End Sub

Private Sub lvAlarms_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    If Item.SubItems(1) = LoadLang("새 알람 추가...", "New...") Then
        ClearAlarmFields
    Else
        cmdDeleteAlarm.Enabled = True
        
        txtAlarmTitle.Text = Item.SubItems(1)
        txtTimeHrs.Text = Split(GetSetting("Calendar", "Alarms", txtAlarmTitle.Text & "Time", "00:00"), ":")(0)
        txtTimeMin.Text = Split(GetSetting("Calendar", "Alarms", txtAlarmTitle.Text & "Time", "00:00"), ":")(1)
        
        On Error Resume Next
        Dim i As Integer
        For i = 0 To 6
            chkDayOfWeeks(i).Value = GetSetting("Calendar", "Alarms", txtAlarmTitle.Text & "W" & CStr(i), 0)
        Next i
        
        txtAlarmMemo.Text = GetSetting("Calendar", "Alarms", txtAlarmTitle.Text & "Memo", "")
        
        txtAlarmTitle.Enabled = False
    End If
End Sub

Private Sub lvTodaysPlan_DblClick()
    On Error Resume Next
End Sub

Private Sub menuhover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    timHidemenu.Enabled = False
    timHidemenu.Enabled = True
    ssRibbonMenu.Visible = True
    Me.BorderStyle = 1
    cmdRibbonFile.Visible = -1
    cmdMnuOptions.Visible = -1
    cmdHelp.Visible = -1
    cmdMnuAbout.Visible = -1
End Sub

Private Sub mnuPlansClear_Click()
    cmdDelAllTodaysPlan_Click
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub sdcmdSavePlanner_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To txtPlannerTF.Count - 1
        SaveSetting "Calendar", "Planner", CStr(i), txtPlannerTF(i).Text
    Next i
        
    lblStatus.Caption = "저장되었습니다."
    Sleep 1000
    lblStatus.Caption = ""
End Sub

Private Sub ssRibbonMenu_Click(PreviousTab As Integer)
    If ssRibbonMenu.Tab = 0 Then
        ssRibbonMenu.Tab = PreviousTab
    End If
End Sub

Private Sub timAlarmChecker_Timer()
    '알람을 찾는다.
    On Error Resume Next
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\ALARMS"
    
    lvAlarmList.Path = "C:\CALPLANS\ALARMS"
    lvAlarmList.Refresh
    
    Dim Alarm As Integer
    Dim ttt As String
    
    For Alarm = 0 To lvAlarmList.ListCount - 1
        ttt = Format(Now, "hh:mm")
        
        If ttt = GetSetting("Calendar", "Alarms", lvAlarmList.List(Alarm) & "Time", "00:00") Then
            If GetSetting("Calendar", "NotifiedAlarms", lvAlarmList.List(Alarm), "abc") = "abc" Then
                If GetSetting("Calendar", "Alarms", lvAlarmList.List(Alarm) & "W" & CStr(DayOfWeek()), 0) = 1 Then
                    SaveSetting "Calendar", "NotifiedAlarms", lvAlarmList.List(Alarm), "1"
                    frmAlarm.lblCaption = lvAlarmList.List(Alarm)
                    frmAlarm.txtAlarmMemo = GetSetting("Calendar", "Alarms", lvAlarmList.List(Alarm) & "Memo", "")
                    frmAlarm.Show vbModal, Me
                End If
            End If
        End If
    Next Alarm
End Sub

Private Sub Timer1_Timer()
    '오늘의 일정을 찾는다.
    On Error Resume Next
    
    Dim yy As Integer
    Dim mm As Integer
    Dim dd As Integer
    
    yy = Format(Now, "YYYY")
    mm = Format(Now, "M")
    dd = Format(Now, "D")
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\" & yy
    MkDir "C:\CALPLANS\" & yy & "\" & mm
    MkDir "C:\CALPLANS\" & yy & "\" & mm & "\" & dd
    
    
    
    lvTodaysPlans.Path = "C:\CALPLANS\" & yy & "\" & mm & "\" & dd
    
    lvTodaysPlans.Refresh
    Dim Plan As Integer
    Dim ttt As Integer
    
    For Plan = 0 To lvTodaysPlans.ListCount - 1
        ttt = CInt(Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(0) & Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(1)) - CInt(Format(Now, "hhmm"))
        '현재시각과 일정시각과의 차이가 10분 미만이면 알림을 띄운다.
        'MsgBox Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(0) & Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(1) & " " & Format(Now, "hhmm") & " " & ttt
        If ttt < 10 And ttt >= -1 Then
            '띄운 적이 없으면 알림
            If GetSetting("Calendar", "NotifiedPlans\" & yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan), "abc") = "abc" Then
                'MsgBox 3
                frmReminder.yy = yy
                frmReminder.mm = mm
                frmReminder.dd = dd
                frmReminder.lblTitle.Caption = lvTodaysPlans.List(Plan)
                frmReminder.lblLoca.Caption = GetSetting("Calendar", yy & mm & dd, lvTodaysPlans.List(Plan) & "Location", "주소 불분명")
                frmReminder.txtContent.Text = GetSetting("Calendar", yy & mm & dd, lvTodaysPlans.List(Plan) & "Cont", "")
                frmReminder.Show
                'SysTray.ShowBalloonTip lvTodaysPlans.List(Plan) & " 일정 시작까지 10분보다 적게 남았습니다. 준비하십시오.", beInformation, "일정관리자"
                'Beep 950, 5
            End If
        End If
    Next Plan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Confirm(LoadLang("일정관리자를 닫으면 예정 일정 알림을 받지 않습니다.", "You will not be notified when you close the program.", "No se le notificara sobre los planes cuando se cierre el programa."), LoadLang("경고", "Warning", "Advertencia"), Me, 48) = True Then
        Dim i As Integer
        
        SaveSetting "Calendar", "Config", "LTB", SSTab1.Tab
        
        If Me.WindowState <> vbMinimized Then
            SaveSetting "Calendar", "Settings", "MainLeft", Me.Left
            SaveSetting "Calendar", "Settings", "MainTop", Me.Top
        End If
        
        End
    Else
        Cancel = 1
        Exit Sub
    End If
    
    'Cancel = 1
    'Me.Hide
    'frmNotifyMgr.Show
End Sub

Private Sub lvContacts_Click()
    On Error Resume Next
    
    If lvContacts.List(lvContacts.ListIndex) = LoadLang("새 연락처 추가...", "New...") Then
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
        
        txtBMonth.Text = ""
        txtBYear.Text = ""
        txtBDay.Text = ""
        
        cmbGroup.ListIndex = 0
        
        cmdDelContact.Enabled = False
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
        
        txtContent.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Content", "")
        
        txtBYear.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "BY", "")
        txtBMonth.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "BM", "")
        txtBDay.Text = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "BD", "")
        
        Dim i As Integer
        
        For i = 0 To cmbGroup.ListCount - 1
            If cmbGroup.List(i) = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Group", "") Then
                cmbGroup.ListIndex = i
                Exit For
            End If
        Next i
        
        cmdDelContact.Enabled = True
        
        'Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (" & txtName.Text & ")"
    End If
End Sub

Private Sub lvTasks_Click()
    If lvTasks.List(lvTasks.ListIndex) = LoadLang("새 작업 추가...", "New...") Then
        cmdDelTask.Enabled = False
    Else
        cmdDelTask.Enabled = True
    End If
    
    txtTaskTitle.Text = lvTasks.List(lvTasks.ListIndex)
    txtPercentage.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Perc", "")
    txtMemo.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Memo", "")
    txtPart.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Part", "")
    txtImpt.Text = GetSetting("Calendar", "Tasks", lvTasks.List(lvTasks.ListIndex) & "Impt", "")
End Sub

Private Sub lvTasks_ItemCheck(Item As Integer)
    If lvTasks.List(Item) <> LoadLang("새 작업 추가...", "New...") Then
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
        If FileExists(Dir1.Path & "\PLNMGR32.HLP") = False Then
            MessageBox "도움말 파일을 찾을 수 없습니다. 풀그림 실행화일 경로에 PLNMGR32.HLP이 있는지 확인하십시오. 없으면 다시 설치하거나 깃허브에서 받아 복사하십시오.", "도움말", Me, 16
            Exit Sub
        End If
        
        nRet = OSWinHelp(Me.hwnd, Dir1.Path & "\PLNMGR32.HLP", 261, 0)
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
        If FileExists(Dir1.Path & "\PLNMGR32.HLP") = False Then
            MessageBox "도움말 파일을 찾을 수 없습니다. 풀그림 실행화일 경로에 PLNMGR32.HLP이 있는지 확인하십시오. 없으면 다시 설치하거나 깃허브에서 받아 복사하십시오.", "도움말", Me, 16
            Exit Sub
        End If
        
        nRet = OSWinHelp(Me.hwnd, Dir1.Path & "\PLNMGR32.HLP", 3, 0)
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
    
    If sbStatusBar.Visible Then
        Me.Height = 7080
    Else
        Me.Height = 6810
    End If
End Sub

'Private Sub mnuViewToolbar_Click()
    'mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
'End Sub

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
    Unload Me
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

'Private Sub mnuFilePageSetup_Click()
'    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "페이지 설정"
'        .CancelError = True
'        .ShowPrinter
'    End With
'
'End Sub

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

'Private Sub mnuFileOpen_Click()
'    Dim sFile As String
'
'
'    With dlgCommonDialog
'        .DialogTitle = "열기"
'        .CancelError = False
'        '작업: Common Dialog 컨트롤의 플래그와 특성을 설정합니다.
'        .Filter = "모든 파일(*.*)|*.*"
'        .ShowOpen
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    '작업: 코드를 추가하여 열려 있는 파일을 처리합니다.
'
'End Sub

Private Sub mnuFileNew_Click()
    '작업: 'mnuFileNew_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileNew_Click' 코드를 추가하십시오."
End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuTodaysPlan.Caption = MonthView1.SelStart & "의 일정"
        PopupMenu mnuDateMenu
    End If
End Sub

Private Sub ssRibbonMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab > 0 Then
        mnuFileBar0.Visible = True
        mnuFileSave.Visible = True
    Else
        mnuFileBar0.Visible = False
        mnuFileSave.Visible = False
    End If
End Sub

Private Sub tglCalWeekNum_Click()
    If MonthView1.ShowWeekNumbers = False Then
        MonthView1.ShowWeekNumbers = True
    Else
        MonthView1.ShowWeekNumbers = False
    End If
    
    SaveSetting "Calendar", "Options", "SWN", tglCalWeekNum.Value
End Sub

Private Sub tglStatusBar_Click()
    mnuViewStatusBar_Click
End Sub

Private Sub timHidemenu_Timer()
    ssRibbonMenu.Visible = 0
    Me.BorderStyle = 0
    cmdRibbonFile.Visible = 0
    cmdMnuOptions.Visible = 0
    cmdHelp.Visible = 0
    cmdMnuAbout.Visible = 0
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

Private Sub LoadAlarms()
    lvAlarms.ColumnHeaders.Clear
    
    On Error Resume Next
    
    lvAlarms.ListItems.Clear
    lvAlarmFiles.Refresh
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\ALARMS"
    
    lvAlarmFiles.Path = "C:\CALPLANS\ALARMS"
    
    lvAlarms.ColumnHeaders.Add , , LoadLang("시간", "Time"), 350
    lvAlarms.ColumnHeaders.Add , , LoadLang("이름", "Name"), 1400

    lvAlarms.ListItems.Add , , "--:--"
    lvAlarms.ListItems(1).SubItems(1) = LoadLang("새 알람 추가...", "New...")
    
    Dim Alarm As Integer
    Dim Title As String
    Dim Time As String
    
    For Alarm = 0 To lvAlarmFiles.ListCount - 1
        Title = lvAlarmFiles.List(Alarm)
        Time = GetSetting("Calendar", "Alarms", Title & "Time", "00:00")

        lvAlarms.ListItems.Add , , Time
        lvAlarms.ListItems(Alarm + 2).SubItems(1) = Title
    Next Alarm
End Sub

Private Sub txtTimeHrs_Change()
    On Error Resume Next
    If Len(txtTimeHrs.Text) = 2 Or (txtTimeHrs.Text >= 3 And txtTimeHrs.Text <= 9) Then
        txtTimeMin.SetFocus '시 입력 칸을 채우면 다음 칸을 활성화한다.
    End If
End Sub

Private Sub txtTimeMin_Change()
    If txtTimeMin.Text = "" Then txtTimeHrs.SetFocus
End Sub

