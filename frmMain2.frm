VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '���� ����
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   11385
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   6960
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   229
      Top             =   1230
      Width           =   2895
   End
   Begin VB.CheckBox chkDuplicateContact 
      Caption         =   "Check1"
      Height          =   255
      Left            =   720
      TabIndex        =   228
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton cmdMnuAbout 
      BackColor       =   &H8000000C&
      Caption         =   "+"
      Height          =   375
      Left            =   600
      TabIndex        =   227
      ToolTipText     =   "���α׷� ������ �����ݴϴ�."
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdTltRef 
      Caption         =   "����(&R)"
      Height          =   300
      Left            =   12000
      TabIndex        =   8
      ToolTipText     =   "������ ��������� �����մϴ�."
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin TabDlg.SSTab ssSidebar 
      Height          =   4815
      Left            =   8760
      TabIndex        =   217
      Top             =   1680
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   10
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   2
      TabMaxWidth     =   2
      WordWrap        =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain2.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ssTodaysPlan"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMain2.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ssFavorites"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMain2.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ssAlmostFinished"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmMain2.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmMain2.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmMain2.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "frmMain2.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmMain2.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "frmMain2.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "frmMain2.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      Begin TabDlg.SSTab ssFavorites 
         Height          =   4815
         Left            =   0
         TabIndex        =   222
         Top             =   0
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "���ã��"
         TabPicture(0)   =   "frmMain2.frx":055A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lvFavoriteContacts"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvContactFiles2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "�׷�"
         TabPicture(1)   =   "frmMain2.frx":0576
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvGroupContacts"
         Tab(1).Control(1)=   "cmbIndexContactGroups"
         Tab(1).Control(2)=   "lvContactFiles3"
         Tab(1).ControlCount=   3
         Begin VB.FileListBox lvContactFiles3 
            Height          =   450
            Left            =   -74280
            TabIndex        =   233
            Top             =   2280
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.FileListBox lvContactFiles2 
            Height          =   450
            Left            =   840
            TabIndex        =   232
            Top             =   1920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cmbIndexContactGroups 
            Height          =   300
            Left            =   -74880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   231
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox lvGroupContacts 
            Height          =   3660
            Left            =   -74880
            TabIndex        =   230
            Top             =   720
            Width           =   1935
         End
         Begin VB.ListBox lvFavoriteContacts 
            Height          =   4200
            Left            =   120
            TabIndex        =   223
            Top             =   360
            Width           =   1935
         End
      End
      Begin TabDlg.SSTab ssTodaysPlan 
         Height          =   4815
         Left            =   -75000
         TabIndex        =   218
         Top             =   0
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BackColor       =   -2147483636
         TabCaption(0)   =   "���� ����"
         TabPicture(0)   =   "frmMain2.frx":0592
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lvTodaysPlan"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvTodaysPlans"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "���� ����"
         TabPicture(1)   =   "frmMain2.frx":05AE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvTmrPlans"
         Tab(1).ControlCount=   1
         Begin VB.FileListBox lvTmrPlans 
            Height          =   4410
            Left            =   -74880
            TabIndex        =   221
            Top             =   360
            Width           =   1935
         End
         Begin VB.FileListBox lvTodaysPlans 
            Height          =   270
            Left            =   1560
            TabIndex        =   220
            Top             =   1320
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.FileListBox lvTodaysPlan 
            Height          =   4410
            Left            =   120
            TabIndex        =   219
            Top             =   360
            Width           =   1935
         End
      End
      Begin TabDlg.SSTab ssAlmostFinished 
         Height          =   4815
         Left            =   -75000
         TabIndex        =   224
         Top             =   0
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "�Ϸ�Ǿ�� �۾�"
         TabPicture(0)   =   "frmMain2.frx":05CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "List2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.ListBox List2 
            Height          =   4200
            Left            =   120
            TabIndex        =   225
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.Timer datetime 
      Interval        =   100
      Left            =   2400
      Top             =   1200
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox hider2 
      Appearance      =   0  '���
      BorderStyle     =   0  '����
      Enabled         =   0   'False
      Height          =   4815
      Left            =   8640
      TabIndex        =   149
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox hider1 
      Appearance      =   0  '���
      BorderStyle     =   0  '����
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   148
      Top             =   6855
      Width           =   8535
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   300
      Left            =   8340
      TabIndex        =   147
      Top             =   6525
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      Height          =   300
      Left            =   255
      TabIndex        =   146
      Top             =   6525
      Width           =   255
   End
   Begin VB.Timer hideshow 
      Interval        =   100
      Left            =   2880
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   12720
      Top             =   1560
   End
   Begin VB.FileListBox lvGroupList 
      Height          =   270
      Left            =   10560
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer timAlarmChecker 
      Interval        =   10000
      Left            =   12960
      Top             =   1800
   End
   Begin VB.FileListBox lvAlarmList 
      Height          =   270
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pbxRibbonBackground 
      Height          =   975
      Left            =   12720
      Picture         =   "frmMain2.frx":05E6
      ScaleHeight     =   915
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.PictureBox pbxTodaysPlanTab 
      Height          =   1335
      Left            =   12840
      Picture         =   "frmMain2.frx":350F8
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdRibbonFile 
      BackColor       =   &H8000000C&
      Caption         =   "cmdRibbonFile"
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1105
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "��"
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      ToolTipText     =   "���α׷��� ���򸻰� ���õ� �׸��Դϴ�."
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
      Caption         =   "_"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      ToolTipText     =   "���α׷��� ���򸻰� ���õ� �׸��Դϴ�."
      Top             =   1200
      Width           =   375
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '�Ʒ� ����
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   8235
      Visible         =   0   'False
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14896
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2021-07-23"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "���� 12:25"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssRibbonMenu 
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   548
      TabMaxWidth     =   1940
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      MouseIcon       =   "frmMain2.frx":5C2BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain2.frx":5C2D6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "home"
      TabPicture(1)   =   "frmMain2.frx":5C2F2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "timHidemenu"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "view"
      TabPicture(2)   =   "frmMain2.frx":5C744
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tglCalWeekNum"
      Tab(2).Control(1)=   "tglStatusBar"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "plans"
      TabPicture(3)   =   "frmMain2.frx":5CB96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   410
         Width           =   2415
         Begin VB.CommandButton cmdPlanList 
            Caption         =   "���� ���"
            Height          =   600
            Left            =   120
            Picture         =   "frmMain2.frx":5CFE8
            Style           =   1  '�׷���
            TabIndex        =   18
            ToolTipText     =   "ǥ���� ��¥�� ���� ����� ǥ���մϴ�."
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdPlanIndex 
            Caption         =   "������ ����"
            Height          =   615
            Left            =   1200
            Picture         =   "frmMain2.frx":5D42A
            Style           =   1  '�׷���
            TabIndex        =   17
            ToolTipText     =   "�ּҷ�, ���� ��ü����Դϴ�."
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  '����
         Height          =   615
         Left            =   2640
         TabIndex        =   14
         Top             =   410
         Width           =   1335
         Begin VB.CommandButton cmdEndPrg 
            Caption         =   "������"
            Height          =   615
            Left            =   120
            Picture         =   "frmMain2.frx":5D86C
            Style           =   1  '�׷���
            TabIndex        =   15
            ToolTipText     =   "���α׷��� �����ϴ�."
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  '����
         Height          =   615
         Left            =   -74880
         TabIndex        =   11
         Top             =   410
         Width           =   2415
         Begin VB.CommandButton cmdTodaysPlan 
            Caption         =   "�̳��� ����"
            Height          =   615
            Left            =   120
            Picture         =   "frmMain2.frx":5DCAE
            Style           =   1  '�׷���
            TabIndex        =   13
            ToolTipText     =   "ǥ���� ��¥�� ���� ����� ǥ���մϴ�."
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelAllTodaysPlan 
            Caption         =   "�̳���   ���� ����"
            Height          =   615
            Left            =   1320
            Picture         =   "frmMain2.frx":5E0F0
            Style           =   1  '�׷���
            TabIndex        =   12
            ToolTipText     =   "������ ���� ������ ��� �����մϴ�."
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Timer timHidemenu 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   12360
         Top             =   240
      End
      Begin MSForms.ToggleButton tglCalWeekNum 
         Height          =   615
         Left            =   -74880
         TabIndex        =   20
         ToolTipText     =   "�޷¿��� ���� ��ȣ�� ǥ���ϰų� ����ϴ�."
         Top             =   360
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1931;1085"
         Value           =   "1"
         Caption         =   "�� ��ȣ"
         FontName        =   "����"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tglStatusBar 
         Height          =   840
         Left            =   -64800
         TabIndex        =   19
         ToolTipText     =   "����ǥ������ ǥ���ϰų� ����ϴ�."
         Top             =   375
         Visible         =   0   'False
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1931;1482"
         Value           =   "1"
         Caption         =   "����ǥ����"
         FontName        =   "����"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5025
      Left            =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   8864
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   10
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   582
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      MouseIcon       =   "frmMain2.frx":5E532
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "frmMain2.frx":5E54E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Dir1"
      Tab(0).Control(1)=   "MonthView1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�ּҷ�"
      TabPicture(1)   =   "frmMain2.frx":5E9A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdResetFields"
      Tab(1).Control(1)=   "cmdDeleteAllContacts"
      Tab(1).Control(2)=   "cmdDelContact"
      Tab(1).Control(3)=   "lvContactFiles"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "cmdSaveContact"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(8)=   "lvContacts"
      Tab(1).Control(9)=   "cmdDuplicateContact"
      Tab(1).Control(10)=   "chkFavorite"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "�� ��"
      TabPicture(2)   =   "frmMain2.frx":5EDF2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvTasks"
      Tab(2).Control(1)=   "cmdSaveTask"
      Tab(2).Control(2)=   "cmdDelTask"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "lvTaskFiles"
      Tab(2).Control(5)=   "cmdDeleteAllTasks"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "�ϰ�ǥ"
      TabPicture(3)   =   "frmMain2.frx":5F244
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
      TabCaption(4)   =   "�˶�"
      TabPicture(4)   =   "frmMain2.frx":5F55E
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label20"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label19"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label18"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lvAlarms"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lvAlarmFiles"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtAlarmMemo"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdDeleteAlarm"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdSaveAlarm"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cmdResetAF"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Frame5"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtTimeMin"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtTimeHrs"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "txtAlarmTitle"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "chkDayOfWeeks(0)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "chkDayOfWeeks(1)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "chkDayOfWeeks(2)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "chkDayOfWeeks(3)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "chkDayOfWeeks(4)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "chkDayOfWeeks(5)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "chkDayOfWeeks(6)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "cmdRelectAllDW"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "cmdUnselectAllDW"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "cmdSelectAllDW"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).ControlCount=   23
      TabCaption(5)   =   " �޸�"
      TabPicture(5)   =   "frmMain2.frx":5F9B0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text1"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "�׸���"
      TabPicture(6)   =   "frmMain2.frx":5FE02
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "����"
      TabPicture(7)   =   "frmMain2.frx":60254
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command10"
      Tab(7).Control(1)=   "Command9"
      Tab(7).Control(2)=   "Command6"
      Tab(7).Control(3)=   "Command7"
      Tab(7).Control(4)=   "Command8"
      Tab(7).Control(5)=   "cmdCalcOperator(2)"
      Tab(7).Control(6)=   "cmdCalcAdvancedOperator(15)"
      Tab(7).Control(7)=   "cmdCalcAdvancedOperator(14)"
      Tab(7).Control(8)=   "cmdCalcAdvancedOperator(13)"
      Tab(7).Control(9)=   "cmdCalcAdvancedOperator(12)"
      Tab(7).Control(10)=   "cmdCalcAdvancedOperator(11)"
      Tab(7).Control(11)=   "cmdCalcAdvancedOperator(9)"
      Tab(7).Control(12)=   "cmdCalcAdvancedOperator(8)"
      Tab(7).Control(13)=   "cmdCalcAdvancedOperator(7)"
      Tab(7).Control(14)=   "cmdCalcAdvancedOperator(6)"
      Tab(7).Control(15)=   "cmdCalcAdvancedOperator(5)"
      Tab(7).Control(16)=   "cmdCalcAdvancedOperator(4)"
      Tab(7).Control(17)=   "cmdCalcAdvancedOperator(3)"
      Tab(7).Control(18)=   "cmdCalcAdvancedOperator(2)"
      Tab(7).Control(19)=   "cmdCalcAdvancedOperator(1)"
      Tab(7).Control(20)=   "cmdCalcAdvancedOperator(0)"
      Tab(7).Control(21)=   "cmdCalcAdvancedOperator(10)"
      Tab(7).Control(22)=   "cmdCalcCalculate"
      Tab(7).Control(23)=   "cmdCalcPoint"
      Tab(7).Control(24)=   "cmdCalcOperator(4)"
      Tab(7).Control(25)=   "cmdCalcOperator(3)"
      Tab(7).Control(26)=   "cmdCalcOperator(1)"
      Tab(7).Control(27)=   "cmdCalcNumbers(9)"
      Tab(7).Control(28)=   "cmdCalcNumbers(8)"
      Tab(7).Control(29)=   "cmdCalcNumbers(7)"
      Tab(7).Control(30)=   "cmdCalcNumbers(6)"
      Tab(7).Control(31)=   "cmdCalcNumbers(5)"
      Tab(7).Control(32)=   "cmdCalcNumbers(4)"
      Tab(7).Control(33)=   "cmdCalcNumbers(3)"
      Tab(7).Control(34)=   "cmdCalcNumbers(1)"
      Tab(7).Control(35)=   "cmdCalcNumbers(2)"
      Tab(7).Control(36)=   "cmdCalcNumbers(0)"
      Tab(7).Control(37)=   "Text2"
      Tab(7).ControlCount=   38
      TabCaption(8)   =   "Ÿ�̸�"
      TabPicture(8)   =   "frmMain2.frx":606A6
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Shape1"
      Tab(8).Control(1)=   "Shape2"
      Tab(8).Control(2)=   "lblTimerH"
      Tab(8).Control(3)=   "lblTimerM"
      Tab(8).Control(4)=   "lblTimerS"
      Tab(8).Control(5)=   "lblTimerSD"
      Tab(8).Control(6)=   "Label27"
      Tab(8).Control(7)=   "Label28"
      Tab(8).Control(8)=   "Label29"
      Tab(8).Control(9)=   "lbltsw"
      Tab(8).Control(10)=   "ssTimerType"
      Tab(8).Control(11)=   "cmdTimerStart"
      Tab(8).Control(12)=   "txtTimerLog"
      Tab(8).Control(13)=   "cmdTimerPause"
      Tab(8).Control(14)=   "cmdTImerStop"
      Tab(8).Control(15)=   "Command14"
      Tab(8).Control(16)=   "timTImer"
      Tab(8).Control(17)=   "pbTimer"
      Tab(8).Control(18)=   "Frame9"
      Tab(8).Control(19)=   "cmdSetTimerTime"
      Tab(8).Control(20)=   "chkEnablesw"
      Tab(8).ControlCount=   21
      TabCaption(9)   =   "Tab 9"
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      Begin VB.CommandButton cmdSelectAllDW 
         Caption         =   "cmdSelectAllDW"
         Height          =   320
         Left            =   4200
         TabIndex        =   244
         Top             =   1785
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnselectAllDW 
         Caption         =   "cmdUnselectAllDW"
         Height          =   320
         Left            =   5520
         TabIndex        =   243
         Top             =   1785
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelectAllDW 
         Caption         =   "cmdRelectAllDW"
         Height          =   320
         Left            =   6840
         TabIndex        =   242
         Top             =   1785
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "�����"
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   241
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "�ݿ���"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   240
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "�����"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   239
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "������"
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   238
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "ȭ����"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   237
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "������"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   236
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDayOfWeeks 
         Caption         =   "�Ͽ���"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   235
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkFavorite 
         Caption         =   "���ã��(&F)"
         Height          =   255
         Left            =   -72960
         TabIndex        =   234
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdDuplicateContact 
         Caption         =   "����(&U)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   226
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chkEnablesw 
         Caption         =   "Ÿ�̸� ���� ��    �����ġ ����"
         Height          =   375
         Left            =   -70680
         TabIndex        =   215
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSetTimerTime 
         Caption         =   "����"
         Height          =   375
         Left            =   -71760
         TabIndex        =   210
         Top             =   3420
         Width           =   855
      End
      Begin VB.Frame Frame9 
         Caption         =   "�ð� ����"
         Height          =   615
         Left            =   -74760
         TabIndex        =   204
         Top             =   3240
         Width           =   3975
         Begin VB.TextBox txtSetH 
            Alignment       =   1  '������ ����
            Height          =   270
            Left            =   120
            TabIndex        =   205
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin ComCtl2.UpDown UpDown3 
            Height          =   270
            Left            =   480
            TabIndex        =   211
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtSetH"
            BuddyDispid     =   196656
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   495
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSetM 
            Alignment       =   1  '������ ����
            Height          =   270
            Left            =   840
            TabIndex        =   208
            Text            =   "0"
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox txtSetS 
            Alignment       =   1  '������ ����
            Height          =   270
            Left            =   1560
            TabIndex        =   207
            Text            =   "0"
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox txtSetSD 
            Alignment       =   1  '������ ����
            Height          =   270
            Left            =   2280
            TabIndex        =   206
            Text            =   "0"
            Top             =   240
            Width           =   420
         End
         Begin ComCtl2.UpDown UpDown4 
            Height          =   270
            Left            =   1200
            TabIndex        =   212
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtSetM"
            BuddyDispid     =   196657
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   495
            Max             =   59
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown5 
            Height          =   270
            Left            =   1920
            TabIndex        =   213
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtSetS"
            BuddyDispid     =   196658
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   495
            Max             =   59
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown6 
            Height          =   270
            Left            =   2640
            TabIndex        =   214
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "txtSetSD"
            BuddyDispid     =   196659
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   495
            Max             =   9
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label30 
            Caption         =   "          :           :           ."
            Height          =   255
            Left            =   150
            TabIndex        =   209
            Top             =   270
            Width           =   2415
         End
      End
      Begin ComctlLib.ProgressBar pbTimer 
         Height          =   255
         Left            =   -74760
         TabIndex        =   203
         Top             =   3960
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Timer timTImer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   -70200
         Top             =   2880
      End
      Begin VB.CommandButton Command14 
         Caption         =   "���"
         Height          =   375
         Left            =   -70440
         TabIndex        =   202
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdTImerStop 
         Caption         =   "�ʱ�ȭ"
         Height          =   375
         Left            =   -71880
         TabIndex        =   201
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdTimerPause 
         Caption         =   "�Ͻ�����"
         Height          =   375
         Left            =   -73320
         TabIndex        =   200
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtTimerLog 
         Height          =   3615
         Left            =   -68520
         MultiLine       =   -1  'True
         TabIndex        =   199
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdTimerStart 
         Caption         =   "����"
         Height          =   375
         Left            =   -74760
         TabIndex        =   198
         Top             =   2160
         Width           =   1335
      End
      Begin TabDlg.SSTab ssTimerType 
         Height          =   375
         Left            =   -74880
         TabIndex        =   190
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Ÿ�̸�"
         TabPicture(0)   =   "frmMain2.frx":60AF8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "�����ġ"
         TabPicture(1)   =   "frmMain2.frx":60B14
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�޸� �����"
         Height          =   375
         Left            =   -71760
         TabIndex        =   189
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "�޸� �б�"
         Height          =   375
         Left            =   -73320
         TabIndex        =   188
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "�޸� ����"
         Height          =   375
         Left            =   -74880
         TabIndex        =   187
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�ʱ�ȭ"
         Height          =   375
         Left            =   -68400
         TabIndex        =   186
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "�����"
         Height          =   375
         Left            =   -70080
         TabIndex        =   185
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdCalcOperator 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   -67680
         TabIndex        =   184
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   15
         Left            =   -74880
         TabIndex        =   183
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   14
         Left            =   -74880
         TabIndex        =   182
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   13
         Left            =   -74880
         TabIndex        =   181
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   12
         Left            =   -72000
         TabIndex        =   180
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   11
         Left            =   -72000
         TabIndex        =   179
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   9
         Left            =   -72000
         TabIndex        =   178
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   8
         Left            =   -72960
         TabIndex        =   177
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   7
         Left            =   -72960
         TabIndex        =   176
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   6
         Left            =   -72960
         TabIndex        =   175
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   5
         Left            =   -72960
         TabIndex        =   174
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   4
         Left            =   -73920
         TabIndex        =   173
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   3
         Left            =   -73920
         TabIndex        =   172
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   2
         Left            =   -73920
         TabIndex        =   171
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   1
         Left            =   -73920
         TabIndex        =   170
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   0
         Left            =   -74880
         TabIndex        =   169
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcAdvancedOperator 
         Enabled         =   0   'False
         Height          =   615
         Index           =   10
         Left            =   -72000
         TabIndex        =   168
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcCalculate 
         Caption         =   "="
         Height          =   615
         Left            =   -68880
         TabIndex        =   167
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcPoint 
         Caption         =   "."
         Height          =   615
         Left            =   -69840
         TabIndex        =   166
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcOperator 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   -67680
         TabIndex        =   165
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcOperator 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   -67680
         TabIndex        =   164
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcOperator 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -67680
         TabIndex        =   163
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "9"
         Height          =   615
         Index           =   9
         Left            =   -68880
         TabIndex        =   162
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "8"
         Height          =   615
         Index           =   8
         Left            =   -69840
         TabIndex        =   161
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "7"
         Height          =   615
         Index           =   7
         Left            =   -70800
         TabIndex        =   160
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "6"
         Height          =   615
         Index           =   6
         Left            =   -68880
         TabIndex        =   159
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "5"
         Height          =   615
         Index           =   5
         Left            =   -69840
         TabIndex        =   158
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "4"
         Height          =   615
         Index           =   4
         Left            =   -70800
         TabIndex        =   157
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "3"
         Height          =   615
         Index           =   3
         Left            =   -68880
         TabIndex        =   156
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "1"
         Height          =   615
         Index           =   1
         Left            =   -70800
         TabIndex        =   155
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "2"
         Height          =   615
         Index           =   2
         Left            =   -69840
         TabIndex        =   154
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCalcNumbers 
         Caption         =   "0"
         Height          =   615
         Index           =   0
         Left            =   -70800
         TabIndex        =   153
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  '������ ����
         Height          =   270
         Left            =   -74760
         TabIndex        =   152
         Top             =   240
         Width           =   7935
      End
      Begin VB.ListBox lvContacts 
         Height          =   4020
         ItemData        =   "frmMain2.frx":60B30
         Left            =   -74880
         List            =   "frmMain2.frx":60B37
         TabIndex        =   139
         Top             =   76
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   -73080
         TabIndex        =   126
         Top             =   98
         Width           =   4935
         Begin VB.TextBox txtEmail 
            Height          =   270
            Left            =   1080
            TabIndex        =   132
            Top             =   550
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   600
            TabIndex        =   131
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtCellPhone 
            Height          =   270
            Left            =   3360
            TabIndex        =   130
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPostalCode 
            Height          =   270
            Left            =   1080
            TabIndex        =   129
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtAddress 
            Height          =   270
            Left            =   2520
            TabIndex        =   128
            Top             =   900
            Width           =   2295
         End
         Begin VB.ComboBox cmbGroup 
            Height          =   300
            Left            =   3360
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   127
            Top             =   560
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "�̸�:"
            Height          =   255
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "�޴���ȭ:"
            Height          =   255
            Left            =   2520
            TabIndex        =   137
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "���ڿ���:"
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "�����ȣ:"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   950
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "�ּ�:"
            Height          =   255
            Left            =   2040
            TabIndex        =   134
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "�׷�:"
            Height          =   255
            Left            =   2880
            TabIndex        =   133
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   975
         Left            =   -73080
         TabIndex        =   117
         Top             =   1516
         Width           =   4935
         Begin VB.TextBox txtCompany 
            Height          =   270
            Left            =   3000
            TabIndex        =   121
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtHome 
            Height          =   270
            Left            =   720
            TabIndex        =   120
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtFax 
            Height          =   270
            Left            =   600
            TabIndex        =   119
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtOtherNumber 
            Height          =   270
            Left            =   2880
            TabIndex        =   118
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "��:"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  '������ ����
            Caption         =   "ȸ��(�б�):"
            Height          =   255
            Left            =   1800
            TabIndex        =   124
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label6 
            Caption         =   "�ѽ�:"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "��Ÿ:"
            Height          =   255
            Left            =   2400
            TabIndex        =   122
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdSaveContact 
         Caption         =   "����(&S)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   116
         Top             =   196
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1575
         Left            =   -73080
         TabIndex        =   109
         Top             =   2596
         Width           =   4935
         Begin VB.TextBox txtContent 
            Height          =   975
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   113
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtBYear 
            Height          =   270
            Left            =   120
            TabIndex        =   112
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtBMonth 
            Height          =   270
            Left            =   840
            TabIndex        =   111
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtBDay 
            Height          =   270
            Left            =   1320
            TabIndex        =   110
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Label22"
            Height          =   255
            Left            =   1920
            TabIndex        =   114
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.FileListBox lvContactFiles 
         Height          =   270
         Left            =   -69240
         TabIndex        =   108
         Top             =   75
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDelContact 
         Caption         =   "delete"
         Height          =   495
         Left            =   -68040
         TabIndex        =   107
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ListBox lvTasks 
         Height          =   4050
         ItemData        =   "frmMain2.frx":60B4E
         Left            =   -74880
         List            =   "frmMain2.frx":60B55
         Style           =   1  'Ȯ�ζ�
         TabIndex        =   106
         Top             =   76
         Width           =   2295
      End
      Begin VB.CommandButton cmdSaveTask 
         Caption         =   "cmdSaveTask"
         Height          =   495
         Left            =   -67920
         TabIndex        =   105
         Top             =   76
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelTask 
         Caption         =   "cmdDelTask"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -67920
         TabIndex        =   104
         Top             =   3076
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4095
         Left            =   -72480
         TabIndex        =   89
         Top             =   98
         Width           =   4455
         Begin VB.TextBox txtPercentage 
            Alignment       =   1  '������ ����
            Height          =   270
            Left            =   3450
            TabIndex        =   96
            Text            =   "0"
            Top             =   1200
            Width           =   420
         End
         Begin VB.TextBox txtTaskTitle 
            Height          =   270
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtMemo 
            Height          =   1335
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   94
            Top             =   2640
            Width           =   4215
         End
         Begin VB.TextBox txtImpt 
            Height          =   270
            Left            =   120
            MaxLength       =   2
            TabIndex        =   92
            Text            =   "1"
            Top             =   1920
            Width           =   480
         End
         Begin VB.TextBox txtPart 
            Height          =   270
            Left            =   1080
            TabIndex        =   90
            Top             =   1920
            Width           =   3255
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   270
            Left            =   600
            TabIndex        =   91
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            OrigLeft        =   600
            OrigTop         =   1920
            OrigRight       =   855
            OrigBottom      =   2175
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   3840
            TabIndex        =   93
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   327681
            BuddyControl    =   "cmdDeleteAllContacts"
            BuddyDispid     =   196728
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
         Begin ComctlLib.ProgressBar TaskProgress 
            Height          =   300
            Left            =   120
            TabIndex        =   97
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label9 
            Alignment       =   1  '������ ����
            Caption         =   "%"
            Height          =   255
            Left            =   4155
            TabIndex        =   102
            Top             =   1245
            Width           =   135
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Label16"
            Height          =   255
            Left            =   1080
            TabIndex        =   98
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.FileListBox lvTaskFiles 
         Height          =   270
         Left            =   -67920
         TabIndex        =   88
         Top             =   675
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.DirListBox Dir1 
         Height          =   300
         Left            =   -66960
         TabIndex        =   87
         Top             =   -44
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdDeleteAllContacts 
         Caption         =   "clear"
         Height          =   495
         Left            =   -68040
         TabIndex        =   86
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmdResetFields 
         Caption         =   "���� �ʱ�ȭ(&R)"
         Height          =   495
         Left            =   -68040
         TabIndex        =   85
         Top             =   3676
         Width           =   1350
      End
      Begin VB.CommandButton cmdDeleteAllTasks 
         Caption         =   "cmdDeleteAllTasks"
         Height          =   495
         Left            =   -67920
         TabIndex        =   84
         Top             =   3676
         Width           =   1215
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   0
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   1
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   2
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   3
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   4
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   79
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   5
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   458
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   6
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   458
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   7
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   8
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   9
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   10
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   11
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   12
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   938
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   13
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   938
         Width           =   2055
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
         Index           =   15
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   16
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   17
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   18
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   19
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   1418
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   20
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   1418
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   21
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   22
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   23
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   60
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   24
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   25
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   26
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   2018
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   27
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   2018
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   28
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   29
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   30
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   31
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   52
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   32
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   33
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   2618
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   34
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   2618
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   35
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   36
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   37
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   38
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   39
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   40
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   3098
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   41
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   3098
         Width           =   2055
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   42
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   43
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   44
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   45
         Left            =   -71760
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   46
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   47
         Left            =   -69840
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   3578
         Width           =   975
      End
      Begin VB.TextBox txtPlannerTF 
         Height          =   495
         Index           =   48
         Left            =   -68880
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3578
         Width           =   2055
      End
      Begin VB.CommandButton sdcmdSavePlanner 
         Caption         =   "sdcmdSavePlanner"
         Height          =   375
         Left            =   -68040
         TabIndex        =   34
         Top             =   98
         Width           =   1215
      End
      Begin VB.TextBox txtAlarmTitle 
         Height          =   270
         Left            =   3240
         TabIndex        =   32
         Top             =   218
         Width           =   4935
      End
      Begin VB.TextBox txtTimeHrs 
         Height          =   270
         Left            =   3240
         TabIndex        =   31
         Top             =   578
         Width           =   375
      End
      Begin VB.TextBox txtTimeMin 
         Height          =   270
         Left            =   3600
         TabIndex        =   30
         Top             =   578
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1335
         Left            =   2640
         TabIndex        =   29
         Top             =   858
         Width           =   5535
      End
      Begin VB.CommandButton cmdResetAF 
         Caption         =   "cmdResetAF"
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   3818
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveAlarm 
         Caption         =   "cmdSaveAlarm"
         Height          =   375
         Left            =   6840
         TabIndex        =   27
         Top             =   3818
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeleteAlarm 
         Caption         =   "cmdDeleteAlarm"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   3818
         Width           =   1335
      End
      Begin VB.TextBox txtAlarmMemo 
         Height          =   1215
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2498
         Width           =   5535
      End
      Begin VB.FileListBox lvAlarmFiles 
         Height          =   270
         Left            =   7440
         TabIndex        =   24
         Top             =   -15
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4095
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   120
         Width           =   8175
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4170
         Left            =   -74880
         TabIndex        =   23
         Top             =   105
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   7355
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483636
         Appearance      =   0
         MonthColumns    =   3
         MonthRows       =   2
         StartOfWeek     =   70909953
         CurrentDate     =   44330
      End
      Begin ComctlLib.ListView lvAlarms 
         Height          =   4095
         Left            =   120
         TabIndex        =   33
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
      Begin VB.Label lbltsw 
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   216
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label29 
         BackStyle       =   0  '����
         Caption         =   "."
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -70080
         TabIndex        =   197
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label28 
         BackStyle       =   0  '����
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -71760
         TabIndex        =   196
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  '����
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -73320
         TabIndex        =   195
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblTimerSD 
         BackStyle       =   0  '����
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -69720
         TabIndex        =   194
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblTimerS 
         BackStyle       =   0  '����
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -71280
         TabIndex        =   193
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTimerM 
         BackStyle       =   0  '����
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -72960
         TabIndex        =   192
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTimerH 
         BackStyle       =   0  '����
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -74520
         TabIndex        =   191
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000003&
         BackStyle       =   1  '�������� ����
         BorderStyle     =   0  '����
         Height          =   975
         Left            =   -74760
         Shape           =   4  '�ձ� �簢��
         Top             =   720
         Width           =   5895
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  '�������� ����
         BorderStyle     =   2  '���
         Height          =   1215
         Left            =   -74880
         Shape           =   4  '�ձ� �簢��
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label lblDOW 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   144
         Top             =   578
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "7             9             12              15             18            21               22-"
         Height          =   225
         Left            =   -74280
         TabIndex        =   143
         Top             =   218
         Width           =   6135
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   2640
         TabIndex        =   142
         Top             =   218
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   2640
         TabIndex        =   141
         Top             =   578
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   2640
         TabIndex        =   140
         Top             =   2258
         Width           =   975
      End
   End
   Begin VB.Image cmdHelp 
      Height          =   240
      Left            =   1680
      Picture         =   "frmMain2.frx":60B6A
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image cmdPinMenu 
      Height          =   240
      Left            =   840
      Picture         =   "frmMain2.frx":612AC
      Stretch         =   -1  'True
      ToolTipText     =   "���� �޴� ����/�����"
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image cmdMnuOptions 
      Height          =   240
      Left            =   1250
      Picture         =   "frmMain2.frx":616EE
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  '�������� ����
      Height          =   375
      Left            =   720
      Shape           =   4  '�ձ� �簢��
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgPressed 
      Height          =   930
      Left            =   10185
      Picture         =   "frmMain2.frx":61B30
      Top             =   6585
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblDatetime 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   151
      Top             =   7365
      Width           =   1455
   End
   Begin VB.Label lblStatustext 
      BackStyle       =   0  '����
      Caption         =   "�޷¿��� ��¥�� �����ʽÿ�."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   150
      Top             =   7365
      Width           =   7935
   End
   Begin VB.Image imgHover 
      Height          =   930
      Left            =   10200
      ToolTipText     =   "���� ��ħ"
      Top             =   6600
      Width           =   930
   End
   Begin VB.Image imgHover2 
      Height          =   930
      Left            =   10185
      Picture         =   "frmMain2.frx":648FA
      Top             =   6585
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image imgDehover 
      Height          =   2655
      Left            =   8880
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6735
      Index           =   0
      Left            =   0
      Picture         =   "frmMain2.frx":676C4
      Top             =   1080
      Width           =   11385
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   10920
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image menuhover 
      Height          =   1455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   9840
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   600
      TabIndex        =   145
      Top             =   8640
      Width           =   9255
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   0
      Picture         =   "frmMain2.frx":1615EE
      Top             =   1080
      Visible         =   0   'False
      Width           =   11385
   End
   Begin VB.Image Image1 
      Height          =   6720
      Index           =   1
      Left            =   0
      Picture         =   "frmMain2.frx":25B51A
      Top             =   1080
      Width           =   11385
   End
   Begin VB.Image Image3 
      Height          =   6720
      Index           =   1
      Left            =   0
      Picture         =   "frmMain2.frx":354B5C
      Top             =   1080
      Visible         =   0   'False
      Width           =   11400
   End
   Begin VB.Image Image1 
      Height          =   6735
      Index           =   2
      Left            =   0
      Picture         =   "frmMain2.frx":44E19E
      Top             =   1080
      Width           =   11385
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   2
      Left            =   0
      Picture         =   "frmMain2.frx":5480C8
      Top             =   1080
      Visible         =   0   'False
      Width           =   11385
   End
End
Attribute VB_Name = "frmMain"
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
 
'https://www.vbforums.com/showthread.php?548604-RESOLVED-how-to-make-form-moveable&p=3389095#post3389095
Dim theme As Integer
Dim ContactChanged As Boolean

Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

'https://www.vbforums.com/showthread.php?462116-Marquee-Progressbar
Private Const WM_USER = &H400
Private Const PBM_SETMARQUEE = WM_USER + 10
'Private Const GWL_STYLE = (-16)
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Dim PrevStyle As Long
Option Explicit
'Private Const FLAG_ICC_FORCE_CONNECTION As Long = &H1
'Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Dim th As Integer
Dim tm As Integer
Dim ts As Integer

Dim pos As POINTAPI

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim contact As Integer
Dim iFileNo As Integer
Dim Task As Integer
Dim ctimert As Integer

'�ۿ°�: http://www.vbforums.com/showthread.php?595990-VB6-System-tray-icon-systray
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' The following code is required:

Private Function SetUpMarquee(ObjProgressBar, Optional TurnOn As Boolean, Optional Speed As Integer = 50)
    Select Case TurnOn
        Case True
            'Set the style to Marquee
            SetWindowLong ObjProgressBar.hwnd, GWL_STYLE, PrevStyle Or PBM_SETMARQUEE
            'Make it automatically scroll
            SetUpMarquee = SendMessage(ObjProgressBar.hwnd, PBM_SETMARQUEE, CLng(TurnOn), CLng(Speed))
        Case False
            'Set the Style Back
            SetWindowLong ObjProgressBar.hwnd, GWL_STYLE, PrevStyle
    End Select
End Function

Sub ContactDataChanged()
    ContactChanged = -1
    cmdSaveContact.FontBold = -1
End Sub

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



Private Sub chkFavorite_Click()
    ContactDataChanged
End Sub

Private Sub cmbGroup_Change()
    ContactDataChanged
End Sub

Private Sub cmbIndexContactGroups_Change()
    GetGroups
End Sub

Private Sub cmbIndexContactGroups_Click()
    GetGroups
End Sub

Private Sub cmbIndexContactGroups_KeyPress(KeyAscii As Integer)
    GetGroups
End Sub

Private Sub cmbIndexContactGroups_Scroll()
    GetGroups
End Sub

Private Sub cmdDeleteAlarm_Click()
    On Error Resume Next
    If Confirm("�� ���� ����ϴµ� ������ �˶��� �����ұ��?", "���", Me) Then
        Kill "C:\CALPLANS\ALARMS\" & lvAlarms.SelectedItem.SubItems(1)
        
        ClearAlarmFields
        
        LoadAlarms
    End If
End Sub

Private Sub cmdDuplicateContact_Click()
    Dim cn As String
    cn = InputBox("����ó �̸�: ", "����ó ����", "")
    If Len(cn) Then
        txtName.Text = cn
        chkDuplicateContact.Value = 1
        cmdSaveContact_Click
        chkDuplicateContact.Value = 0
    End If
End Sub

Private Sub cmdMnuAbout_Click()
    mnuHelpAbout_Click
End Sub

Private Sub cmdMnuOptions_Click()
    cmdOptions_Click
End Sub

Private Sub cmdPinMenu_Click()
    If cmdPinMenu.Tag = "+" Then
        cmdPinMenu.Tag = "-"
        SaveSetting "Calendar", "Options", "AlwaysRibbon", "1"
        hideshow.Enabled = 0
        showrm
    Else
        cmdPinMenu.Tag = "+"
        SaveSetting "Calendar", "Options", "AlwaysRibbon", "0"
        hideshow.Enabled = -1
    End If
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
    '�Է°��� �˻��Ѵ�.
    If Mid$(txtTimeMin.Text, 1, 1) = "0" Then
        txtTimeMin.Text = Mid$(txtTimeMin.Text, 2, 1)
    End If
    If InStr(1, txtAlarmTitle.Text, "?") > 0 Or InStr(1, txtAlarmTitle.Text, "\") > 0 Or InStr(1, txtAlarmTitle.Text, "|") > 0 Or InStr(1, txtAlarmTitle.Text, ".") > 0 Or InStr(1, txtAlarmTitle.Text, "/") > 0 Or InStr(1, txtAlarmTitle.Text, "*") > 0 Or InStr(1, txtAlarmTitle.Text, ":") > 0 Or InStr(1, txtAlarmTitle.Text, ChrW$(34)) > 0 Or txtAlarmTitle.Text = LoadLang("�� �˶� �߰�...", "New...") Then
        MessageBox "������ ���� �ùٸ��� �ʽ��ϴ�.", "�Է� �� ����", Me, 16
    End If
    If IsNumeric(txtTimeHrs.Text) = False Or IsNumeric(txtTimeMin.Text) = False Then
        MessageBox "�ð��� ���� �ùٸ��� �ʽ��ϴ�.", "�Է� �� ����", Me, 16
        Exit Sub
    End If
    If GetSetting("Calendar", "Options", "NoTimeCheck", 0) = 0 Then
        If txtTimeHrs.Text > 23 Or txtTimeMin.Text > 59 Or txtTimeHrs.Text < 0 Or txtTimeMin.Text < 0 Then
            MessageBox "�ð����� �ô� 0���� 23, ���� 0���� 59������ �����̿��� �մϴ�.", "�Է� �� ����", Me, 16
            Exit Sub
    End If
    End If
    If txtAlarmTitle.Text = "" Then
        MessageBox "������ ���� �ʼ��Դϴ�.", "�Է� �� ����", Me, 16
        Exit Sub
    End If
    
    '������ �߰��ϱ� ���� �ش� ������ ������ �����ϴ��� Ȯ���Ѵ�.
    If FileExists("C:\CALPLANS\ALARMS\" & txtAlarmTitle.Text) = True And lvAlarms.SelectedItem.SubItems(1) = LoadLang("�� �˶� �߰�...", "New...") Then
        MessageBox "�ش� �̸��� �˶��� �̹� �����մϴ�.", "ó�� �� ����", Me, 16
    End If
    
    '�ش� �˶��� �������� �˸��� ������ �����.
    'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
    Dim iFileNo As Integer
    iFileNo = FreeFile
    '������ ����.
    Open "C:\CALPLANS\ALARMS\" & txtAlarmTitle.Text For Output As #iFileNo
    
    '������ ������ ���� �����Ƿ� �� ĭ����...
    Print #iFileNo, ""
    
    '������ �ݴ´�.
    Close #iFileNo
    
    Dim txtTime As String
    
    '������Ʈ���� ������ ��Ÿ ������ �����Ѵ�.
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

Sub SetTimerLabel()
    If Len(lblTimerH.Caption) = 1 Then lblTimerH.Caption = "0" & lblTimerH.Caption
    If Len(lblTimerM.Caption) = 1 Then lblTimerM.Caption = "0" & lblTimerM.Caption
    If Len(lblTimerS.Caption) = 1 Then lblTimerS.Caption = "0" & lblTimerS.Caption
End Sub

Sub SetTimerTime()
    On Error GoTo errln
    If _
        (IsNumeric(txtSetH.Text)) And _
        (IsNumeric(txtSetM.Text) And txtSetM.Text <= 59 And txtSetM.Text >= 0) And _
        (IsNumeric(txtSetS.Text) And txtSetS.Text <= 59 And txtSetS.Text >= 0) And _
        (IsNumeric(txtSetSD.Text) And txtSetSD.Text <= 9 And txtSetSD.Text >= 0) And _
        (1) _
    Then
        lblTimerH.Caption = txtSetH.Text
        lblTimerM.Caption = txtSetM.Text
        lblTimerS.Caption = txtSetS.Text
        lblTimerSD.Caption = txtSetSD.Text
        
        th = txtSetH.Text
        tm = txtSetM.Text
        ts = txtSetS.Text
        
        SetTimerLabel
        
        pbTimer.Max = 10 * (ts + tm * 60 + th * 3600) + txtSetSD.Text
    Else
        MsgBox "�ð��� �ùٸ��� �ʽ��ϴ�.", 16, "����"
        Exit Sub
    End If
    
    th = txtSetH.Text
    tm = txtSetM.Text
    ts = txtSetS.Text
    pbTimer.Max = 10 * (ts + tm * 60 + th * 3600) + txtSetSD.Text
    
    Exit Sub
    
errln:
    'MsgBox "�ð��� �ùٸ��� �ʽ��ϴ�.", 16, "����"
End Sub

Private Sub cmdSetTimerTime_Click()
    SetTimerTime
End Sub

Private Sub cmdTimerPause_Click()
    timTImer.Enabled = 0
    If ssTimerType.Tab Then SetUpMarquee pbTimer, False, 50
    cmdTimerStart.Caption = "���"
    cmdTimerStart.Enabled = -1
    cmdTimerPause.Enabled = 0
End Sub

Private Sub cmdTimerStart_Click()
    'SetTimerTime
    If ssTimerType.Tab Then
        SetUpMarquee pbTimer, True, 50
    Else
        SetUpMarquee pbTimer, False, 50
    End If
    timTImer.Enabled = -1
    cmdSetTimerTime.Enabled = 0
    cmdTimerStart.Enabled = 0
    cmdTimerPause.Enabled = -1
End Sub

Private Sub cmdTImerStop_Click()
    lblTimerH.Caption = "00"
    lblTimerM.Caption = "00"
    lblTimerS.Caption = "00"
    lblTimerSD.Caption = "00"
    timTImer.Enabled = 0
    
    SetUpMarquee pbTimer, False, 50
    pbTimer.Value = 0
    
    cmdTimerStart.Caption = "����"
    
    cmdSetTimerTime.Enabled = -1
    cmdTimerStart.Enabled = -1
    cmdTimerPause.Enabled = 0
    
    lbltsw.Visible = 0
End Sub

Sub GetGroups()
    lvContactFiles3.Refresh
    lvContactFiles3.Path = "C:\CALPLANS\CONTACTS"
    lvGroupContacts.Clear
    SaveSetting "Calendar", "Contacts", txtName.Text & "Group", cmbGroup.Text
    Dim contact As Integer
    For contact = 0 To lvContactFiles3.ListCount - 1
        If GetSetting("Calendar", "Contacts", lvContactFiles3.List(contact) & "Group") = cmbIndexContactGroups.Text Then
            lvGroupContacts.AddItem lvContactFiles3.List(contact)
        End If
    Next contact
End Sub

Sub GetFavorites()
    lvContactFiles2.Refresh
    lvFavoriteContacts.Clear
    lvContactFiles2.Path = "C:\CALPLANS\CONTACTS"
    Dim contact As Integer
    For contact = 0 To lvContactFiles2.ListCount - 1
        If GetSetting("Calendar", "Contacts", lvContactFiles2.List(contact) & "Fav", 0) = 1 Then
            lvFavoriteContacts.AddItem lvContactFiles2.List(contact)
        End If
    Next contact
End Sub

Private Sub cmdTltRef_Click()
    lvTodaysPlan.Refresh
    lvTodaysPlans.Refresh
    lvTmrPlans.Refresh
    
    GetFavorites
    GetGroups
End Sub

' End required code
' /////////////////////////////////////////////

Sub LoadContacts()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CONTACTS"
    
    lvContacts.Clear
    lvContacts.AddItem LoadLang("�� ����ó �߰�...", "New...")
    
    lvContactFiles.Refresh
    
    lvContacts.ListIndex = 0
    lvContactFiles.Path = "C:\CALPLANS\CONTACTS"
    
    For contact = 0 To lvContactFiles.ListCount - 1
        lvContacts.AddItem lvContactFiles.List(contact)
    Next contact
End Sub

Private Sub cmdDelContact_Click()
    If Confirm(lvContacts.List(lvContacts.ListIndex) & " ����ó�� �����Ͻðڽ��ϱ�?", "�ּҷ� ����", Me) Then
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
        
        If lvContacts.List(lvContacts.ListIndex) = "�Ѹ�" Then
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
    If Confirm(txtTaskTitle.Text & " �۾��� �����Ͻðڽ��ϱ�?", "�۾� ����", Me) Then
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
        PopupMenu frmMainOld.mnuHelp, , cmdHelp.Left, cmdHelp.Top + cmdHelp.height ' - 2350 - ssTodaysPlan.width + 100, 400
    Else
        PopupMenu frmMainOld.mnuHelp, , Me.width - 2350, 400
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
    If Confirm("�ѹ��� ����մϴ�. ��� �Է»����� ���� �ʱ�ȭ�Ͻðڽ��ϱ�?", "�ʱ�ȭ", Me) Then
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
        MessageBox "�̸��� ���� �ùٸ��� �ʽ��ϴ�.", "�Է� �� ����", Me, 16
        Exit Sub
    End If
    
    cmdSaveContact.FontBold = 0
    
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
    
    SaveSetting "Calendar", "Contacts", txtName.Text & "Fav", chkFavorite.Value
    
    If txtName.Text = "�Ѹ�" And txtBYear.Text = "1983" And (txtBMonth.Text = "4" Or txtBMonth.Text = "04") And txtBDay.Text = "22" Then
        SaveSetting "Calendar", "Config", "EggEnabled", "1"
    End If
    
    If lvContacts.List(lvContacts.ListIndex) = LoadLang("�� ����ó �߰�...", "New...") Or chkDuplicateContact.Value Then
        '�ش� ����ó�� �������� �˸��� ������ �����.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFileNo = FreeFile
        '������ ����.
        Open "C:\CALPLANS\CONTACTS\" & txtName.Text For Output As #iFileNo
        
        '������ ������ ���� �����Ƿ� �� ĭ����...
        Print #iFileNo, ""
        
        '������ �ݴ´�.
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
    
    lvTasks.AddItem LoadLang("�� �۾� �߰�...", "New...")
    
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
        MessageBox "������ ���� �ùٸ��� �ʽ��ϴ�.", "�Է� �� ����", Me, 16
        Exit Sub
    End If
    
    
    If IsNumeric(txtImpt.Text) = False Or txtImpt.Text < 1 Or txtImpt.Text > 10 Then
        MessageBox "�߿䵵�� 1(����)���� 10(����)�������� �մϴ�.", "�Է� �� ����", Me, 16
        Exit Sub
    End If
    
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Perc", txtPercentage.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Memo", txtMemo.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Impt", txtImpt.Text
    SaveSetting "Calendar", "Tasks", txtTaskTitle.Text & "Part", txtPart.Text
    
    If lvTasks.List(lvTasks.ListIndex) = LoadLang("�� �۾� �߰�...", "New...") Then
        '�ش� �۾��� �������� �˸��� ������ �����.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        iFileNo = FreeFile
        '������ ����.
        Open "C:\CALPLANS\TASKS\" & txtTaskTitle.Text For Output As #iFileNo
        
        '������ ������ ���� �����Ƿ� �� ĭ����...
        Print #iFileNo, ""
        
        '������ �ݴ´�.
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
    If Confirm("�����Ͻðڽ��ϱ�?", "����", Me) Then
        If Confirm("���� *�Ұ���*�մϴ�. ������ " & MonthView1.SelStart & "�� ��� ������ �����Ͻðڽ��ϱ�?", "����", Me, , True) Then
            On Error Resume Next
            Shell "CMD /C RD /S /Q " & ChrW$(34) & "C:\CALPLANS\" & Split(MonthView1.SelStart, "-")(0) & "\" & Split(MonthView1.SelStart, "-")(1) & "\" & Split(MonthView1.SelStart, "-")(2) & ChrW$(34)
            Shell "COMMAND /C DELTREE /Y " & ChrW$(34) & "C:\CALPLANS\" & Split(MonthView1.SelStart, "-")(0) & "\" & Split(MonthView1.SelStart, "-")(1) & "\" & Split(MonthView1.SelStart, "-")(2) & ChrW$(34)
            
            MessageBox "�����Ǿ����ϴ�.", "����", Me
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
    
    Select Case theme
        Case 0
            SSTab1.BackColor = RGB(66, 156, 255)
            ssTodaysPlan.BackColor = RGB(66, 156, 255)
            'ssRibbonMenu.BackColor = RGB(66, 156, 255)
            'cmdRibbonFile.BackColor = RGB(66, 156, 255)
            'cmdHelp.BackColor = RGB(66, 156, 255)
            'cmdMnuOptions.BackColor = RGB(66, 156, 255)
            'cmdPinMenu.BackColor = RGB(66, 156, 255)
            Command1.BackColor = RGB(66, 156, 255)
            Command2.BackColor = RGB(66, 156, 255)
            cmdMnuAbout.BackColor = RGB(66, 156, 255)
            Me.BackColor = RGB(255, 0, 255)
            
            'ssTodaysPlan.BackColor = Me.BackColor
            'cmdHelp.BackColor = RGB(66, 156, 255)
            
            hider1.BackColor = RGB(66, 156, 255)
            hider2.BackColor = RGB(66, 156, 255)
        Case 1
            SSTab1.BackColor = RGB(248, 164, 24)
            ssTodaysPlan.BackColor = RGB(248, 164, 24)
            'ssRibbonMenu.BackColor = RGB(248, 164, 24)
            'cmdRibbonFile.BackColor = RGB(248, 164, 24)
            'cmdHelp.BackColor = RGB(248, 164, 24)
            'cmdMnuOptions.BackColor = RGB(248, 164, 24)
            'cmdPinMenu.BackColor = RGB(248, 164, 24)
            Command1.BackColor = RGB(248, 164, 24)
            Command2.BackColor = RGB(248, 164, 24)
            cmdMnuAbout.BackColor = RGB(248, 164, 24)
            Me.BackColor = RGB(255, 0, 255)
            
            'ssTodaysPlan.BackColor = Me.BackColor
            'cmdHelp.BackColor = RGB(248, 164, 24)
            
            hider1.BackColor = RGB(248, 164, 24)
            hider2.BackColor = RGB(248, 164, 24)
        Case 2
            SSTab1.BackColor = RGB(236, 220, 46)
            ssTodaysPlan.BackColor = RGB(236, 220, 46)
            'ssRibbonMenu.BackColor = RGB(236, 220, 46)
            'cmdRibbonFile.BackColor = RGB(236, 220, 46)
            'cmdHelp.BackColor = RGB(236, 220, 46)
            'cmdMnuOptions.BackColor = RGB(236, 220, 46)
            'cmdPinMenu.BackColor = RGB(236, 220, 46)
            Command1.BackColor = RGB(236, 220, 46)
            Command2.BackColor = RGB(236, 220, 46)
            cmdMnuAbout.BackColor = RGB(236, 220, 46)
            Me.BackColor = RGB(255, 0, 255)
            
            'ssTodaysPlan.BackColor = Me.BackColor
            'cmdHelp.BackColor = RGB(236, 220, 46)
            
            hider1.BackColor = RGB(236, 220, 46)
            hider2.BackColor = RGB(236, 220, 46)
    End Select
    
End Sub

Private Sub cmdUnselectAllDW_Click()
    Dim i As Integer
    For i = 0 To 6
        chkDayOfWeeks(i).Value = 0
    Next i
End Sub

Private Sub cmdRibbonFile_Click()
    PopupMenu frmMainOld.mnuFile, , cmdRibbonFile.Left, cmdRibbonFile.Top + cmdRibbonFile.height
End Sub

Private Sub Command1_Click()
    Form_Unload 0
End Sub

Private Sub Command11_Click()
    SetTimerTime
End Sub

Private Sub Command14_Click()
    txtTimerLog.Text = txtTimerLog.Text & lblTimerH.Caption & ":" & lblTimerM.Caption & ":" & lblTimerS.Caption & "." & lblTimerSD.Caption & vbNewLine
End Sub

Private Sub Command15_Click()

End Sub

Private Sub Command2_Click()
    Me.WindowState = 1
End Sub

Private Sub Command4_Click()
    If SSTab1.Tab > 0 Then
        SSTab1.Tab = SSTab1.Tab - 1
    End If
End Sub

Private Sub Command5_Click()
    If SSTab1.Tab < SSTab1.Tabs - 1 Then
        SSTab1.Tab = SSTab1.Tab + 1
    End If
End Sub

Private Sub datetime_Timer()
    lblDatetime.Caption = Format(Now, "H�� M��")
End Sub

Private Sub Form_Load()
    'MsgBox DayOfWeek()
    'MessageBox PlayFair("dlfjs qkqhrkxdms sdfhuj", "ultra"), "3", Me
    theme = GetSetting("Calendar", "Options", "Theme2", 0)
    Image1(0).Visible = 0
    Image3(0).Visible = 0
    Image1(1).Visible = 0
    Image3(1).Visible = 0
    Image1(2).Visible = 0
    Image3(2).Visible = 0
    'MsgBox theme
    
    If theme = 1 Then Shape3.Visible = -1
    
    ContactChanged = 0
    
    ctimert = 0
    
    PrevStyle = GetWindowLong(pbTimer.hwnd, GWL_STYLE)
    
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbMagenta, 0&, LWA_COLORKEY
    
    

    If GetSetting("Calendar", "Options", "TP", 0) = 1 Then
        'Me.Width = 8715
    End If
    
    If GetSetting("Calendar", "Options", "NoRibbon", 0) = 1 Then
        SSTab1.Top = 120
        ssTodaysPlan.height = 4695
        lvTodaysPlan.height = 3870
        cmdTltRef.Top = 4440
        'Me.Height = 5900
        
        ssRibbonMenu.Visible = False
        cmdHelp.Visible = False
        cmdRibbonFile.Visible = False
        cmdMnuAbout.Visible = False
        cmdMnuOptions.Visible = False
        
        'mnuDateMenu.Visible = True
        'mnuFile.Visible = True
        'mnuView.Visible = True
        'mnuHelp.Visible = True
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
    
    cmbGroup.AddItem LoadLang("���� �� ��", "None")
    cmbIndexContactGroups.AddItem LoadLang("���� �� ��", "None")
    
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
                    MessageBox "���������� Ǯ�׸��� �����մϴ�." & vbCrLf & vbCrLf & _
                           "    PLNMGR32.EXE [/R]" & vbCrLf & vbCrLf & _
                           "    /R  �ּ�ȭ�� ���·� �����մϴ�.", _
                           "����ġ ����", Me

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
            MessageBox LoadLang("����ġ�� Ʋ���ϴ�", "Switch is wrong", "El comando no es valido.") & " - " & Command, LoadLang("����", "Error", "Error"), Me, 16
            End
    End Select
    
    'mnuHelpAbout.Caption = App.Title & " ����(&A)"
    
    'frmNotifyMgr.Show

    Me.Left = GetSetting("Calendar", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("Calendar", "Settings", "MainTop", 1000)
    
    Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab)
    Me.Caption = Me.Caption & " (" & MonthView1.Year & "�� " & MonthView1.Month & "��)"
    
    If GetSetting("Calendar", "Config", "FirstRun", "0") = "0" Then
        'SaveSetting "Calendar", "Config", "FirstRun", "1"
        
        frmWizard.Show vbModal, Me
        
        If GetWinver(1) >= 6 And GetWinver(2) >= 1 Then
        Else
            'MessageBox LoadLang("��ǻ�Ͱ� Windows Vista Ȥ�� Windows XP ������ � ü���� �����ϰ� �ֽ��ϴ�. �޷��� �ùٷ� ǥ�õ��� ���� �� �ֽ��ϴ�.", "Your PC is running Windows VIsta or earlier. The calendar may display incorrectly.", "La computadora esta ejecutando un sistema operativo de Windows Vista o Windows XP o inferior. Es posible que el calendario no se muestre correctamente."), LoadLang("���", "Warning", "Advertencia"), Me, 48
        End If
        
        MessageBox LoadLang("��ǻ�͸� ������ ������ �˸��� �������� ", "Add ", "Agregue ") & "[" & ChrW$(34) & Dir1.Path & "\PLNMGR32.EXE" & ChrW$(34) & " /R]" & _
               LoadLang("(��� �����) �ٷΰ��⸦ �������α׷��� �߰��Ͻʽÿ�.", "(Path Copied) to your startup program to be notified when you start your computer.", "(Ruta copiada) a su programa de inicio para recibir una notificacion cuando inicie su computadora."), LoadLang("�˸��� Ȱ��ȭ", "Tip", "Propina"), Me
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
    DOWLS(0) = LoadLang("��", "S")
    DOWLS(1) = LoadLang("��", "M")
    DOWLS(2) = LoadLang("ȭ", "T")
    DOWLS(3) = LoadLang("��", "W")
    DOWLS(4) = LoadLang("��", "T")
    DOWLS(5) = LoadLang("��", "F")
    DOWLS(6) = LoadLang("��", "S")
    
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
        cmbIndexContactGroups.AddItem lvGroupList.List(i)
    Next i
    
    'mnuFile.Caption = LoadLang("����(&F)", "&File", "Archivo(&F)")
    'mnuView.Caption = LoadLang("����(&V)", "&View", "&Ver")
    'mnuHelp.Caption = LoadLang("����(&H)", "&Help", "Ayuda(&H)")
    
    'Me.mnuFileExit.Caption = LoadLang("���(&X)", "E&xit", "Salida(&X)")
    'mnuFileProperties.Caption = LoadLang("���� ���(&I)", "L&ist of Plans", "L&ista de horarios") & "..."
    'mnuFilePlanBrowser.Caption = LoadLang("��� ����/������ ����(&B)", "&Browse the Data", "Indice de datos(&B)") & "..."
    'mnuFileSave.Caption = LoadLang("����(&S)", "&Save", "Tienda(&S)")
    
    'mnuViewStatusBar.Caption = LoadLang("���� ǥ����(&S)", "&Status Bar", "Barra de e&stado")
    'mnuViewOptions.Caption = LoadLang("�ɼ�(&O)", "&Options", "Ambientaci&on")
    
    'mnuDateMenu.Caption = LoadLang("����(&P)", "&Plans", "&Planes")
    'mnuTodaysPlan.Caption = LoadLang("�̳��� ����(&T)", "Selec&ted Date's Plans", "&Planes de la fecha seleccionada")
    'mnuPlansClear.Caption = LoadLang("������ ��¥�� ���� ��� ����(&D)", "Clear selected &Date's Plans", "Borrar los planes &de la fecha seleccionada")
    
    ssRibbonMenu.TabCaption(1) = LoadLang("Ȩ", "Home", "Inicio")
    ssRibbonMenu.TabCaption(2) = LoadLang("����", "View", "Ver")
    ssRibbonMenu.TabCaption(3) = LoadLang("����", "Plan", "Planes")
    
    cmdPlanList.Caption = LoadLang("���� ���", "Plan List", "Lista de planes")
    cmdPlanIndex.Caption = LoadLang("������ ����", "Data Index", "Indice de datos")
    cmdEndPrg.Caption = LoadLang("������", "Exit", "Salida")
    
    tglStatusBar.Caption = LoadLang("����ǥ����", "Status Bar", "Barra de estado")
    tglCalWeekNum.Caption = LoadLang("�� ��ȣ", "Week Number", "Numero de la semana")
    
    cmdTodaysPlan.Caption = LoadLang("�̳��� ����", "Selected Day's Plans", "Planes del dia seleccionado")
    cmdDelAllTodaysPlan.Caption = LoadLang("�̳���   ���� ����", "Delete Plans", "Eliminar planes")
    
    cmdMnuAbout.ToolTipText = LoadLang("���α׷� ����", "About this application...")
    cmdMnuOptions.ToolTipText = LoadLang("ȯ�� ����", "Settings...")
    cmdHelp.ToolTipText = LoadLang("����", "Help")
    
    cmdRibbonFile.Caption = LoadLang("����", "File", "Archivo") & " ��"
    
    ssTodaysPlan.TabCaption(0) = LoadLang("����", "Today", "Hoy")
    ssTodaysPlan.TabCaption(1) = LoadLang("����", "Tomorrow", "Manana")
    
    SSTab1.TabCaption(0) = LoadLang("����", "Plans", "Planes")
    SSTab1.TabCaption(1) = LoadLang("�ּҷ�", "Contacts", "Contactos")
    SSTab1.TabCaption(2) = LoadLang("�� ��", "Tasks", "Tareas")
    SSTab1.TabCaption(3) = LoadLang("�ϰ�ǥ", "Schedule", "Calendario")
    SSTab1.TabCaption(4) = LoadLang("�˶�", "Alarms", "Alarmas")
    
    cmdTltRef.Caption = LoadLang("����(&R)", "&Refresh", "Actualiza&r")
    
    Frame1.Caption = LoadLang("�⺻ ����", "Basic Information", "Informacion basica")
    Frame2.Caption = LoadLang("��ȭ��ȣ", "Phone Numbers", "Numeros de telefono")
    Frame3.Caption = LoadLang("��Ÿ ����", "Other Informations", "Otra informacion")

    Label22.Caption = LoadLang("�޸�", "Note", "Nota") & ":"
    Label21.Caption = LoadLang("����", "Birthday", "Cumpleanos") & ":"
    
     Label1.Caption = LoadLang("�̸�", "Name", "Nombre") & ":"
     Label2.Caption = LoadLang("�޴���ȭ", "Cell-phone", "Celular") & ":"
     Label3.Caption = LoadLang("���ڿ���", "E-mail", "Correo electronico") & ":"
    Label17.Caption = LoadLang("�׷�", "Group", "Grupo") & ":"
    Label12.Caption = LoadLang("�����ȣ", "Postal", "Postal") & ":"
    Label13.Caption = LoadLang("�ּ�", "Address", "Direccion") & ":"
     Label4.Caption = LoadLang("��", "Home", "Casa") & ":"
     Label5.Caption = LoadLang("ȸ��", "Company", "Empresa") & ":"
     Label6.Caption = LoadLang("�ѽ�", "Fax", "Fax") & ":"
     Label7.Caption = LoadLang("��Ÿ", "Other", "Otros") & ":"
    
    cmdSaveContact.Caption = LoadLang("����(&S)", "&Save", "Tienda(&S)")
    cmdDelContact.Caption = LoadLang("����(&D)", "&Delete", "Eliminar(&D)")
    cmdDeleteAllContacts.Caption = LoadLang("��� ����(&E)", "Cl&ear contatcs", "Eliminar todo(&E)")
    cmdResetFields.Caption = LoadLang("���� �ʱ�ȭ(&R)", "&Reset Fields", "&Agregar")
    
    Frame4.Caption = LoadLang("�� �� ����", "Task Information")
    Label10.Caption = LoadLang("����", "Title") & ":"
    Label8.Caption = LoadLang("�Ϸ���", "Percent Complete") & ":"
    Label14.Caption = LoadLang("�߿䵵", "Importance") & ":"
    Label16.Caption = LoadLang("������", "Participants") & ":"
    Label11.Caption = LoadLang("�޸�", "Note") & ":"
    
    cmdSaveTask.Caption = LoadLang("����(&S)", "&Save", "Tienda(&S)")
    cmdDelTask.Caption = LoadLang("����", "&Delete", "Eliminar(&D)")
    cmdDeleteAllTasks.Caption = LoadLang("��� ����(&E)", "Cl&ear Tasks", "Eliminar todo(&E)")
    
    sdcmdSavePlanner.Caption = LoadLang("����(&S)", "&Save", "Tienda(&S)")
    
    Label18.Caption = LoadLang("�̸�", "Name", "Nombre") & ":"
    Label19.Caption = LoadLang("�ð�", "Time", "Tiempo") & ":"
    Label20.Caption = LoadLang("�޸�", "Note", "Nota") & ":"
    Frame5.Caption = LoadLang("����", "-", "-")
    
    chkDayOfWeeks(0).Caption = LoadLang("�Ͽ���", "Sunday", "Domingo")
    chkDayOfWeeks(1).Caption = LoadLang("������", "Monday", "Lunes")
    chkDayOfWeeks(2).Caption = LoadLang("ȭ����", "Tuesday", "Martes")
    chkDayOfWeeks(3).Caption = LoadLang("������", "Wednesday", "Miercoles")
    chkDayOfWeeks(4).Caption = LoadLang("�����", "Thursday", "Jueves")
    chkDayOfWeeks(5).Caption = LoadLang("�ݿ���", "Friday", "Viernes")
    chkDayOfWeeks(6).Caption = LoadLang("�����", "Saturday", "Sabado")
    
    cmdSelectAllDW.Caption = LoadLang("��� ����(&A)", "Select &All", "Seleccion&ar todo")
    cmdUnselectAllDW.Caption = LoadLang("���� ����(&L)", "Dese&lect All", "Dese&leccionar todo")
    cmdRelectAllDW.Caption = LoadLang("���� ����(&I)", "&Invert", "&Invertir seleccion")
    
    cmdResetAF.Caption = LoadLang("�ʱ�ȭ(&R)", "&Reset Fields", "&Restablecer")
    cmdDeleteAlarm.Caption = LoadLang("����(&D)", "&Delete", "Eliminar(&D)")
    cmdSaveAlarm.Caption = LoadLang("�߰�(&A)", "&Add", "&Agregar")
    
    Me.Caption = LoadLang(App.Title, "Plan Manager 3")
    
    MonthView1.Value = Split(Format(Now, "YYYY-M-D"), "-")(0) & "-" & Split(Format(Now, "YYYY-M-D"), "-")(1) & "-" & Split(Format(Now, "YYYY-M-D"), "-")(2)

    'Me.Show
    frmTip.Show vbModal, Me
    
    'Me.BorderStyle = 0
    Me.Caption = Me.Caption
    
    Me.height = 8290 '7995 '8500
    'Me.Width = 11400
    
    If GetSetting("Calendar", "Options", "AlwaysRibbon", 0) = 1 Then
        hideshow.Enabled = 0
        showrm
    End If
    
    Load frmMainOld
    frmMainOld.Hide
    
    Me.Hide
    If GetSetting("Calendar", "Options", "Password", "") = "" Then
        Me.Show
    Else
        frmConfirmPassword.Show 1, Me
    End If
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Container.Name = ssSidebar.Name Then
            Select Case theme
                Case 0
                    ctrl.BackColor = RGB(66, 156, 255)
                Case 1
                    ctrl.BackColor = RGB(248, 164, 24)
                Case 2
                    ctrl.BackColor = RGB(236, 220, 46)
            End Select
        End If
    Next ctrl
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "1"
End Sub

Private Sub hideshow_Timer()
    GetCursorPos pos
    'Label23.Caption = pos.x * 15 & " " & pos.y * 15 & " " & Screen.Width / 1280  ' 15
    
    If pos.y * 15 > Me.Top And pos.y * 15 < Me.Top + ssRibbonMenu.height + 500 And pos.x * 15 > Me.Left And pos.x * 15 < Me.Left + ssRibbonMenu.width Then
        showrm
    Else
        hiderm
    End If
End Sub

Private Sub imghover_click()
    cmdTltRef_Click
End Sub

Private Sub lblStatustexttext_Click()

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, _
    HTCAPTION, 0&)
    End If
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, _
    HTCAPTION, 0&)
    End If
End Sub

Private Sub imgDehover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgHover2.Visible = 0
End Sub

Private Sub imgHover_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPressed.Visible = -1
End Sub

Private Sub imgHover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgHover2.Visible = -1
End Sub

Private Sub imgHover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPressed.Visible = 0
    cmdTltRef_Click
End Sub

Private Sub imgHover2_Click()
    cmdTltRef_Click
End Sub



Private Sub lvAlarms_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    If Item.SubItems(1) = LoadLang("�� �˶� �߰�...", "New...") Then
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

Private Sub showrm()
    'Exit Sub
    timHidemenu.Enabled = False
    timHidemenu.Enabled = True
    ssRibbonMenu.Visible = True
    cmdRibbonFile.Visible = -1
    'cmdMnuOptions.Visible = -1
    'cmdHelp.Visible = -1
    'cmdMnuAbout.Visible = -1
    Image3(theme).Visible = -1
    'Image3.ZOrder 0
    Image1(theme).Visible = 0
    'Me.Height = 8500
    'Me.Width = 11500
    'Me.BorderStyle = 1
    'Me.Caption = Me.Caption
    SetWindowRgn hwnd, 0&, -1
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
        
    lblStatustext.Caption = "����Ǿ����ϴ�."
    'Sleep 1000
    'lblStatustext.Caption = ""
End Sub

Private Sub ssRibbonMenu_Click(PreviousTab As Integer)
    If ssRibbonMenu.Tab = 0 Then
        ssRibbonMenu.Tab = PreviousTab
    End If
End Sub

Private Sub ssTimerType_Click(PreviousTab As Integer)
    If timTImer.Enabled And ctimert <> ssTimerType.Tab Then
        ssTimerType.Tab = PreviousTab
    Else
        ctimert = ssTimerType.Tab
    End If
End Sub

Private Sub timAlarmChecker_Timer()
    '�˶��� ã�´�.
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
    '������ ������ ã�´�.
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
        '����ð��� �����ð����� ���̰� 10�� �̸��̸� �˸��� ����.
        'MsgBox Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(0) & Split(GetSetting("Calendar", yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(1) & " " & Format(Now, "hhmm") & " " & ttt
        If ttt < 10 And ttt >= -1 Then
            '��� ���� ������ �˸�
            If GetSetting("Calendar", "NotifiedPlans\" & yy & "\" & mm & "\" & dd, lvTodaysPlans.List(Plan), "abc") = "abc" Then
                'MsgBox 3
                frmReminder.yy = yy
                frmReminder.mm = mm
                frmReminder.dd = dd
                frmReminder.lblTitle.Caption = lvTodaysPlans.List(Plan)
                frmReminder.lblLoca.Caption = GetSetting("Calendar", yy & mm & dd, lvTodaysPlans.List(Plan) & "Location", "�ּ� �Һи�")
                frmReminder.txtContent.Text = GetSetting("Calendar", yy & mm & dd, lvTodaysPlans.List(Plan) & "Cont", "")
                frmReminder.Show
                'SysTray.ShowBalloonTip lvTodaysPlans.List(Plan) & " ���� ���۱��� 10�к��� ���� ���ҽ��ϴ�. �غ��Ͻʽÿ�.", beInformation, "����������"
                'Beep 950, 5
            End If
        End If
    Next Plan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Confirm(LoadLang("���������ڸ� ������ ���� ���� �˸��� ���� �ʽ��ϴ�.", "You will not be notified when you close the program.", "No se le notificara sobre los planes cuando se cierre el programa."), LoadLang("���", "Warning", "Advertencia"), Me, 48) = True Then
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
    
    If ContactChanged Then
        'If MsgBox("���� ���� ���� ����ó�� �������� ���� ������ �ֽ��ϴ�.", 48 + vbOKCancel, "���") = vbCancel Then Exit Sub
    End If
    
    If lvContacts.List(lvContacts.ListIndex) = LoadLang("�� ����ó �߰�...", "New...") Then
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
        cmbIndexContactGroups.ListIndex = 0
        
        cmdDelContact.Enabled = False
        cmdDuplicateContact.Enabled = False
        
        chkFavorite.Value = 0
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
        
        chkFavorite.Value = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Fav", 0)
        
        Dim i As Integer
        
        For i = 0 To cmbGroup.ListCount - 1
            If cmbGroup.List(i) = GetSetting("Calendar", "Contacts", lvContacts.List(lvContacts.ListIndex) & "Group", "") Then
                cmbGroup.ListIndex = i
                Exit For
            End If
        Next i
        
        cmdDelContact.Enabled = True
        cmdDuplicateContact.Enabled = True
        
        'Me.Caption = App.Title & " - " & SSTab1.TabCaption(SSTab1.Tab) & " (" & txtName.Text & ")"
    End If
End Sub

Private Sub lvTasks_Click()
    If lvTasks.List(lvTasks.ListIndex) = LoadLang("�� �۾� �߰�...", "New...") Then
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
    If lvTasks.List(Item) <> LoadLang("�� �۾� �߰�...", "New...") Then
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


    '�� ������Ʈ�� ���� ���� ������ ������ ����ڿ��� �޽����� ǥ���մϴ�.
    '����ڴ� [������Ʈ �Ӽ�] ��ȭ ���ڿ��� ���� ���α׷��� ����
    '���� ������ ������ �� �ֽ��ϴ�.
    If Len(App.HelpFile) = 0 Then
        MessageBox "���� ������ ǥ���� �� �����ϴ�. �� ���α׷��� ������ ������ �����ϴ�.", App.Title, Me, 16
    Else
        On Error Resume Next
        If FileExists(Dir1.Path & "\PLNMGR32.HLP") = False Then
            MessageBox "���� ������ ã�� �� �����ϴ�. Ǯ�׸� ����ȭ�� ��ο� PLNMGR32.HLP�� �ִ��� Ȯ���Ͻʽÿ�. ������ �ٽ� ��ġ�ϰų� ����꿡�� �޾� �����Ͻʽÿ�.", "����", Me, 16
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


    '�� ������Ʈ�� ���� ���� ������ ������ ����ڿ��� �޽����� ǥ���մϴ�.
    '����ڴ� [������Ʈ �Ӽ�] ��ȭ ���ڿ��� ���� ���α׷��� ����
    '���� ������ ������ �� �ֽ��ϴ�.
    If Len(App.HelpFile) = 0 Then
        MessageBox "���� ������ ǥ���� �� �����ϴ�. �� ���α׷��� ������ ������ �����ϴ�.", App.Title, Me, 16
    Else
        On Error Resume Next
        If FileExists(Dir1.Path & "\PLNMGR32.HLP") = False Then
            MessageBox "���� ������ ã�� �� �����ϴ�. Ǯ�׸� ����ȭ�� ��ο� PLNMGR32.HLP�� �ִ��� Ȯ���Ͻʽÿ�. ������ �ٽ� ��ġ�ϰų� ����꿡�� �޾� �����Ͻʽÿ�.", "����", Me, 16
            Exit Sub
        End If
        
        nRet = OSWinHelp(Me.hwnd, Dir1.Path & "\PLNMGR32.HLP", 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuToolsOptions_Click()
    '�۾�: 'mnuToolsOptions_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuToolsOptions_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewWebBrowser_Click()
    '�۾�: 'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewOptions_Click()
    '�۾�: 'mnuViewOptions_Click' �ڵ带 �߰��Ͻʽÿ�.
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    '�۾�: 'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewStatusBar_Click()
    'mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    'sbStatusBar.Visible = mnuViewStatusBar.Checked
    
    If sbStatusBar.Visible Then
        Me.height = 7080
    Else
        Me.height = 6810
    End If
End Sub

'Private Sub mnuViewToolbar_Click()
    'mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
'End Sub

Private Sub mnuEditPasteSpecial_Click()
    '�۾�: 'mnuEditPasteSpecial_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditPasteSpecial_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditPaste_Click()
    '�۾�: 'mnuEditPaste_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditPaste_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditCopy_Click()
    '�۾�: 'mnuEditCopy_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditCopy_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditCut_Click()
    '�۾�: 'mnuEditCut_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditCut_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditUndo_Click()
    '�۾�: 'mnuEditUndo_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditUndo_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSend_Click()
    '�۾�: 'mnuFileSend_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSend_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFilePrint_Click()
    '�۾�: 'mnuFilePrint_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFilePrint_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFilePrintPreview_Click()
    '�۾�: 'mnuFilePrintPreview_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFilePrintPreview_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

'Private Sub mnuFilePageSetup_Click()
'    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "������ ����"
'        .CancelError = True
'        .ShowPrinter
'    End With
'
'End Sub

Private Sub mnuFileProperties_Click()
    '�۾�: 'mnuFileProperties_Click' �ڵ带 �߰��Ͻʽÿ�.
    MonthView1_DateClick MonthView1.SelStart
End Sub

Private Sub mnuFileSaveAll_Click()
    '�۾�: 'mnuFileSaveAll_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSaveAll_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSaveAs_Click()
    '�۾�: 'mnuFileSaveAs_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSaveAs_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSave_Click()
    '�۾�: 'mnuFileSave_Click' �ڵ带 �߰��Ͻʽÿ�.
    If SSTab1.Tab = 1 Then
        cmdSaveContact_Click
    Else
        cmdSaveTask_Click
    End If
End Sub

Private Sub mnuFileClose_Click()
    '�۾�: 'mnuFileClose_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileClose_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

'Private Sub mnuFileOpen_Click()
'    Dim sFile As String
'
'
'    With dlgCommonDialog
'        .DialogTitle = "����"
'        .CancelError = False
'        '�۾�: Common Dialog ��Ʈ���� �÷��׿� Ư���� �����մϴ�.
'        .Filter = "��� ����(*.*)|*.*"
'        .ShowOpen
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    '�۾�: �ڵ带 �߰��Ͽ� ���� �ִ� ������ ó���մϴ�.
'
'End Sub

Private Sub mnuFileNew_Click()
    '�۾�: 'mnuFileNew_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileNew_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'mnuTodaysPlan.Caption = MonthView1.SelStart & "�� ����"
        'PopupMenu mnuDateMenu
    End If
End Sub

Private Sub ssRibbonMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab > 0 Then
        'mnuFileBar0.Visible = True
        'mnuFileSave.Visible = True
    Else
        'mnuFileBar0.Visible = False
        'mnuFileSave.Visible = False
    End If
    
    ssSidebar.Tab = SSTab1.Tab
    
    cmdTltRef_Click
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

Private Sub hiderm()
    'Exit Sub
    ssRibbonMenu.Visible = 0
    cmdRibbonFile.Visible = 0
    'cmdMnuOptions.Visible = 0
    'cmdHelp.Visible = 0
    'cmdMnuAbout.Visible = 0
    Image3(theme).Visible = 0
    Image1(theme).Visible = -1
    
    'Me.BorderStyle = 0
    'Me.Caption = Me.Caption
    
    Me.ScaleMode = 3
    Dim width As Integer
    Dim height As Integer
    Dim bw As Integer
    Dim th As Integer
    Dim rgn As Long
    width = Me.ScaleX(Me.width, 1, 3)
    height = Me.ScaleY(Me.height, 1, 3)
    bw = (width - Me.ScaleWidth) / 2
    th = height - bw - Me.ScaleHeight
    rgn = CreateRectRgn(0 + bw, 0 + th, width - bw, height - bw)
    SetWindowRgn Me.hwnd, rgn, -1
    Me.ScaleMode = 1
End Sub

Private Sub timTImer_Timer()
    If ssTimerType.Tab Then '��ž��ġ
        If lblTimerSD.Caption = 9 Then
            lblTimerSD.Caption = 0
            lblTimerS.Caption = lblTimerS.Caption + 1
        End If
        
        If lblTimerS.Caption = 60 Then
            lblTimerS.Caption = 0
            lblTimerM.Caption = lblTimerM.Caption + 1
        End If
        
        If lblTimerM.Caption = 60 Then
            lblTimerM.Caption = 0
            lblTimerH.Caption = lblTimerH.Caption + 1
        End If
        
        lblTimerSD.Caption = lblTimerSD.Caption + 1
        
        SetTimerLabel
    Else 'Ÿ�̸�
        If pbTimer.Value = pbTimer.Max Or (lblTimerH.Caption = 0 And lblTimerM.Caption = 0 And lblTimerS.Caption = 0 And lblTimerSD.Caption = 0) Then
            timTImer.Enabled = 0
            
            lblTimerH.Caption = 0
            lblTimerM.Caption = 0
            lblTimerS.Caption = 0
            lblTimerSD.Caption = 0
            SetTimerLabel
            
            If chkEnablesw.Value Then
                MessageBox "Ÿ�̸� ��!", "�˸�", Me, 48, 3000, False
            Else
                MessageBox "Ÿ�̸� ��!", "�˸�", Me, 48
            End If
            
            cmdTImerStop_Click
            pbTimer.Value = 0
            
            If chkEnablesw.Value Then
                ssTimerType.Tab = 1
                lbltsw.Visible = -1
                cmdTimerStart_Click
            End If
            
            Exit Sub
        End If
        
        If lblTimerSD.Caption = 0 Then
            lblTimerSD.Caption = 9
            lblTimerS.Caption = lblTimerS.Caption - 1
        End If
        
        If lblTimerS.Caption = -1 Then
            lblTimerS.Caption = 59
            lblTimerM.Caption = lblTimerM.Caption - 1
        End If
        
        If lblTimerM.Caption = -1 And lblTimerH.Caption Then
            lblTimerM.Caption = 59
            lblTimerH.Caption = lblTimerH.Caption - 1
        End If
        
        lblTimerSD.Caption = lblTimerSD.Caption - 1
        
        SetTimerLabel
        
        pbTimer.Value = pbTimer.Value + 1
    End If
End Sub

Private Sub txtAddress_Change()
    ContactDataChanged
End Sub

Private Sub txtBDay_Change()
    ContactDataChanged
End Sub

Private Sub txtBMonth_Change()
    ContactDataChanged
End Sub

Private Sub txtBYear_Change()
    ContactDataChanged
End Sub

Private Sub txtCellPhone_Change()
    ContactDataChanged
End Sub

Private Sub txtCompany_Change()
    ContactDataChanged
End Sub

Private Sub txtContent_Change()
    ContactDataChanged
End Sub

Private Sub txtEmail_Change()
    ContactDataChanged
End Sub

Private Sub txtFax_Change()
    ContactDataChanged
End Sub

Private Sub txtHome_Change()
    ContactDataChanged
End Sub

Private Sub txtName_Change()
    ContactDataChanged
End Sub

Private Sub txtOtherNumber_Change()
    ContactDataChanged
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
    
    lvAlarms.ColumnHeaders.Add , , LoadLang("�ð�", "Time"), 350
    lvAlarms.ColumnHeaders.Add , , LoadLang("�̸�", "Name"), 1400

    lvAlarms.ListItems.Add , , "--:--"
    lvAlarms.ListItems(1).SubItems(1) = LoadLang("�� �˶� �߰�...", "New...")
    
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

Private Sub txtPostalCode_Change()
    ContactDataChanged
End Sub

Private Sub txtTimeHrs_Change()
    On Error Resume Next
    If Len(txtTimeHrs.Text) = 2 Or (txtTimeHrs.Text >= 3 And txtTimeHrs.Text <= 9) Then
        txtTimeMin.SetFocus '�� �Է� ĭ�� ä��� ���� ĭ�� Ȱ��ȭ�Ѵ�.
    End If
End Sub

Private Sub txtTimeMin_Change()
    If txtTimeMin.Text = "" Then txtTimeHrs.SetFocus
End Sub



