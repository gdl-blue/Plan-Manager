VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "ȯ�� ����"
   ClientHeight    =   4350
   ClientLeft      =   -75
   ClientTop       =   3000
   ClientWidth     =   8175
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOptionHelp 
      Caption         =   "����(&H)..."
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "���"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ��"
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
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483636
      TabCaption(0)   =   "��Ÿ��"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkAlwaysRm"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNoRibbon"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkTP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "����� ������"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDelPlans"
      Tab(1).Control(1)=   "cmdDelContacts"
      Tab(1).Control(2)=   "cmdDelTasks"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "ǥ��"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1"
      Tab(2).Control(1)=   "cmbStartPage"
      Tab(2).Control(2)=   "radSelST"
      Tab(2).Control(3)=   "radCFQ"
      Tab(2).Control(4)=   "Frame11"
      Tab(2).Control(5)=   "Frame4"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "�˻�"
      TabPicture(3)   =   "frmOptions.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkNoTimeCHeck"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(2)=   "Label9"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "����� �з�"
      TabPicture(4)   =   "frmOptions.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdDelGroup"
      Tab(4).Control(1)=   "cmdAddNewGroup"
      Tab(4).Control(2)=   "cmdClearGroups"
      Tab(4).Control(3)=   "txtNewGroup"
      Tab(4).Control(4)=   "lvGroups"
      Tab(4).Control(5)=   "lvCustomCates"
      Tab(4).Control(6)=   "cmdClearCates"
      Tab(4).Control(7)=   "cmdDelSelCate"
      Tab(4).Control(8)=   "cmdAddNewCate"
      Tab(4).Control(9)=   "txtCategory"
      Tab(4).Control(10)=   "Label14"
      Tab(4).Control(11)=   "Label11"
      Tab(4).Control(12)=   "Label8"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmOptions.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "����"
      TabPicture(6)   =   "frmOptions.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdChangePassword"
      Tab(6).Control(1)=   "chkPasswordRequired"
      Tab(6).Control(2)=   "Frame5"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "�Ҹ�"
      TabPicture(7)   =   "frmOptions.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdPlayRT"
      Tab(7).Control(1)=   "cmdPlayNS"
      Tab(7).Control(2)=   "Frame34"
      Tab(7).Control(3)=   "Frame12"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   " "
      TabPicture(8)   =   "frmOptions.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "��޼���"
      TabPicture(9)   =   "frmOptions.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cmdApplyAdvanced"
      Tab(9).Control(1)=   "txtAdvancedSetting"
      Tab(9).Control(2)=   "Frame15"
      Tab(9).ControlCount=   3
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "����(&C)"
         Height          =   375
         Left            =   -70200
         TabIndex        =   94
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdApplyAdvanced 
         Appearance      =   0  '���
         Caption         =   "����"
         Height          =   255
         Left            =   -69960
         TabIndex        =   93
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkNoTimeCHeck 
         Caption         =   "���� �߰� �� �ð��� �ùٸ��� �˻� ����(&T)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   92
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "���� �� '�˰� ��ʴϱ�' ǥ��(&P)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   2640
         Width           =   2895
      End
      Begin VB.ComboBox cmbStartPage 
         Height          =   300
         Left            =   -74400
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   90
         Top             =   1560
         Width           =   5295
      End
      Begin VB.OptionButton radSelST 
         Caption         =   "ȭ�� ����(&T)"
         Height          =   255
         Left            =   -74640
         TabIndex        =   89
         Top             =   1200
         Width           =   5175
      End
      Begin VB.OptionButton radCFQ 
         Caption         =   "������ ���� �������� ����(&Q)"
         Height          =   255
         Left            =   -74640
         TabIndex        =   88
         Top             =   1920
         Width           =   5295
      End
      Begin VB.CommandButton cmdDelPlans 
         Caption         =   "��� ����(&D)"
         Height          =   375
         Left            =   -70320
         TabIndex        =   87
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelContacts 
         Caption         =   "��� ����(&E)"
         Height          =   375
         Left            =   -70320
         TabIndex        =   86
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelTasks 
         Caption         =   "��� ����(&L)"
         Height          =   375
         Left            =   -70320
         TabIndex        =   85
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chkTP 
         Caption         =   "���������� �����(&O)"
         Height          =   255
         Left            =   3960
         TabIndex        =   84
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkNoRibbon 
         Caption         =   "���� �޴� ��Ȱ��(&N)"
         Height          =   255
         Left            =   4080
         TabIndex        =   83
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkAlwaysRm 
         Caption         =   "�޴� �׻� ���̱�(&A)"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   960
         Width           =   2055
      End
      Begin VB.Frame Frame7 
         Caption         =   "���� �޴�"
         Height          =   735
         Left            =   120
         TabIndex        =   80
         Top             =   3240
         Width           =   3615
         Begin VB.ComboBox cmbBGColor 
            Height          =   300
            Left            =   120
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   81
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "�׸�"
         Height          =   735
         Left            =   3840
         TabIndex        =   78
         Top             =   3240
         Width           =   3735
         Begin VB.ComboBox cmbThemeSelect 
            Height          =   300
            Left            =   120
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   79
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.ComboBox txtAdvancedSetting 
         Height          =   300
         Left            =   -74280
         TabIndex        =   75
         Top             =   680
         Width           =   2175
      End
      Begin VB.Frame Frame15 
         Caption         =   "�׸�:    "
         Height          =   2055
         Left            =   -74880
         TabIndex        =   74
         Top             =   720
         Width           =   5895
         Begin VB.TextBox txtAdvancedValue 
            Height          =   270
            Left            =   240
            TabIndex        =   77
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label6 
            Caption         =   "������:"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "���"
         Height          =   615
         Left            =   6120
         TabIndex        =   64
         Top             =   2520
         Width           =   1455
         Begin VB.ComboBox cmbLanguage 
            Height          =   300
            Left            =   120
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   65
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�׸�"
         Height          =   735
         Left            =   1920
         TabIndex        =   61
         Top             =   4560
         Width           =   7455
         Begin VB.CommandButton cmdTheSet 
            Caption         =   "�׸�(&T)..."
            Height          =   375
            Left            =   6120
            TabIndex        =   62
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "�׸��� �����Ϸ��� ���� ���߸� �����ʽÿ�."
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   300
            Width           =   5895
         End
      End
      Begin VB.CommandButton cmdPlayRT 
         Caption         =   "���(&R)"
         Height          =   320
         Left            =   -70680
         TabIndex        =   60
         Top             =   3080
         Width           =   1335
      End
      Begin VB.CommandButton cmdPlayNS 
         Caption         =   "���(&N)"
         Height          =   320
         Left            =   -70680
         TabIndex        =   59
         Top             =   1640
         Width           =   1335
      End
      Begin VB.CheckBox chkPasswordRequired 
         Caption         =   "���α׷��� ������ �� ��ȣ �Է� �ʿ�"
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   720
         Width           =   3255
      End
      Begin VB.Frame Frame34 
         Caption         =   "���� �˸���"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   5895
         Begin VB.Frame Frame9 
            BorderStyle     =   0  '����
            Height          =   975
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   5655
            Begin VB.PictureBox grpNotificationContainer 
               Appearance      =   0  '���
               BorderStyle     =   0  '����
               ForeColor       =   &H80000008&
               Height          =   1300
               Left            =   0
               ScaleHeight     =   1305
               ScaleWidth      =   5415
               TabIndex        =   50
               Top             =   0
               Width           =   5415
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "����� 1"
                  Height          =   495
                  Index           =   5
                  Left            =   3600
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "�ߺ�- ��- �ߺ�-"
                  Height          =   495
                  Index           =   4
                  Left            =   1800
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "��- ��- ��-"
                  Height          =   495
                  Index           =   3
                  Left            =   0
                  TabIndex        =   68
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "�ߺ�- ���� 3"
                  Height          =   495
                  Index           =   8
                  Left            =   3600
                  TabIndex        =   73
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "����� 2"
                  Height          =   495
                  Index           =   7
                  Left            =   1800
                  TabIndex        =   72
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "��- �ߺ�-"
                  Height          =   495
                  Index           =   6
                  Left            =   0
                  TabIndex        =   71
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "�ߺ�- �ߺ�-"
                  Height          =   255
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   67
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "������-"
                  Height          =   255
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   57
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton optNotificationSound 
                  Caption         =   "��- ��-"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   51
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   975
               Left            =   5400
               Max             =   1
               TabIndex        =   49
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "�˶���"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   47
         Top             =   2160
         Width           =   5895
         Begin VB.VScrollBar VScroll2 
            Enabled         =   0   'False
            Height          =   975
            Left            =   5520
            Max             =   1
            TabIndex        =   53
            Top             =   240
            Width           =   255
         End
         Begin VB.Frame Frame13 
            BorderStyle     =   0  '����
            Caption         =   "Frame13"
            Height          =   975
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   5655
            Begin VB.PictureBox grpRingtoneContainer 
               Appearance      =   0  '���
               BorderStyle     =   0  '����
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   0
               ScaleHeight     =   975
               ScaleWidth      =   5415
               TabIndex        =   54
               Top             =   0
               Width           =   5415
               Begin VB.OptionButton optRingtone 
                  Caption         =   "�Ʊ���� �Ѹ�"
                  Height          =   255
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   66
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.OptionButton optRingtone 
                  Caption         =   "�����"
                  Height          =   255
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton optRingtone 
                  Caption         =   "�⺻��"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   55
                  Top             =   0
                  Width           =   1815
               End
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "�޷�"
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   5895
         Begin VB.ComboBox cmbWSD 
            Height          =   300
            Left            =   1440
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   44
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label5 
            Caption         =   "���� ����:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " "
         Height          =   2535
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   6015
         Begin VB.TextBox txtConfirmPassword 
            Height          =   270
            IMEMode         =   3  '��� ����
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   42
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox txtNewPassword 
            Height          =   270
            IMEMode         =   3  '��� ����
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   40
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCurrentPassword 
            Height          =   270
            IMEMode         =   3  '��� ����
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   38
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label17 
            Caption         =   "��й�ȣ Ȯ��:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label16 
            Caption         =   "�� ��й�ȣ:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label15 
            Caption         =   "���� ��й�ȣ:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdDelGroup 
         Caption         =   "���� �׷� ����"
         Height          =   375
         Left            =   -72600
         TabIndex        =   35
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNewGroup 
         Caption         =   "�Է� �߰�(&D)"
         Height          =   375
         Left            =   -68760
         TabIndex        =   34
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearGroups 
         Caption         =   "�׷� ��ü����"
         Height          =   375
         Left            =   -70320
         TabIndex        =   33
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtNewGroup 
         Height          =   270
         Left            =   -74880
         TabIndex        =   30
         Top             =   3720
         Width           =   6015
      End
      Begin VB.FileListBox lvGroups 
         Height          =   1350
         Left            =   -72600
         TabIndex        =   29
         Top             =   960
         Width           =   2175
      End
      Begin VB.FileListBox lvCustomCates 
         Height          =   1350
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ʱ�ȭ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdPrgReset 
            Caption         =   "�ʱ�ȭ(&R)"
            Height          =   375
            Left            =   4560
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblResetN1 
            Caption         =   "������ �ʱ�ȭ"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblResetCount 
            Caption         =   "7"
            Height          =   255
            Left            =   1320
            TabIndex        =   26
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "���α׷� ��ü �����͸� �ʱ�ȭ�մϴ�."
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblResetN2 
            Caption         =   "�ܰ� ���Դϴ�."
            Height          =   255
            Left            =   1440
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "�� ����"
         Height          =   615
         Left            =   -74880
         TabIndex        =   21
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Frame Frame8 
         Caption         =   "���̾ƿ�"
         Height          =   1695
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   5895
      End
      Begin VB.CommandButton cmdClearCates 
         Caption         =   "�з� ��ü����"
         Height          =   375
         Left            =   -70320
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "�� �˻� ����"
         Height          =   615
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton cmdDelSelCate 
         Caption         =   "���� �з� ����"
         Height          =   375
         Left            =   -74880
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNewCate 
         Caption         =   "�Է� �߰�(&A)"
         Height          =   375
         Left            =   -68760
         TabIndex        =   15
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtCategory 
         Height          =   270
         Left            =   -74880
         TabIndex        =   14
         Top             =   3120
         Width           =   6015
      End
      Begin VB.Frame Frame4 
         Caption         =   "����"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   6015
         Begin VB.Label Label7 
            Caption         =   "���� ȭ��:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�� ������"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   6015
         Begin VB.FileListBox lvTaskFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   6
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.FileListBox lvContactFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   5
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.FileListBox lvPlanFiles 
            Height          =   270
            Left            =   3240
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "�� �۾����:"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "�� �ּҷ�:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "�� ����:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Label Label14 
         Caption         =   "Categorias:                      Grupo:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label11 
         Caption         =   "�� �׷� �߰�:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "[*] �� ������ �����ϸ� ���α׷��� �ùٷ� �۵����� ���� �� �ֽ��ϴ�."
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   3720
         Width           =   7335
      End
      Begin VB.Label Label8 
         Caption         =   "�� ���� �з� �߰�:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   2880
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ResetCount As Integer

'����� �ܺμҽ�
'http://www.vbforums.com/showthread.php?617573-RESOLVED-Scroll-bars-for-frame-inside-a-tab

Option Explicit
Dim lngOriginalTop         As Long
Dim lngIncrement           As Long
Dim lngOriginalTop2        As Long
Dim lngIncrement2          As Long

Dim RTI As Integer
Dim NSI As Integer

Dim Loaded As Boolean

Private Sub Check2_Click()

End Sub

Private Sub chkPasswordRequired_Click()
    If chkPasswordRequired.Value = 0 And GetSetting("Calendar", "Options", "Password", "") <> "" Then
        frmCheckDeactivatePassword.Show vbModal, Me
    End If
    
    Dim ctrl As Control
    If chkPasswordRequired.Value = 0 Then
        For Each ctrl In Me.Controls
            If ctrl.Container.Name = Frame5.Name Then
                ctrl.Enabled = False
            End If
        Next ctrl
    Else
        For Each ctrl In Me.Controls
            If ctrl.Container.Name = Frame5.Name Then
                ctrl.Enabled = True
            End If
        Next ctrl
        
        If GetSetting("Calendar", "Options", "Password", "") = "" Then
            txtCurrentPassword.Enabled = 0
            Label15.Enabled = 0
        End If
    End If
End Sub

Private Sub cmdAddNewCate_Click()
    If txtCategory.Text <> "(�������� ����)" Then
        On Error Resume Next
        
        MkDir "C:\CALPLANS"
        MkDir "C:\CALPLANS\CTGORIES"
        
        If Len(txtCategory.Text) < 1 Then
            MessageBox "�з��� ������ �Է����ֽʽÿ�.", "����", Me, 16
            Exit Sub
        End If
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        Dim iFileNo As Integer
        iFileNo = FreeFile
        '������ ����.
        
        Open "C:\CALPLANS\CTGORIES\" & txtCategory.Text For Output As #iFileNo
        
        '������ ������ ���� �����Ƿ� �� ĭ����...
        Print #iFileNo, ""
        
        '������ �ݴ´�.
        Close #iFileNo
        
        lvCustomCates.Refresh
    Else
        MessageBox "�̹� �����ϰų� �ùٸ��� �ʽ��ϴ�.", "����", Me, 16
    End If
End Sub

Private Sub cmdCalSet_Click()
    SSTab1.Tab = 7
End Sub

Private Sub cmdApplyAdvanced_Click()
    If txtAdvancedSetting.Text = "" Then Exit Sub
    If UCase(txtAdvancedSetting.Text) = "PASSWORD" Then Exit Sub
    SaveSetting "Calendar", "Options", txtAdvancedSetting.Text, txtAdvancedValue.Text
End Sub

Private Sub cmdChangePassword_Click()
    If GetSetting("Calendar", "Options", "Password", "") <> txtCurrentPassword.Text Then
        MsgBox "���� ��ȣ�� �ùٸ��� �ʽ��ϴ�.", 16, "��ȣ"
    ElseIf txtConfirmPassword.Text <> txtNewPassword.Text Then
        MsgBox "��ȣ Ȯ���� ��ġ�ϱ� �ʽ��ϴ�.", 16, "��ȣ"
    ElseIf txtConfirmPassword.Text = "" Then
        MsgBox "��ȣ�� ���� �ʼ��Դϴ�.", 16, "��ȣ"
    Else
        SaveSetting "Calendar", "Options", "Password", txtNewPassword.Text
        MsgBox "��ȣ�� ����Ǿ����ϴ�.", 64, "��ȣ"
        txtCurrentPassword.Enabled = -1
        Label15.Enabled = -1
        txtConfirmPassword.Text = ""
        txtNewPassword.Text = ""
        txtCurrentPassword.Text = ""
    End If
End Sub

Private Sub cmdClearCates_Click()
    If Confirm("������ " & lvCustomCates.ListCount & "���� �з��� *���* �����Ͻðڽ��ϱ�?", "����", Me) Then
        On Error Resume Next
        Dim i As Integer
        For i = 0 To lvCustomCates.ListCount - 1
            Kill "C:\CALPLANS\CTGORIES\" & lvCustomCates.List(i)
        Next i
        
        lvCustomCates.Refresh
        MessageBox "��� �����Ǿ����ϴ�.", "����", Me, 48
    End If
End Sub

Private Sub cmdClearGroups_Click()
    If Confirm("������ " & lvGroups.ListCount & "���� �׷��� *���* �����Ͻðڽ��ϱ�?", "����", Me) Then
        On Error Resume Next
        Dim i As Integer
        For i = 0 To lvGroups.ListCount - 1
            Kill "C:\CALPLANS\CTGROUPS\" & lvGroups.List(i)
        Next i
        
        lvGroups.Refresh
        MessageBox "��� �����Ǿ����ϴ�.", "����", Me, 48
    End If
End Sub

Sub cmdDelContacts_Click()
    If Confirm("������ �����Ͻðڽ��ϱ�?", "����", Me) Then
        If Confirm("���� *�Ұ���*�մϴ�. ������ ��� �ּҷ��� �����Ͻðڽ��ϱ�?", "����", Me, , True) Then
            On Error Resume Next
            lvContactFiles.Path = "C:\CALPLANS\CONTACTS"
            
            Dim contact As Integer
            Dim ContactName As String
            For contact = 0 To lvContactFiles.ListCount - 1
                ContactName = lvContactFiles.List(contact)
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
            Next contact
            
            frmMain.LoadContacts
            
            MessageBox "�ּҷ� ����Ÿ�� ��� �����ƽ��ϴ�.", "����", Me, 64
        End If
    End If
End Sub

Private Sub cmdDelPlans_Click()
    Dim DelYear As String
    DelYear = InputBox("������ ������ �Է��Ͻʽÿ�.", "���� ��� �����")
    If DelYear <> "" Then
        If IsNumeric(DelYear) = False Then
            MessageBox "������ ���� �ùٸ��� �ʽ��ϴ�.", "����", Me, 16
            Exit Sub
        End If
    
        On Error Resume Next
        If Confirm("������ �����Ͻðڽ��ϱ�?", "����", Me) Then
            If Confirm("���� *�Ұ���*�մϴ�. ������ " & DelYear & "���� ��� ������ �����Ͻðڽ��ϱ�?", "����", Me, , True) Then
                On Error Resume Next
                Shell "CMD /C RD /S /Q " & ChrW$(34) & "C:\CALPLANS\" & DelYear & ChrW$(34)
                Shell "COMMAND /C DELTREE /Y " & ChrW$(34) & "C:\CALPLANS\" & DelYear & ChrW$(34)
            End If
        End If
    End If
End Sub

Private Sub cmdDelSelCate_Click()
    On Error Resume Next
    Kill "C:\CALPLANS\CTGORIES\" & lvCustomCates.List(lvCustomCates.ListIndex)
    
    lvCustomCates.Refresh
End Sub

Sub cmdDelTasks_Click()
    If Confirm("������ �����Ͻðڽ��ϱ�?", "����", Me) Then
        If Confirm("���� *�Ұ���*�մϴ�. ������ ��� �۾��� �����Ͻðڽ��ϱ�?", "����", Me, , True) Then
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
            
            MessageBox "�۾���� ����Ÿ�� ��� �����ƽ��ϴ�.", "����", Me
        End If
    End If
    
    frmMain.LoadTasks
End Sub

Private Sub cmdOptionHelp_Click()
    MessageBox "������ ���õ� ������ �����ϴ�.", "����", Me, 16
End Sub

Private Sub cmdPlayNS_Click()
    PlayNotification NSI
End Sub

Private Sub cmdPlayRT_Click()
    PlayRingtone RTI
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
    
    SaveSetting "Calendar", "Options", "Theme2", cmbThemeSelect.ListIndex
    
    SaveSetting "Calendar", "Options", "TP", chkTP.Value
    
    SaveSetting "Calendar", "Options", "NoRibbon", chkNoRibbon.Value
    
    SaveSetting "Calendar", "Options", "AlwaysRibbon", chkAlwaysRm.Value
    
    SaveSetting "Calendar", "Options", "Language", cmbLanguage.ListIndex
    
    If radSelST.Value = False Then
        SaveSetting "Calendar", "Options", "SST", False
    Else
        SaveSetting "Calendar", "Options", "SST", True
    End If
    
    If GetSetting("Calendar", "Options", "TP", 0) = 1 Then
        frmMain.width = 8715
    Else
        frmMain.width = 11040
    End If
    
    Dim i As Integer
    
    SaveSetting "Calendar", "Options", "Notification", NSI
    
    SaveSetting "Calendar", "Options", "Ringtone", RTI
    
    If Confirm(LoadLang("������ ���������� ����Ǿ����� ȿ���� �����Ϸ��� ���α׷��� ������ؾ� �մϴ�. ���α׷��� �����մϴ�.", "You must restart the application to take effect.", "La configuracion se ha aplicado correctamente y debe reiniciar el programa para que surta efecto. Salga del programa."), LoadLang("�˸�", "Information", "Informacion"), Me, 48) Then
        End
    End If
    
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
        If Confirm("������ ���. ������ ���α׷� ��ü�� �ʱ�ȭ�Ͻðڽ��ϱ�?", "�ʱ�ȭ", Me) Then
            If Confirm("��ǻ�Ͱ� Windows 95, 98 Ȥ�� Me�� �������Դϱ�?", "����", Me) Then
                Shell "COMMAND /C RD /S /Q C:\CALPLANS"
            Else
                Shell "CMD /C RD /S /Q C:\CALPLANS"
            End If
            MessageBox "�ʱ�ȭ �Ϸ�. ���α׷��� �����մϴ�...", "�ʱ�ȭ", Me
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

Private Sub cmdAddNewGroup_Click()
    If txtNewGroup.Text <> "���� �� ��" Then
        On Error Resume Next
        
        MkDir "C:\CALPLANS"
        MkDir "C:\CALPLANS\CTGROUPS"
        
        If Len(txtNewGroup.Text) < 1 Then
            MessageBox "�׷��� �̸��� �Է����ֽʽÿ�.", "����", Me, 16
            Exit Sub
        End If
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        Dim iFileNo As Integer
        iFileNo = FreeFile
        '������ ����.
        
        Open "C:\CALPLANS\CTGROUPS\" & txtNewGroup.Text For Output As #iFileNo
        
        '������ ������ ���� �����Ƿ� �� ĭ����...
        Print #iFileNo, ""
        
        '������ �ݴ´�.
        Close #iFileNo
        
        lvGroups.Refresh
    Else
        MessageBox "�̹� �����ϰų� �ùٸ��� �ʽ��ϴ�.", "����", Me, 16
    End If
End Sub

Private Sub cmdDelGroup_Click()
    On Error Resume Next
    Kill "C:\CALPLANS\CTGROUPS\" & lvGroups.List(lvGroups.ListIndex)
    
    lvGroups.Refresh
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
    Loaded = False
    lngOriginalTop = grpNotificationContainer.Top
    lngIncrement = (grpNotificationContainer.height - Frame9.height) / VScroll1.Max
    
    lngOriginalTop2 = grpRingtoneContainer.Top
    lngIncrement2 = (grpRingtoneContainer.height - Frame12.height) / VScroll2.Max
   
    ResetCount = 7
    'chkNoResize.Value = GetSetting("Calendar", "Options", "NoResize", "0")
    
    chkNoTimeCHeck.Value = GetSetting("Calendar", "Options", "NoTimeCheck", 0)
    
    chkTP.Value = GetSetting("Calendar", "Options", "TP", 0)
    chkAlwaysRm.Value = GetSetting("Calendar", "Options", "AlwaysRibbon", 0)
    
    If GetSetting("Calendar", "Options", "SST", True) = True Then
        radSelST.Value = True
    Else
        radCFQ.Value = True
    End If
    
    chkNoRibbon.Value = GetSetting("Calendar", "Options", "NoRibbon", 0)
    
    Dim ctrl2 As Control
    For Each ctrl2 In Me.Controls
        If ctrl2.Container.Name = Frame5.Name Then
            ctrl2.Enabled = False
        End If
    Next ctrl2
    
    
    On Error Resume Next
    cmbWSD.AddItem LoadLang("�Ͽ���", "Sunday", "Domingo")
    cmbWSD.AddItem LoadLang("������", "Monday", "Lunes")
    cmbWSD.AddItem LoadLang("ȭ����", "Tuesday", "Martes")
    cmbWSD.AddItem LoadLang("������", "Wednesday", "Miercoles")
    cmbWSD.AddItem LoadLang("�����", "Thursday", "Jueves")
    cmbWSD.AddItem LoadLang("�ݿ���", "Friday", "Viernes")
    cmbWSD.AddItem LoadLang("�����", "Saturday", "Sabado")
    
    cmbStartPage.AddItem LoadLang("����", "Plans", "Planes")
    cmbStartPage.AddItem LoadLang("�ּҷ�", "Contacts", "Contactos")
    cmbStartPage.AddItem LoadLang("�� ��", "Tasks", "Tareas")
    cmbStartPage.AddItem LoadLang("�ϰ�ǥ", "Schedule", "Calendario")
    cmbStartPage.AddItem LoadLang("�˶�", "Alarms", "Alarmas")
    
    cmbBGColor.AddItem LoadLang("�ý���: �������α׷� �۾�����", "System Scheme: Application Background")
    cmbBGColor.AddItem LoadLang("�ý���: ���� ǥ���", "System Scheme: Button Face")
    cmbBGColor.AddItem LoadLang("����", "Red", "Rojo")
    cmbBGColor.AddItem LoadLang("���", "Yellow", "Amarillo")
    cmbBGColor.AddItem LoadLang("�ʷ�", "Green", "Verde")
    cmbBGColor.AddItem LoadLang("����", "Cyan", "Cian")
    cmbBGColor.AddItem LoadLang("û��", "Dark Cyan", "Cian oscuro")
    cmbBGColor.AddItem LoadLang("�Ķ�", "Blue", "Azul")
    cmbBGColor.AddItem LoadLang("����", "Black", "Negro")
    
    cmbThemeSelect.AddItem LoadLang("Ǫ�� �ܿ�", "Winter")
    cmbThemeSelect.AddItem LoadLang("������ Ʈ��", "Orange Truck")
    cmbThemeSelect.AddItem LoadLang("���� �����", "Lemon Submarine")
    
    Me.Caption = LoadLang("ȯ�漳��", "Settings", "Ambientacion")
    Command1.Caption = LoadLang("Ȯ��", "OK", "Tienda")
    Command2.Caption = LoadLang("���", "Cancel", "Cancelar")
    cmdOptionHelp.Caption = LoadLang("����(&H)", "&Help", "Ayuda(&H)") & "..."
    cmdTheSet.Caption = LoadLang("�׸�(&T)", "&Theme", "&Tema") & "..."
    Frame8.Caption = LoadLang("���̾ƿ�", "Layout", "Diseno")
    chkTP.Caption = LoadLang("���������� �����(&O)", "Hide t&oday's plan list", "Ocultar la lista de planes de h&oy")
    chkNoRibbon.Caption = LoadLang("���� �޴� ��Ȱ��(&N)", "Disable ribbo&n menu", "Deshabilitar me&nu de ribbon")
    Frame1.Caption = LoadLang("�޷�", "Calendar", "Calendario")
    Frame16.Caption = LoadLang("�׸�", "Theme", "Tema")
    'Frame10.Caption = LoadLang("�׸�", "Theme", "Tema")
    
    Frame14.Caption = LoadLang("���", "Language", "Idioma")
    
    Label5.Caption = LoadLang("���� ����", "Start of week", "Dia de inicio") & ":"
    'Label13.Caption = LoadLang("�׸��� �����Ϸ��� ���� ���߸� �����ʽÿ�.", "To apply theme, click the button.", "Haga clic en el boton Siguiente para aplicar el tema.")
    
    SSTab1.TabCaption(0) = LoadLang("��Ÿ��", "Appearence", "Apariencia") 'Pantalla de visualizacion
    SSTab1.TabCaption(1) = LoadLang("����� ������", "User Data", "Datos del usuario")
    SSTab1.TabCaption(2) = LoadLang("�⺻ ����", "Defaults", "Predeterminados")
    SSTab1.TabCaption(3) = LoadLang("�Է°˻�", "Value Checking", "Comprobacion")
    SSTab1.TabCaption(4) = LoadLang("�з� �� �׷�", "Categories", "Categorias")
    'SSTab1.TabCaption(5) = LoadLang("�׸�", "Theme", "Tema")
    SSTab1.TabCaption(6) = LoadLang("����", "Security", "Contrasena")
    SSTab1.TabCaption(7) = LoadLang("�Ҹ�", "Sounds", "Sonido")
    
    Frame2.Caption = LoadLang("�� ������", "My data", "Mis datos")
    Label1.Caption = LoadLang("�� ����", "My Plans", "Mis planes") & ":"
    Label2.Caption = LoadLang("�� �ּҷ�", "My Contacts", "Mis contactos") & ":"
    Label3.Caption = LoadLang("�� �۾�", "My Tasks", "Mis tareas") & ":"
    
    cmdDelPlans.Caption = LoadLang("��� ����(&D)", "&Delete All", "Eliminar to&do")
    cmdDelContacts.Caption = LoadLang("��� ����(&E)", "D&elete All", "&Eliminar todo")
    cmdDelTasks.Caption = LoadLang("��� ����(&L)", "De&lete All", "E&liminar todo")
    
    Frame4.Caption = LoadLang("���� �� �۾�", "On startup", "En el arranque")
    Label7.Caption = LoadLang("���� ������", "Startup Page", "Pagina de inicio") & ":"
    radSelST.Caption = LoadLang("������ ����(&T)", "Selec&t a page", "Selecciona una pagina(&T)")
    radCFQ.Caption = LoadLang("������ ���� �������� ����(&Q)", "Resume where you &quited", "Reanudar donde dejo(&Q)")
    
    Frame11.Caption = LoadLang("�� ����")
    
    If LoadLang(1, 2, 3) <> 1 Then Frame11.Visible = False
    
    Frame6.Caption = LoadLang("�ð�", "Time", "Hora")
    chkNoTimeCHeck.Caption = LoadLang("���� �߰� �� �ð��� �ùٸ��� �˻� ����(&T)", "Do not check if the &time is invalid", "No verifique si el &tiempo no es valido")
    Label9.Caption = "[*] " & LoadLang("�� ������ �����ϸ� ���α׷��� �ùٷ� �۵����� ���� �� �ֽ��ϴ�.", "Changing settings in this page may cause internal errors", "Cambiar la configuracion en esta pagina puede causar errores internos")
    
    Label14.Caption = LoadLang("���� �з� ���:                 �׷� ���:", "Plan categories:               Contact groups:", "Categorias:                      Grupo:")
    cmdDelSelCate.Caption = LoadLang("���� ����(&S)", "Delete &selected", "Eliminar &seleccionado")
    cmdDelGroup.Caption = LoadLang("���� ����(&E)", "Delete s&elected", "&Eliminar seleccionado")
    
    cmdClearCates.Caption = LoadLang("�з� ��ü����", "Clear Categories", "Eliminar todo categorias")
    cmdClearGroups.Caption = LoadLang("�׷� ��ü����", "Clear Groups", "Eliminar todo grupos")
    
    Label8.Caption = LoadLang("�� ���� �з� �߰�", "New Category", "Nueva categoria") & ":"
    Label11.Caption = LoadLang("�� �׷� �߰�", "New Group", "Nueva grupo") & ":"
    
    cmdAddNewCate.Caption = LoadLang("�߰�(&A)", "&Add", "&Anadir")
    cmdAddNewGroup.Caption = LoadLang("�߰�(&D)", "A&dd", "Ana&dir")
    
    Frame7.Caption = LoadLang("���� �޴� ����", "Ribbon Menu background color", "Tema")
    'Label10.Caption = LoadLang("���� �Ǹ��� ����", "Background Color", "Color de fondo") & ":"
    
    Frame34.Caption = LoadLang("���� �˸���", "Notification Sound", "Sonido de notificacion")
    Frame12.Caption = LoadLang("�˶���", "Alarm Ringtone", "Tono de alarma")
    
    optNotificationSound(0).Caption = LoadLang("��- ��-", "Beep- Beep-")
    optNotificationSound(1).Caption = LoadLang("������-", "Bee-eep-")
    optNotificationSound(2).Caption = LoadLang("�ߺ�- �ߺ�-", "Beepbeep-")
    optNotificationSound(3).Caption = LoadLang("��- ��- ��-", "Beep- beep- beep-")
    optNotificationSound(4).Caption = LoadLang("�ߺ�- ��- �ߺ�-", "Beepeep- Beep- Beepeep-")
    optNotificationSound(5).Caption = LoadLang("�����", "Stair Tone", "Tono de escalera") & " 1"
    optNotificationSound(6).Caption = LoadLang("��- �ߺ�-", "Beep- Beepeep-")
    optNotificationSound(7).Caption = LoadLang("�����", "Stair Tone", "Tono de escalera") & " 2"
    optNotificationSound(8).Caption = LoadLang("�ߺ�- ���� 3", "Beep-beep- ��3")
    
    cmdPlayNS.Caption = LoadLang("���(&P)", "&Preview", "Vista &previa")
    cmdPlayRT.Caption = LoadLang("���(&R)", "P&review", "Vista p&revia")
    
    optRingtone(0).Caption = LoadLang("�⺻��", "Basic Tone", "Tono basico")
    optRingtone(1).Caption = LoadLang("�����", "Stair Tone", "Tono de escalera")
    optRingtone(2).Caption = LoadLang("�Ʊ���� �Ѹ�", "Dooly theme", "Tema del Dooly")
    
    cmbBGColor.ListIndex = GetSetting("Calendar", "Options", "BGColor", 0)
    cmbThemeSelect.ListIndex = GetSetting("Calendar", "Options", "Theme2", 0)
    
    cmbLanguage.AddItem "�ѱ���"
    cmbLanguage.AddItem "English"
    cmbLanguage.AddItem "Espanol"
    'cmbLanguage.AddItem "����"
    
    cmbLanguage.ListIndex = GetSetting("Calendar", "Options", "Language", 0)
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\CTGORIES"
    MkDir "C:\CALPLANS\CTGROUPS"
    
    lvCustomCates.Path = "C:\CALPLANS\CTGORIES"
    lvGroups.Path = "C:\CALPLANS\CTGROUPS"
    
    cmbStartPage.ListIndex = GetSetting("Calendar", "Options", "StartPage", 0)
    
    cmbWSD.ListIndex = GetSetting("Calendar", "Options", "WSD", 0)
    
    NSI = GetSetting("Calendar", "Options", "Notification", 0)
    RTI = GetSetting("Calendar", "Options", "Ringtone", 0)
    
    If GetSetting("Calendar", "Config", "EggEnabled", "0") = "1" Then
        optRingtone(2).Visible = True
    End If
    
    optNotificationSound.Item(NSI).Value = True
    optRingtone.Item(RTI).Value = True
    
    Loaded = True
    
    If GetSetting("Calendar", "Options", "Password", "") <> "" Then
        chkPasswordRequired.Value = 1
        
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If ctrl.Container.Name = Frame5.Name Then
                ctrl.Enabled = True
            End If
        Next ctrl
    End If
End Sub

Private Sub optNotificationSound_Click(Index As Integer)
    NSI = Index
End Sub

Private Sub optRingtone_Click(Index As Integer)
    RTI = Index
End Sub

Private Sub radCFQ_Click()
    cmbStartPage.Enabled = False
End Sub

Private Sub radSelST_Click()
    cmbStartPage.Enabled = True
End Sub

Private Sub txtAdvancedSetting_Change()
    On Error Resume Next
    If UCase(txtAdvancedSetting.Text) = "PASSWORD" Then
        txtAdvancedValue.Text = ""
    Else
        txtAdvancedValue.Text = GetSetting("Calendar", "Options", txtAdvancedSetting.Text)
    End If
End Sub

Private Sub VScroll1_Change()
    grpNotificationContainer.Top = lngOriginalTop - (VScroll1.Value * lngIncrement)
End Sub
