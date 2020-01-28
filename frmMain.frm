VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "일정표"
   ClientHeight    =   5205
   ClientLeft      =   150
   ClientTop       =   825
   ClientWidth     =   7350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows 기본값
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      TabCaption(0)   =   "일정"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MonthView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "주소록"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "할 일"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4170
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   7355
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         MonthColumns    =   3
         MonthRows       =   2
         ShowToday       =   0   'False
         StartOfWeek     =   20185089
         CurrentDate     =   43858
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4935
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7779
            Text            =   "날짜를 누르십시오."
            TextSave        =   "날짜를 누르십시오."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2020-01-28"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "오후 4:33"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "새 파일(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "열기(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "닫기(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "저장(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "다른 이름으로 저장(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "모두 저장(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "속성(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "페이지 설정(&U)"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "인쇄 미리보기(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "인쇄(&P)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "보내기(&D)..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "편집(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "실행 취소(&U)"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "잘라내기(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "복사(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "붙여넣기(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "선택하여 붙여넣기(&S)..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "보기(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "도구 모음(&T)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "상태 표시줄(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "새로 고침(&R)"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "옵션(&O)..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "웹 브라우저(&W)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "도구(&T)"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "옵션(&O)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "목차(&C)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "찾기(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "정보(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
    Me.Left = GetSetting("Calendar", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("Calendar", "Settings", "MainTop", 1000)
    Me.Width = 7440
    Me.Height = 5970
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
        MsgBox "도움말 목차를 표시할 수 없습니다. 이 프로젝트와 연관된 도움말이 없습니다.", vbInformation, Me.Caption
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
        MsgBox "도움말 목차를 표시할 수 없습니다. 이 프로젝트와 연관된 도움말이 없습니다.", vbInformation, Me.Caption
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
    MsgBox "'mnuViewOptions_Click' 코드를 추가하십시오."
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
    '폼을 언로드합니다.
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
    MsgBox "'mnuFileProperties_Click' 코드를 추가하십시오."
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
    MsgBox "'mnuFileSave_Click' 코드를 추가하십시오."
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

