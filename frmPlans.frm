VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPlans 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 목록"
   ClientHeight    =   4020
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7170
   Icon            =   "frmPlans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdTody 
      Caption         =   "T"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      ToolTipText     =   "오늘의 날짜로 이동"
      Top             =   80
      Width           =   375
   End
   Begin VB.CommandButton cmdYest 
      Caption         =   "◀"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      ToolTipText     =   "어제로"
      Top             =   80
      Width           =   375
   End
   Begin VB.CommandButton cmdTomr 
      Caption         =   "▶"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      ToolTipText     =   "내일로"
      Top             =   80
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveTo 
      Caption         =   "→"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      ToolTipText     =   "선택한 날짜로 이동"
      Top             =   80
      Width           =   375
   End
   Begin VB.ComboBox txtDay 
      Height          =   300
      Left            =   3000
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox txtMonth 
      Height          =   300
      Left            =   1680
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtYear 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   120
      MaxLength       =   4
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "닫기(&C)"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.FileListBox lvPlanFiles 
      Height          =   450
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdViewPlan 
      Caption         =   "보기(&V)..."
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelBtn 
      Caption         =   "삭제(&D)"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddBtn 
      Caption         =   "추가(&C)..."
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin ComctlLib.ListView lstPlanList 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
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
   Begin VB.Label Label1 
      Caption         =   "년                월                   일"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   165
      Width           =   3135
   End
End
Attribute VB_Name = "frmPlans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public CurrentDate As Date
Dim Year As Integer
Dim Month As Integer
Dim Day As Integer
Dim Plans As String
Dim Plan As Integer

Dim Title As String
Dim Time As String
Dim Category As String
Dim Imprty As String
Dim PlanData As String
Dim PlanItem As ListItem

Private Sub cmdAddBtn_Click()
    frmAddPlan.CurrentDate = CurrentDate
    frmAddPlan.Show vbModal, Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelBtn_Click()
    On Error Resume Next
    If Confirm("'" & lstPlanList.SelectedItem.Text & "' 일정을 삭제하시겠습니까?", "일정 삭제", Me) Then
        Kill "C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\" & lstPlanList.SelectedItem.Text
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Cate"
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Time"
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Location"
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Cont"
        DeleteSetting "Calendar", "NotifiedPlans\" & Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text
    End If
    
    frmMain.lvTodaysPlan.Refresh
    frmMain.lvTodaysPlans.Refresh
    frmMain.lvTmrPlans.Refresh
    
    LoadPlans
End Sub

Sub LoadPlans()
    On Error Resume Next
    lstPlanList.ListItems.Clear
    lvPlanFiles.Refresh
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\" & Year
    MkDir "C:\CALPLANS\" & Year & "\" & Month
    MkDir "C:\CALPLANS\" & Year & "\" & Month & "\" & Day
    
    lvPlanFiles.Path = "C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\"

    For Plan = 0 To lvPlanFiles.ListCount - 1
        'PlanData = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFIles.List(Plan), "(지정되지 않음)")

        Title = lvPlanFiles.List(Plan)
        Time = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFiles.List(Plan) & "Time", "(지정되지 않음)")
        Category = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFiles.List(Plan) & "Cate", "(지정되지 않음)")
        Imprty = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFiles.List(Plan) & "Impt", "1")

        lstPlanList.ListItems.Add , , Title
        lstPlanList.ListItems(Plan + 1).SubItems(1) = Left$(Time, 2) & ":" & Right$(Time, 2)
        lstPlanList.ListItems(Plan + 1).SubItems(2) = Category
        lstPlanList.ListItems(Plan + 1).SubItems(3) = Imprty
    Next Plan
End Sub

Private Sub cmdTody_Click()
    CurrentDate = Format(Now, "YYYY-M-D")
    Form_Load
End Sub

Private Sub cmdTomr_Click()
    CurrentDate = DateAdd("d", 1, CurrentDate)
    Form_Load
End Sub

Private Sub cmdViewPlan_Click()
    On Error GoTo exitsub
    frmPlanView.CurrentDate = CurrentDate
    frmPlanView.Caption = lstPlanList.SelectedItem.SubItems(2) & " 일정 - " & lstPlanList.SelectedItem.Text
    frmPlanView.Category = lstPlanList.SelectedItem.SubItems(2)
    frmPlanView.Title = lstPlanList.SelectedItem.Text
    frmPlanView.lblTimeHrs.Caption = Split(lstPlanList.SelectedItem.SubItems(1), ":")(0) & "시"
    frmPlanView.lblTimeMin.Caption = Split(lstPlanList.SelectedItem.SubItems(1), ":")(1) & "분"
    frmPlanView.lblLocation.Text = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Location", "알 수 없음")
    frmPlanView.txtContent.Text = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Cont", "자세한 내용 없음")
    frmPlanView.txtImprty.Text = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Impt", "1")
    frmPlanView.txtParts.Text = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Part", "")
    frmPlanView.Show vbModal, Me
    
exitsub:
    Exit Sub
End Sub

Private Sub cmdMoveTo_Click()
    On Error Resume Next
    If IsNumeric(txtYear.Text) = False Or IsNumeric(txtYear.Text) = False Or IsNumeric(txtYear.Text) = False Then
        MessageBox "날짜의 값은 숫자이여야 합니다.", "오류", Me, 16
        Exit Sub
    ElseIf txtYear.Text < 1 Or txtYear.Text > 9999 Then
        MessageBox "해(연도)의 값은 1 이상 9999 이하이여야 합니다. 88, 05 등의 값은 1988, 2005으로 치환됩니다.", "오류", Me, 16
        Exit Sub
    ElseIf txtMonth.Text < 1 Or txtMonth.Text > 12 Then
        MessageBox "달의 값은 1 이상 12 이하이여야 합니다.", "오류", Me, 16
        Exit Sub
    ElseIf (txtDay.Text < 1 Or txtDay.Text > 31) Then
        MessageBox "일의 값은 달의 값이 7 이하이면서 홀수이거나 8 이상이면서 짝수일 때에는 1 이상 31 이하, 달이 2일 때는 1 이상 29 이하, 그 외에는 1 이상 30 이하이여야 합니다.", "오류", Me, 16
        Exit Sub
    End If
    
    CurrentDate = txtYear.Text & "-" & txtMonth.Text & "-" & txtDay.Text
    Form_Load
End Sub

Private Sub cmdYest_Click()
    CurrentDate = DateAdd("d", -1, CurrentDate)
    Form_Load
End Sub

Private Sub Form_Load()
    Year = Split(CurrentDate, "-")(0)
    Month = Split(CurrentDate, "-")(1)
    Day = Split(CurrentDate, "-")(2)
    Me.Caption = Year & "년 " & Month & "월 " & Day & "일의 일정 목록"
    
    lstPlanList.ColumnHeaders.Clear
    
    txtMonth.Clear
    txtDay.Clear
    
    txtYear.Text = Year
    txtMonth.Text = Month
    txtDay.Text = Day
    
    On Error Resume Next
    
    Dim i As Integer
    For i = 1 To 12
        txtMonth.AddItem CStr(i)
    Next i
    For i = 1 To 31
        txtDay.AddItem CStr(i)
    Next i
    
    lstPlanList.ColumnHeaders.Add , , "제목", 2000
    lstPlanList.ColumnHeaders.Add , , "시간", 350
    lstPlanList.ColumnHeaders.Add , , "분류", 850
    lstPlanList.ColumnHeaders.Add , , "중요도", 450
    
    LoadPlans
    
    lstPlanList.SortKey = 1
End Sub

Private Sub lstPlanList_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lstPlanList.SortKey = ColumnHeader.SubItemIndex
    
    If lstPlanList.SortOrder = lvwAscending Then
        lstPlanList.SortOrder = lvwDescending
    Else
        lstPlanList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstPlanList_DblClick()
    On Error Resume Next
    cmdViewPlan_Click
End Sub

Private Sub lstPlanList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmdDelBtn_Click
    End If
End Sub
