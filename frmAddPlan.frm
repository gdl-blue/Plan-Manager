VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAddPlan 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 추가"
   ClientHeight    =   4410
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7095
   Icon            =   "frmAddPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtParticipants 
      Height          =   270
      Left            =   5400
      TabIndex        =   28
      Top             =   80
      Width           =   1575
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   4440
      TabIndex        =   26
      Top             =   75
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtImprt"
      BuddyDispid     =   196610
      OrigLeft        =   5280
      OrigTop         =   80
      OrigRight       =   5535
      OrigBottom      =   335
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtImprt 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   4080
      TabIndex        =   25
      Text            =   "1"
      Top             =   80
      Width           =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "이 일정이 있는 날짜"
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
      Begin VB.CommandButton cmdRPT 
         Caption         =   "되풀이"
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         ToolTipText     =   "이번달이 끝날때까지 일정 매주 되풀이"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdDP 
         Caption         =   "-"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   2400
         Width           =   255
      End
      Begin VB.ListBox lvPlanCP 
         Height          =   2040
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdCP 
         Caption         =   "+"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtCD 
         Height          =   270
         Left            =   1200
         TabIndex        =   19
         Top             =   2460
         Width           =   375
      End
      Begin VB.TextBox txtCM 
         Height          =   270
         Left            =   720
         TabIndex        =   18
         Top             =   2460
         Width           =   375
      End
      Begin VB.TextBox txtCY 
         Height          =   270
         Left            =   120
         MaxLength       =   4
         TabIndex        =   17
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "  -       -"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   2520
         Width           =   855
      End
   End
   Begin VB.CheckBox chkRepeat 
      Caption         =   "이번달이 끝날때까지 일정 매주 되풀이"
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.FileListBox lvCateFiles 
      Height          =   270
      Left            =   7920
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtContent 
      Height          =   2415
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   13
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ComboBox txtCategory 
      Height          =   300
      ItemData        =   "frmAddPlan.frx":0442
      Left            =   3480
      List            =   "frmAddPlan.frx":044F
      TabIndex        =   11
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtLocation 
      Height          =   270
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtTimeMin 
      Height          =   270
      Left            =   630
      MaxLength       =   2
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtTimeHrs 
      Height          =   270
      Left            =   120
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "24시 형식으로 입력합니다."
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtTitle 
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "추가(&A)"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "참여자:"
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "중요도:"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "내용:"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "분류:"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "위치:"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   " :"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "시간:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "24시 형식으로 입력합니다."
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "제목:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public CurrentDate As Date
Dim Year As Integer
Dim Month As Integer
Dim Day As Integer
Dim txtTime As String
Dim Category As Integer

Private Sub CancelButton_Click()
    If Confirm("일정 추가를 취소하시겠습니까? 임시 저장되지 않습니다.", "일정 추가", Me) Then
        Unload Me
    End If
End Sub

Private Sub cmdCP_Click()
    On Error Resume Next
    If IsNumeric(txtCY.Text) = False Or IsNumeric(txtCM.Text) = False Or IsNumeric(txtCD.Text) = False Then
        MessageBox "해, 달, 일의 값이 틀립니다. 숫자로 다시 입력하십시오.", "오류", Me, 16
        Exit Sub
    End If
    
    Dim cy As Integer
    Dim CM As Integer
    Dim CD As Integer
    
    cy = CStr(CInt(txtCY.Text))
    CM = CStr(CInt(txtCM.Text))
    CD = CStr(CInt(txtCD.Text))
    
    MkDir "C:\CALPLANS\" & cy & "\" & CM & "\" & CD
    
    lvPlanCP.AddItem cy & "-" & CM & "-" & CD
End Sub

Private Sub cmdDP_Click()
    On Error Resume Next
    If Confirm("추가할 일정을 선택한 날짜에 추가하지 않으시겠습니까?", "경고", Me) Then
        lvPlanCP.RemoveItem lvPlanCP.ListIndex
    End If
End Sub

Private Sub cmdRPT_Click()
    On Error Resume Next
    Dim i As Integer
    Dim DT As String
    DT = CurrentDate
    
    For i = 1 To 5
        DT = DateAdd("d", 7, DT)
        If Split(DT, "-")(1) <> Split(CurrentDate, "-")(1) Then
            Exit Sub
        End If
        lvPlanCP.AddItem CStr(CInt(Split(DT, "-")(0))) & "-" & CStr(CInt(Split(DT, "-")(1))) & "-" & CStr(CInt(Split(DT, "-")(2)))
    Next i
End Sub

'Private Sub cmdTody_Click()
'    CurrentDate = Format(Now, "YYYY-M-D")
'    Form_Load
'End Sub
'
'Private Sub cmdTomr_Click()
'    CurrentDate = DateAdd("d", 1, CurrentDate)
'    Form_Load
'End Sub
'
'Private Sub cmdYest_Click()
'    CurrentDate = DateAdd("d", -1, CurrentDate)
'    Form_Load
'End Sub

Private Sub Form_Load()
    Year = Split(CurrentDate, "-")(0)
    Month = Split(CurrentDate, "-")(1)
    Day = Split(CurrentDate, "-")(2)
    Me.Caption = "일정 추가 - " & Year & "년 " & Month & "월 " & Day & "일"
    
    txtCY.Text = Year
    txtCM.Text = Month
    txtCD.Text = Day
    
    On Error Resume Next
    MkDir "C:\CALPLANS\CTGORIES"
    
    lvCateFiles.Path = "C:\CALPLANS\CTGORIES"
    
    txtCategory.Clear
    txtCategory.AddItem "업무"
    txtCategory.AddItem "여가생활"
    txtCategory.AddItem "약속"
    txtCategory.AddItem "취미"
    
    For Category = 0 To lvCateFiles.ListCount - 1
        txtCategory.AddItem lvCateFiles.List(Category)
    Next Category
    
    lvPlanCP.AddItem CStr(CInt(Year)) & "-" & CStr(CInt(Month)) & "-" & CStr(CInt(Day))
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    If lvPlanCP.ListCount < 1 Then
        MessageBox "일정을 추가할 날짜가 없습니다.", "오류", Me, 16
        Exit Sub
    End If
    '입력값을 검사한다.
    If Mid$(txtTimeMin.Text, 1, 1) = "0" Then
        txtTimeMin.Text = Mid$(txtTimeMin.Text, 2, 1)
    End If
    If InStr(1, txtTitle.Text, "?") > 0 Or InStr(1, txtTitle.Text, "\") > 0 Or InStr(1, txtTitle.Text, "|") > 0 Or InStr(1, txtTitle.Text, ".") > 0 Or InStr(1, txtTitle.Text, "/") > 0 Or InStr(1, txtTitle.Text, "*") > 0 Or InStr(1, txtTitle.Text, ":") > 0 Or InStr(1, txtTitle.Text, ChrW$(34)) > 0 Then
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
    If txtTitle.Text = "" Then
        MessageBox "제목의 값은 필수입니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    If IsNumeric(txtImprt.Text) = False Or txtImprt.Text < 1 Or txtImprt.Text > 10 Then
        MessageBox "중요도는 1(낮음)부터 10(높음)까지여야 합니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    If txtCategory.Text = "" Then
        txtCategory.Text = "(지정되지 않음)"
    End If
    
    Dim dd As Integer
    Dim DCY As String
    Dim DCM As String
    Dim DCD As String
    
    For dd = 0 To lvPlanCP.ListCount - 1
        DCY = CStr(CInt(Split(lvPlanCP.List(dd), "-")(0)))
        DCM = CStr(CInt(Split(lvPlanCP.List(dd), "-")(1)))
        DCD = CStr(CInt(Split(lvPlanCP.List(dd), "-")(2)))
        
        '일정을 추가하기 전에 해당 제목의 일정이 존재하는지 확인한다.
        If FileExists("C:\CALPLANS\" & DCY & "\" & DCM & "\" & DCD & "\" & txtTitle.Text) = True Then
            MessageBox "해당 제목의 일정이 일부 날짜에 이미 존재합니다.", "처리 중 오류", Me, 16
        End If
        
        '해당 일정이 존재함을 알리는 파일을 만든다.
        'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
        Dim iFileNo As Integer
        iFileNo = FreeFile
        '파일을 연다.
        Open "C:\CALPLANS\" & DCY & "\" & DCM & "\" & DCD & "\" & txtTitle.Text For Output As #iFileNo
        
        '파일의 내용은 보지 않으므로 빈 칸으로...
        Print #iFileNo, ""
        
        '파일을 닫는다.
        Close #iFileNo
        
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
        
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Time", txtTime
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Location", txtLocation.Text
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Cate", txtCategory.Text
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Cont", txtContent.Text
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Part", txtParticipants.Text
        SaveSetting "Calendar", DCY & "\" & DCM & "\" & DCD, txtTitle.Text & "Impt", txtImprt.Text
        
        frmPlans.LoadPlans
        
        '분류를 추가한다.
        
        If txtCategory.Text <> "업무" And txtCategory.Text <> "여가생활" And txtCategory.Text <> "약속" And txtCategory.Text <> "취미" And txtCategory.Text <> "(지정되지 않음)" Then
            'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
            iFileNo = FreeFile
            '파일을 연다.
            
            Open "C:\CALPLANS\CTGORIES\" & txtCategory.Text For Output As #iFileNo
            
            '파일의 내용은 보지 않으므로 빈 칸으로...
            Print #iFileNo, ""
            
            '파일을 닫는다.
            Close #iFileNo
        End If
    Next dd
    
    frmMain.lvTodaysPlan.Refresh
    frmMain.lvTodaysPlans.Refresh
    frmMain.lvTmrPlans.Refresh
    
    Unload Me
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

Private Sub txtTimeMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
        If txtTimeMin.Text = "" Then txtTimeHrs.SetFocus
    End If
End Sub
