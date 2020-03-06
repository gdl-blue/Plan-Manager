VERSION 5.00
Begin VB.Form frmAddPlan 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 추가"
   ClientHeight    =   3510
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmAddPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.FileListBox lvCateFiles 
      Height          =   270
      Left            =   5160
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtContent 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   13
      Top             =   2400
      Width           =   4335
   End
   Begin VB.ComboBox txtCategory 
      Height          =   300
      ItemData        =   "frmAddPlan.frx":0442
      Left            =   120
      List            =   "frmAddPlan.frx":044F
      TabIndex        =   11
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtLocation 
      Height          =   270
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtTimeMin 
      Height          =   270
      Left            =   720
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
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "내용:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "분류:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "위치:"
      Height          =   255
      Left            =   1320
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

Private Sub Form_Load()
    Year = Split(CurrentDate, "-")(0)
    Month = Split(CurrentDate, "-")(1)
    Day = Split(CurrentDate, "-")(2)
    Me.Caption = "일정 추가 - " & Year & "년 " & Month & "월 " & Day & "일"
    
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
End Sub

Private Sub OKButton_Click()
    '입력값을 검사한다.
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
    If txtCategory.Text = "" Then
        txtCategory.Text = "(지정되지 않음)"
    End If
    
    '일정을 추가하기 전에 해당 제목의 일정이 존재하는지 확인한다.
    If FileExists("C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\" & txtTitle.Text) = True Then
        MessageBox "해당 제목의 일정이 이미 존재합니다...", "처리 중 오류", Me, 16
    End If
    
    '해당 일정이 존재함을 알리는 파일을 만든다.
    'https://stackoverflow.com/questions/21108664/how-to-create-txt-file
    Dim iFileNo As Integer
    iFileNo = FreeFile
    '파일을 연다.
    Open "C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\" & txtTitle.Text For Output As #iFileNo
    
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
    
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Time", txtTime
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Location", txtLocation.Text
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Cate", txtCategory.Text
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Cont", txtContent.Text
    
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
