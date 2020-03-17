VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmEditPlan 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 수정"
   ClientHeight    =   4185
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmEditPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtParts 
      Height          =   270
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   3495
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   600
      TabIndex        =   17
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtImprty"
      BuddyDispid     =   196625
      OrigLeft        =   600
      OrigTop         =   1560
      OrigRight       =   855
      OrigBottom      =   1815
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtImprty 
      Height          =   270
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   480
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "저장(&S)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtTimeHrs 
      Height          =   270
      Left            =   120
      MaxLength       =   2
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtTimeMin 
      Height          =   270
      Left            =   720
      MaxLength       =   2
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtLocation 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.ComboBox txtCategory 
      Height          =   300
      ItemData        =   "frmEditPlan.frx":0442
      Left            =   120
      List            =   "frmEditPlan.frx":044F
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtContent 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.FileListBox lvCateFiles 
      Height          =   270
      Left            =   5160
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "참여자:"
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "중요도:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "제목:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "시간:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   " :"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "위치:"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "분류:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "내용:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "frmEditPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim txtTime As String
Dim Category As Integer
Public Title As String
Public CurrentDate As Date
Public Year As Integer
Public Month As Integer
Public Day As Integer
Dim iFileNo As Integer

Private Sub CancelButton_Click()
    If Confirm("일정 수정를 취소하시겠습니까? 임시 저장되지 않습니다.", "일정 추가", Me) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "일정 수정 - " & Title
    
    On Error Resume Next
    MkDir "C:\CALPLANS\CTGORIES"
    
    txtCategory.Clear
    txtCategory.AddItem "업무"
    txtCategory.AddItem "여가생활"
    txtCategory.AddItem "약속"
    txtCategory.AddItem "취미"
    
    lvCateFiles.Path = "C:\CALPLANS\CTGORIES"
    
    For Category = 0 To lvCateFiles.ListCount - 1
        txtCategory.AddItem lvCateFiles.List(Category)
    Next Category
End Sub

Private Sub OKButton_Click()
    '입력값을 검사한다.
    If IsNumeric(txtTimeHrs.Text) = False Or IsNumeric(txtTimeMin.Text) = False Then
        MessageBox "시간의 값이 올바르지 않습니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    If GetSetting("Calendar", "Options", "NoTimeCheck", 0) = 0 Then
        If txtTimeHrs.Text > 23 Or txtTimeMin.Text > 59 Or txtTimeHrs.Text < 0 Or txtTimeMin.Text < 0 Then
            MessageBox "시간에서 시는 0부터 24, 분은 0부터 59까지의 정수이여야 합니다.", "입력 값 오류", Me, 16
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
    
    If IsNumeric(txtImprty.Text) = False Or txtImprty.Text < 1 Or txtImprty.Text > 10 Then
        MessageBox "중요도는 1(낮음)부터 10(높음)까지여야 합니다.", "입력 값 오류", Me, 16
        Exit Sub
    End If
    
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
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Part", txtParts.Text
    SaveSetting "Calendar", Year & "\" & Month & "\" & Day, txtTitle.Text & "Impt", txtImprty.Text
    
    frmPlans.LoadPlans
    
    '분류를 추가한다.
    If FileExists("C:\CALPLANS\CTGORIES\" & txtCategory.Text) = False Then
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
    End If
    
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

