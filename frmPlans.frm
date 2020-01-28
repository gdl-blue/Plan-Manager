VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPlans 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 목록"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   Icon            =   "frmPlans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox lvPlanFIles 
      Height          =   450
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdViewPlan 
      Caption         =   "보기(&V)..."
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelBtn 
      Caption         =   "삭제(&D)..."
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddBtn 
      Caption         =   "추가(&C)..."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin ComctlLib.ListView lstPlanList 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
Dim PlanData As String
Dim PlanItem As ListItem

Private Sub cmdAddBtn_Click()
    frmAddPlan.CurrentDate = CurrentDate
    frmAddPlan.Show vbModal, Me
End Sub

Private Sub cmdDelBtn_Click()
    On Error Resume Next
    If MsgBox("'" & lstPlanList.SelectedItem.Text & "' 일정을 삭제하시겠습니까?", vbQuestion + vbOKCancel, "일정 삭제") = vbOK Then
        Kill "C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\" & lstPlanList.SelectedItem.Text
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Cate"
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Time"
        DeleteSetting "Calendar", Year & "\" & Month & "\" & Day, lstPlanList.SelectedItem.Text & "Location"
    End If
    
    LoadPlans
End Sub

Sub LoadPlans()
    On Error Resume Next
    lstPlanList.ListItems.Clear
    lvPlanFIles.Refresh
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\" & Year
    MkDir "C:\CALPLANS\" & Year & "\" & Month
    MkDir "C:\CALPLANS\" & Year & "\" & Month & "\" & Day
    
    lvPlanFIles.Path = "C:\CALPLANS\" & Year & "\" & Month & "\" & Day & "\"
    
    For Plan = 0 To lvPlanFIles.ListCount - 1
        'PlanData = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFIles.List(Plan), "(지정되지 않음)")
    
        Title = lvPlanFIles.List(Plan)
        Time = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFIles.List(Plan) & "Time", "(지정되지 않음)")
        Category = GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvPlanFIles.List(Plan) & "Cate", "(지정되지 않음)")
        
        lstPlanList.ListItems.Add , , Title
        lstPlanList.ListItems(Plan + 1).SubItems(1) = Left$(Time, 2) & ":" & Right$(Time, 2)
        lstPlanList.ListItems(Plan + 1).SubItems(2) = Category
    Next Plan
End Sub

Private Sub Form_Load()
    Year = Split(CurrentDate, "-")(0)
    Month = Split(CurrentDate, "-")(1)
    Day = Split(CurrentDate, "-")(2)
    Me.Caption = Year & "년 " & Month & "월 " & Day & "일의 일정 목록"
    
    lstPlanList.ColumnHeaders.Add , , "제목", 2000
    lstPlanList.ColumnHeaders.Add , , "시간", 350
    lstPlanList.ColumnHeaders.Add , , "분류", 850
    
    LoadPlans
End Sub
