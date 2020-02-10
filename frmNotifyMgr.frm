VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotifyMgr 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일정 알리미"
   ClientHeight    =   90
   ClientLeft      =   24585
   ClientTop       =   14610
   ClientWidth     =   2220
   ControlBox      =   0   'False
   Icon            =   "frmNotifyMgr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   2220
   StartUpPosition =   2  '화면 가운데
   Visible         =   0   'False
   WindowState     =   1  '최소화
   Begin VB.FileListBox lvTodaysPlans 
      Height          =   810
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2220
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   120586241
      CurrentDate     =   43858
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1200
      Top             =   960
   End
End
Attribute VB_Name = "frmNotifyMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10초에 한 번씩 일정을 확인한다.
Dim Year As Integer
Dim Month As Integer
Dim Day As Integer
Dim Plan As Integer
Dim ttt As Integer

Private Sub Form_Activate()
    Me.WindowState = 1
End Sub

Private Sub Form_GotFocus()
    Me.WindowState = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    '오늘의 일정을 찾는다.
    On Error Resume Next
    
    Year = Split(MonthView1.SelStart, "-")(0)
    Month = Split(MonthView1.SelStart, "-")(1)
    Day = Split(MonthView1.SelStart, "-")(2)
    
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\" & Year
    MkDir "C:\CALPLANS\" & Year & "\" & Month
    MkDir "C:\CALPLANS\" & Year & "\" & Month & "\" & Day
    
    lvTodaysPlans.Path = "C:\CALPLANS\" & Year & "\" & Month & "\" & Day
    lvTodaysPlans.Refresh
    
    For Plan = 0 To lvTodaysPlans.ListCount - 1
        ttt = CInt(Split(GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(0) & Split(GetSetting("Calendar", Year & "\" & Month & "\" & Day, lvTodaysPlans.List(Plan) & "Time", "00:00"), ":")(1)) - CInt(Format(Now, "hhmm"))
        '현재시각과 일정시각과의 차이가 10분 미만이면 알림을 띄운다.
        If ttt < 10 And ttt >= -1 Then
            '띄운 적이 없으면 알림
            If GetSetting("Calendar", "NotifiedPlans\" & Year & "\" & Month & "\" & Day, lvTodaysPlans.List(Plan), "abc") = "abc" Then
                MsgBox lvTodaysPlans.List(Plan) & " 일정 시작까지 10분보다 적게 남았습니다. 준비하십시오. 이 알림은 다시 표시되지 않습니다.", vbInformation, "일정관리자"
                SaveSetting "Calendar", "NotifiedPlans\" & Year & "\" & Month & "\" & Day, lvTodaysPlans.List(Plan), ""
            End If
        End If
    Next Plan
End Sub
