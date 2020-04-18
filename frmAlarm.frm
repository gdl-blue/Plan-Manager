VERSION 5.00
Begin VB.Form frmAlarm 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "알람"
   ClientHeight    =   3150
   ClientLeft      =   2760
   ClientTop       =   3960
   ClientWidth     =   6030
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAlarmMemo 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Timer timTimeChecker 
      Interval        =   500
      Left            =   360
      Top             =   1320
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "CancelButton"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblTime 
      Caption         =   "현재 시각: 00:00:00"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "알람 이름"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'https://stackoverflow.com/questions/17651725/vb6-how-to-make-a-floating-window-top-most

  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2

  Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
        
Public yy As Integer
Public mm As Integer
Public dd As Integer

  Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
     As Long

     If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
           0, FLAGS)
     Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
           0, 0, FLAGS)
        SetTopMostWindow = False
     End If
  End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblTime.Caption = LoadLang("현재 시각", "TIme") & ": " & Format(Now, "hh:mm:ss")
    Me.Caption = LoadLang("알람", "Alarm")
    CancelButton.Caption = LoadLang("음소거(&M)", "&Mute")
    SetTopMostWindow Me.hwnd, True
    PlayRingtone
End Sub

Private Sub timTimeChecker_Timer()
    lblTime.Caption = LoadLang("현재 시각", "TIme") & ": " & Format(Now, "hh:mm:ss")
End Sub
