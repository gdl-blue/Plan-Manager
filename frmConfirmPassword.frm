VERSION 5.00
Begin VB.Form frmConfirmPassword 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환영합니다!"
   ClientHeight    =   1470
   ClientLeft      =   2760
   ClientTop       =   3990
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "frmConfirmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "종료(&X)"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConfirmPassword.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "암호를 입력해주세요."
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmConfirmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, _
ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Function DisableCloseButton(frm As Form) As Boolean
Dim lHndSysMenu As Long
Dim lAns1 As Long, lAns2 As Long
lHndSysMenu = GetSystemMenu(frm.hWnd, 0)
lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function

Private Sub CancelButton_Click()
    End
End Sub

Private Sub Form_Load()
    DisableCloseButton Me
End Sub

Private Sub OKButton_Click()
    If GetSetting("Calendar", "Options", "Password", "") = Text1.Text Then
        Unload Me
    Else
        MsgBox "암호가 올바르지 않습니다.", 16, "오류"
    End If
End Sub
