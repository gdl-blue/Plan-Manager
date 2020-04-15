VERSION 5.00
Begin VB.Form frmCheckDeactivatePassword 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "비밀번호 비활성화"
   ClientHeight    =   1470
   ClientLeft      =   2760
   ClientTop       =   3990
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "frmCheckDeactivatePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
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
   Begin VB.Label Label1 
      Caption         =   "현재 비밀번호:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmCheckDeactivatePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    frmOptions.chkPasswordRequired.Value = 1
    Unload Me
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
