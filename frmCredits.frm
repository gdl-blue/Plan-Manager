VERSION 5.00
Begin VB.Form frmCredits 
   Caption         =   "frmCredits"
   ClientHeight    =   3720
   ClientLeft      =   30
   ClientTop       =   480
   ClientWidth     =   5640
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   4170
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCredits.frx":0442
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
    
    Me.Caption = LoadLang("코드 출처", "Credits")
End Sub

Private Sub Form_Resize()
    Text1.Width = Me.Width
    Text1.Height = Me.Height
End Sub
