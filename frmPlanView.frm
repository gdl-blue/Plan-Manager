VERSION 5.00
Begin VB.Form frmPlanView 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "내 일정"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "frmPlanView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "frmPlanView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Title As String

Private Sub Form_Load()
    Me.Caption = "내 일정 - " & Title
End Sub
