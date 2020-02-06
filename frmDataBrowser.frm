VERSION 5.00
Begin VB.Form frmDataBrowser 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "모든 일정 색인"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   Icon            =   "frmDataBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.FileListBox File1 
      Height          =   2970
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   3030
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmDataBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
    On Error Resume Next
    If Left$(Dir1.Path, 11) <> "C:\CALPLANS" Then Dir1.Path = "C:\CALPLANS"
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MkDir "C:\CALPLANS"
    MkDir "C:\CALPLANS\TASKS"
    MkDir "C:\CALPLANS\CONTACTS"
    
    Dir1.Path = "C:\CALPLANS"
End Sub
