Attribute VB_Name = "ringtones"
Private Declare Function 소리 Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, Optional ByVal dwDuration As Long = 250) As Long

Sub PlayNotification(Optional ByVal Number As Integer = -1)
    Dim playn As Integer
    playn = Number
    
    If Number = -1 Then playn = CInt(GetSetting("Calendar", "Options", "Notification", 0))

    Select Case playn
        Case 0
            Beep 750, 300
            
            Pause 0.1
            
            Beep 750, 300
        Case 1
            Beep 950, 700
        Case 2
            Beep 850, 250
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
        Case 3
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
        Case 4
            Beep 850, 250
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
        Case 5
            Beep 523, 250
            Beep 659, 250
            Beep 782, 250
        Case 6
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
        Case 7
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
            Beep 850, 250
        Case 8
            Beep 850, 250
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
            
            Pause 0.1
            
            Beep 850, 250
            Beep 850, 250
    End Select
End Sub

Sub PlayRingtone(Optional ByVal Number As Integer = -1)
    Dim 옥 As Long
    
    Dim 장 As Long
    Dim 단 As Long
    Dim 중 As Long
    
    옥 = 2
    
    장 = 500
    단 = 100
    중 = 250
    
    Const 도 As Long = 522
    Const 레 As Long = 586
    Const 미 As Long = 658
    Const 파 As Long = 698
    Const 솔 As Long = 782
    Const 라 As Long = 880
    Const 시 As Long = 986
    
    Dim playn As Integer
    playn = Number
    
    If Number = -1 Then playn = CInt(GetSetting("Calendar", "Options", "Ringtone", 0))
    
    Select Case playn
        Case 0
            Beep 750, 300
            Beep 850, 300
            Beep 750, 300
            Beep 650, 300
            
            Beep 750, 300
            Beep 750, 300
            Beep 750, 300
            
            Beep 850, 300
            Beep 850, 300
            Beep 850, 300
            
            Beep 750, 300
            Beep 750, 300
            Beep 750, 300
            
            Beep 650, 300
            Beep 550, 300
            Beep 450, 300
            
        Case 1
            Beep 350, 300
            Beep 450, 300
            Beep 550, 300
            Beep 650, 300
            Beep 750, 300
            Beep 850, 300
            Beep 950, 300
            Beep 1050, 300
            
            Beep 950, 300
            Beep 850, 300
            Beep 750, 300
            Beep 650, 300
            Beep 550, 300
            Beep 450, 300
            Beep 350, 300
            
            Beep 450, 300
            Beep 550, 300
            Beep 650, 300
            Beep 750, 300
            Beep 850, 300
            Beep 950, 300
            Beep 1050, 300
            
        Case 2
            '베이직 언어 Play 함수 있었으면...

            소리 도, 중 + 40
            소리 미
            소리 도
            소리 시 / 옥, 장 + 단 + 단 + 단
            
            소리 시 / 옥, 중 + 40
            소리 미
            소리 시 / 옥
            소리 라 / 옥, 장 + 단 + 단 + 단

            소리 도, 중 + 40
            소리 미
            소리 도
            소리 시 / 옥, 장 + 단 + 단 + 단
            
            소리 솔, 중 + 단
            소리 미, 중 + 단 + 단 + 단 + 단 + 단
            
            소리 레, 중 + 단
            소리 미, 중 + 단 + 단 + 단 + 단 + 단

            소리 도, 중 + 40
            소리 미
            소리 도
            소리 시 / 옥, 장 + 단 + 단 + 단
            
            소리 솔, 중 + 40
            소리 미
            소리 도, 장 + 단 + 단 + 단 + 단
            
            소리 라 / 옥, 중 + 40
            소리 시 / 옥
            소리 시 / 옥, 중 + 중 + 단 + 단
            
            소리 시 / 옥, 중 + 40
            소리 라 / 옥
            소리 시 / 옥
            소리 도, 중 + 중 + 단 + 단
    End Select
End Sub

