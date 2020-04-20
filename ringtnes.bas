Attribute VB_Name = "ringtones"
Sub PlayNotification(Optional ByVal Number As Integer = -1)
    Dim playn As Integer
    playn = Number
    
    If Number = -1 Then playn = CInt(GetSetting("Calendar", "Options", "Notification", 0))

    Select Case playn
        Case 0
            Beep 750, 300
            Beep 750, 300
        Case 1
            Beep 950, 700
    End Select
End Sub

Sub PlayRingtone(Optional ByVal Number As Integer = -1)
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
            
'            Beep 560, 300 '도
'            Beep 620, 300 '레
'            Beep 690, 300 '미
'            Beep 770, 300 '파
'            Beep 850, 300 '솔
'            Beep 950, 300 '라
'            Beep 1020, 300 '시
'            Beep 1090, 300 '도+1
            
            Beep 560, 200 '도
            Beep 690, 200 '미
            Beep 560, 200 '도
            Beep 460, 600 '시
            
            Beep 440, 200
            Beep 690, 200 '미
            Beep 440, 200
            Beep 390, 600
            
            Beep 770, 300 '파
            Beep 690, 300 '미
            
            Beep 560, 200 '도
            Beep 690, 200 '미
            Beep 560, 200 '도
            Beep 460, 600 '시
            
            Beep 770, 300 '파
            Beep 690, 500 '미
            
            Beep 600, 300 '레
            Beep 690, 600 '미
            
            Beep 560, 200 '도
            Beep 690, 200 '미
            Beep 560, 200 '도
            Beep 460, 600 '시
            
            Beep 690, 200 '미
            Beep 850, 200 '솔
            Beep 690, 200 '미
            Beep 560, 600 '도
            
            Beep 540, 300
            Beep 500, 500
            
            Beep 390, 300 '라
            Beep 460, 200 '시
            Beep 460, 500 '시
            
            Beep 460, 300 '시
            Beep 390, 300 '라
            Beep 460, 300 '시
            Beep 620, 300 '레
            Beep 560, 600 '도
    End Select
End Sub

