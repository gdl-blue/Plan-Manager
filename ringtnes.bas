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
    End Select
End Sub

