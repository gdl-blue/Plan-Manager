Attribute VB_Name = "SubClassControl"

Global gHookHWND As Long
 
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type PAINTSTRUCT
    hdc                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

Private Const GWL_WNDPROC = (-4)
Private Const STRETCHMODE = vbPaletteModeContainer


Private Const WM_PAINT = &HF

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private pRect As RECT

Public Function SubClassSSTAB(MySSTAB As SSTab, Pct As PictureBox)
    Pct.AutoRedraw = True
    Pct.AutoSize = True
    Pct.ScaleMode = vbPixels
    Pct.BackColor = Pct.BackColor
    
    'Save Grid fontname to use with DC's
    SetProp MySSTAB.hwnd, "lpPROC", SetWindowLong(MySSTAB.hwnd, GWL_WNDPROC, AddressOf MySubclassedGrid)
    SetProp MySSTAB.hwnd, "PctOBJ", ObjPtr(Pct)      'Save a pointer to PictureBox
    SetProp MySSTAB.hwnd, "GridOBJ", ObjPtr(MySSTAB)  'Save a pointer to Control
End Function

Public Sub UnSubClassSSTAB(ByVal hw As Long)
    Dim RetVal As Long
    RetVal = SetWindowLong(hw, GWL_WNDPROC, GetProp(hw, "lpPROC")) 'unsubclass Control
    'Clean up windows database
    RemoveProp hw, "lpPROC"
    RemoveProp hw, "PctOBJ"
    RemoveProp hw, "GridOBJ"
End Sub

Private Function MySubclassedGrid(ByVal hw As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim PicTEMP As PictureBox
Dim PicBACKGROUND As PictureBox
Dim GridTEMP As SSTab, GridREAL As SSTab

    gHookHWND = hw
    
    'Make GridTEMP a illegal reference - do not press END - Crash
    CopyMemory GridTEMP, GetProp(hw, "GridOBJ"), 4
    
    'Make it legal
    Set GridREAL = GridTEMP
    
    'Destroy illegal - no more crash
    CopyMemory GridTEMP, 0&, 4
    
    'Same story for PicTEMP
    CopyMemory PicTEMP, GetProp(hw, "PctOBJ"), 4
    Set PicBACKGROUND = PicTEMP
    CopyMemory PicTEMP, 0&, 4

    Select Case lMsg
         Case Is = WM_PAINT
            
            'We must do all the painting job
            Dim controlDC As Long, tempDC As Long, intDC As Long, tempBMP, intBMP As Long
            Dim aPS As PAINTSTRUCT
            Dim aDC As Long
            Dim Altura As Long
            Dim tppX, tppY As Long
            Dim BackBuffDC, BackBuffBMP As Long
            GetClientRect hw, pRect
            
                        
            'Start painting control ...
            Call BeginPaint(hw, aPS)
            aDC = aPS.hdc 'store painting DC
            
            'Prepare Double buffering ...No flickering
            BackBuffDC = CreateCompatibleDC(aDC)
            BackBuffBMP = CreateCompatibleBitmap(aDC, pRect.Right, pRect.Bottom)
            DeleteObject SelectObject(BackBuffDC, BackBuffBMP)
            
            'This is the big thing ! We are sendind WM_PAINT to our backbuffer
            MySubclassedGrid = CallWindowProc(GetProp(hw, "lpPROC"), hw, lMsg, ByVal BackBuffDC, 0&)
                    
            With pRect
              'We just want to place a background picture, so let's Strech it
              Call SetStretchBltMode(BackBuffDC, STRETCHMODE)
                    
              Call StretchBlt(BackBuffDC, tppX, tppY, pRect.Right, pRect.Bottom, _
                    PicBACKGROUND.hdc, 0, 0, PicBACKGROUND.ScaleWidth, PicBACKGROUND.ScaleHeight, vbSrcAnd)
            End With
            
            'We have all the changes into backbuffer. Let's bring in back to control.hDc
            With aPS.rcPaint
               BitBlt aDC, .Left, .Top, .Right - .Left, .Bottom - .Top, BackBuffDC, .Left, .Top, vbSrcCopy
            End With
            
            DeleteDC BackBuffDC
            DeleteObject BackBuffBMP
            Call EndPaint(hw, aPS)
            MySubclassedGrid = 0 'When a function intercepts WM_PAINT it must return 0
            
        Case Else
            'Call default windows procedure, stored in windows database in propertie lpPROC
            MySubclassedGrid = CallWindowProc(GetProp(hw, "lpPROC"), hw, lMsg, wParam, lParam)
    End Select
End Function




