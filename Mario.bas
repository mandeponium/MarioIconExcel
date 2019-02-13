Attribute VB_Name = "Mario"
Option Explicit
#If VBA7 Then
    '64 bit declares here
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As Long, ByVal bErase As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    '32 bit declares here
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Long, ByVal bErase As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If


Public Const KeyPressed As Integer = -32767
Private Const WM_SETREDRAW As Long = &HB&
Private Const WM_USER = &H400
Private Const EM_GETEVENTMASK = (WM_USER + 59)
Private Const EM_SETEVENTMASK = (WM_USER + 69)
Private Const KEY_DOWN    As Integer = &H8000   'If the most significant bit is set, the key is down
Private Const KEY_PRESSED As Integer = &H1      'If the least significant bit is set, the key was pressed after the previous call to GetAsyncKeyState

Public marioCoords() As Variant


Public xCurrent As Long
Public yCurrent As Long


Sub InitializeMario()
    ReDim marioCoords(1 To 13, 1 To 16)
    
    marioCoords(1, 1) = -4142
    marioCoords(2, 1) = -4142
    marioCoords(3, 1) = -4142
    marioCoords(4, 1) = 3
    marioCoords(5, 1) = 3
    marioCoords(6, 1) = 3
    marioCoords(7, 1) = 3
    marioCoords(8, 1) = 3
    marioCoords(9, 1) = 3
    marioCoords(10, 1) = -4142
    marioCoords(11, 1) = -4142
    marioCoords(12, 1) = -4142
    marioCoords(13, 1) = -4142
    marioCoords(1, 2) = -4142
    marioCoords(2, 2) = -4142
    marioCoords(3, 2) = 3
    marioCoords(4, 2) = 3
    marioCoords(5, 2) = 3
    marioCoords(6, 2) = 3
    marioCoords(7, 2) = 3
    marioCoords(8, 2) = 3
    marioCoords(9, 2) = 3
    marioCoords(10, 2) = 3
    marioCoords(11, 2) = 3
    marioCoords(12, 2) = 3
    marioCoords(13, 2) = -4142
    marioCoords(1, 3) = -4142
    marioCoords(2, 3) = -4142
    marioCoords(3, 3) = 12
    marioCoords(4, 3) = 12
    marioCoords(5, 3) = 12
    marioCoords(6, 3) = 40
    marioCoords(7, 3) = 40
    marioCoords(8, 3) = 40
    marioCoords(9, 3) = 1
    marioCoords(10, 3) = 40
    marioCoords(11, 3) = -4142
    marioCoords(12, 3) = -4142
    marioCoords(13, 3) = -4142
    marioCoords(1, 4) = -4142
    marioCoords(2, 4) = 12
    marioCoords(3, 4) = 40
    marioCoords(4, 4) = 12
    marioCoords(5, 4) = 40
    marioCoords(6, 4) = 40
    marioCoords(7, 4) = 40
    marioCoords(8, 4) = 40
    marioCoords(9, 4) = 1
    marioCoords(10, 4) = 40
    marioCoords(11, 4) = 40
    marioCoords(12, 4) = 40
    marioCoords(13, 4) = -4142
    marioCoords(1, 5) = -4142
    marioCoords(2, 5) = 12
    marioCoords(3, 5) = 40
    marioCoords(4, 5) = 12
    marioCoords(5, 5) = 12
    marioCoords(6, 5) = 40
    marioCoords(7, 5) = 40
    marioCoords(8, 5) = 40
    marioCoords(9, 5) = 40
    marioCoords(10, 5) = 1
    marioCoords(11, 5) = 40
    marioCoords(12, 5) = 40
    marioCoords(13, 5) = 40
    marioCoords(1, 6) = -4142
    marioCoords(2, 6) = 12
    marioCoords(3, 6) = 12
    marioCoords(4, 6) = 40
    marioCoords(5, 6) = 40
    marioCoords(6, 6) = 40
    marioCoords(7, 6) = 40
    marioCoords(8, 6) = 40
    marioCoords(9, 6) = 1
    marioCoords(10, 6) = 1
    marioCoords(11, 6) = 1
    marioCoords(12, 6) = 1
    marioCoords(13, 6) = -4142
    marioCoords(1, 7) = -4142
    marioCoords(2, 7) = -4142
    marioCoords(3, 7) = -4142
    marioCoords(4, 7) = 40
    marioCoords(5, 7) = 40
    marioCoords(6, 7) = 40
    marioCoords(7, 7) = 40
    marioCoords(8, 7) = 40
    marioCoords(9, 7) = 40
    marioCoords(10, 7) = 40
    marioCoords(11, 7) = 40
    marioCoords(12, 7) = -4142
    marioCoords(13, 7) = -4142
    marioCoords(1, 8) = -4142
    marioCoords(2, 8) = -4142
    marioCoords(3, 8) = 3
    marioCoords(4, 8) = 3
    marioCoords(5, 8) = 23
    marioCoords(6, 8) = 3
    marioCoords(7, 8) = 3
    marioCoords(8, 8) = 23
    marioCoords(9, 8) = 3
    marioCoords(10, 8) = -4142
    marioCoords(11, 8) = -4142
    marioCoords(12, 8) = -4142
    marioCoords(13, 8) = -4142
    marioCoords(1, 9) = -4142
    marioCoords(2, 9) = 3
    marioCoords(3, 9) = 3
    marioCoords(4, 9) = 3
    marioCoords(5, 9) = 23
    marioCoords(6, 9) = 3
    marioCoords(7, 9) = 3
    marioCoords(8, 9) = 23
    marioCoords(9, 9) = 3
    marioCoords(10, 9) = 3
    marioCoords(11, 9) = 3
    marioCoords(12, 9) = -4142
    marioCoords(13, 9) = -4142
    marioCoords(1, 10) = 3
    marioCoords(2, 10) = 3
    marioCoords(3, 10) = 3
    marioCoords(4, 10) = 3
    marioCoords(5, 10) = 23
    marioCoords(6, 10) = 23
    marioCoords(7, 10) = 23
    marioCoords(8, 10) = 23
    marioCoords(9, 10) = 3
    marioCoords(10, 10) = 3
    marioCoords(11, 10) = 3
    marioCoords(12, 10) = 3
    marioCoords(13, 10) = -4142
    marioCoords(1, 11) = 40
    marioCoords(2, 11) = 40
    marioCoords(3, 11) = 3
    marioCoords(4, 11) = 23
    marioCoords(5, 11) = 6
    marioCoords(6, 11) = 23
    marioCoords(7, 11) = 23
    marioCoords(8, 11) = 6
    marioCoords(9, 11) = 23
    marioCoords(10, 11) = 3
    marioCoords(11, 11) = 40
    marioCoords(12, 11) = 40
    marioCoords(13, 11) = -4142
    marioCoords(1, 12) = 40
    marioCoords(2, 12) = 40
    marioCoords(3, 12) = 40
    marioCoords(4, 12) = 23
    marioCoords(5, 12) = 23
    marioCoords(6, 12) = 23
    marioCoords(7, 12) = 23
    marioCoords(8, 12) = 23
    marioCoords(9, 12) = 23
    marioCoords(10, 12) = 40
    marioCoords(11, 12) = 40
    marioCoords(12, 12) = 40
    marioCoords(13, 12) = -4142
    marioCoords(1, 13) = 40
    marioCoords(2, 13) = 40
    marioCoords(3, 13) = 23
    marioCoords(4, 13) = 23
    marioCoords(5, 13) = 23
    marioCoords(6, 13) = 23
    marioCoords(7, 13) = 23
    marioCoords(8, 13) = 23
    marioCoords(9, 13) = 23
    marioCoords(10, 13) = 23
    marioCoords(11, 13) = 40
    marioCoords(12, 13) = 40
    marioCoords(13, 13) = -4142
    marioCoords(1, 14) = -4142
    marioCoords(2, 14) = -4142
    marioCoords(3, 14) = 23
    marioCoords(4, 14) = 23
    marioCoords(5, 14) = 23
    marioCoords(6, 14) = -4142
    marioCoords(7, 14) = -4142
    marioCoords(8, 14) = 23
    marioCoords(9, 14) = 23
    marioCoords(10, 14) = 23
    marioCoords(11, 14) = -4142
    marioCoords(12, 14) = -4142
    marioCoords(13, 14) = -4142
    marioCoords(1, 15) = -4142
    marioCoords(2, 15) = 12
    marioCoords(3, 15) = 12
    marioCoords(4, 15) = 12
    marioCoords(5, 15) = -4142
    marioCoords(6, 15) = -4142
    marioCoords(7, 15) = -4142
    marioCoords(8, 15) = -4142
    marioCoords(9, 15) = 12
    marioCoords(10, 15) = 12
    marioCoords(11, 15) = 12
    marioCoords(12, 15) = -4142
    marioCoords(13, 15) = -4142
    marioCoords(1, 16) = 12
    marioCoords(2, 16) = 12
    marioCoords(3, 16) = 12
    marioCoords(4, 16) = 12
    marioCoords(5, 16) = -4142
    marioCoords(6, 16) = -4142
    marioCoords(7, 16) = -4142
    marioCoords(8, 16) = -4142
    marioCoords(9, 16) = 12
    marioCoords(10, 16) = 12
    marioCoords(11, 16) = 12
    marioCoords(12, 16) = 12
    marioCoords(13, 16) = -4142

End Sub
    


Private Function ReadDirectionKeyDown() As String
    ReadDirectionKeyDown = ""
    
    If (GetAsyncKeyState(vbKeyUp) And KEY_DOWN) = KEY_DOWN Then
        ReadDirectionKeyDown = "up"
    ElseIf (GetAsyncKeyState(vbKeyDown) And KEY_DOWN) = KEY_DOWN Then
        ReadDirectionKeyDown = "down"
    ElseIf (GetAsyncKeyState(vbKeyRight) And KEY_DOWN) = KEY_DOWN Then
        ReadDirectionKeyDown = "right"
    ElseIf (GetAsyncKeyState(vbKeyLeft) And KEY_DOWN) = KEY_DOWN Then
        ReadDirectionKeyDown = "left"
    End If

End Function
'
'
'Private Function ReadDirectionFromKey() As String
'
'    ReadDirectionFromKey = ""
'
'    Select Case True
'
'        Case GetAsyncKeyState(vbKeyRight):
'            ReadDirectionFromKey = "right"
'
'        Case GetAsyncKeyState(vbKeyLeft):
'            ReadDirectionFromKey = "left"
'
'        Case GetAsyncKeyState(vbKeyUp):
'            ReadDirectionFromKey = "up"
'
'        Case GetAsyncKeyState(vbKeyDown):
'            ReadDirectionFromKey = "down"
'
'    End Select
'
'End Function


Sub ResetBoard()

    Sheets(1).Cells.Delete
    Columns("A:FZ").ColumnWidth = 2.2

End Sub

Sub Game()
    
    DoEvents
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    Application.ScreenUpdating = True
    
    InitializeMario
    
    
    Dim direction As String
    
    x = 3
    y = 8
    xCurrent = x
    yCurrent = y
    
    Dim increment As Long
    
    increment = 3
    
    
    Dim lastFrameTime As Long
    lastFrameTime = timeGetTime
    
    Do
        DoEvents
        
        'All game code goes here.
        '*********************************
                
        'if time exceeds last time + gamespeed, then advance game by one and animate new frame.
        If timeGetTime - lastFrameTime > 50 Then
            DoEvents
            
            direction = ReadDirectionKeyDown
            
            
            Select Case direction
            
                Case "up"
                    y = y - increment
                Case "down"
                    y = y + increment
                Case "left"
                    x = x - increment
                Case "right"
                    x = x + increment
                    
                Case Else
                
            End Select
            
            
            AnimateMario x, y, direction, increment
            
            lastFrameTime = timeGetTime
        End If

        '*********************************
                
    Loop
    
    
    ResetBoard

End Sub


Sub AnimateMario(xStart As Long, yStart As Long, direction As String, increment As Long)
    
    '13 by 16
    
    Dim x As Long
    Dim y As Long
    
    'clear the frame first
    'Range(Cells(yStart, xStart), Cells(yStart + 13, xStart + 16)).Interior.ColorIndex = -4142
    
    
    'Here's a single frame
    For x = 0 To 12
        For y = 0 To 15
            
            'If marioCoords(x + 1, y + 1) <> -4142 Then
            
                Cells(yStart + y, xStart + x).Interior.ColorIndex = marioCoords(x + 1, y + 1)
                
            'End If
            
        Next
    Next
    
    
    
    Select Case direction

        Case "up"
            Range(Cells(yStart + 16, xStart), Cells(yStart + 17 + increment, xStart + 13)).Interior.ColorIndex = -4142
        
        Case "down"
            Range(Cells(yStart - 1, xStart), Cells(yStart - 1 - increment, xStart + 13)).Interior.ColorIndex = -4142
        
        Case "left"
            Range(Cells(yStart, xStart + 13), Cells(yStart + 16, xStart + 14 + increment)).Interior.ColorIndex = -4142
        
        Case "right"
            Range(Cells(yStart, xStart - 1 - increment), Cells(yStart + 16, xStart - 1)).Interior.ColorIndex = -4142

        Case Else ' no movement

    End Select


    
    'else no movement



'-4142   -4142   -4142   3   3   3   3   3   3   -4142   -4142   -4142   -4142
'-4142   -4142   3   3   3   3   3   3   3   3   3   3   -4142
'-4142   -4142   12  12  12  40  40  40  1   40  -4142   -4142   -4142
'-4142   12  40  12  40  40  40  40  1   40  40  40  -4142
'-4142   12  40  12  12  40  40  40  40  1   40  40  40
'-4142   12  12  40  40  40  40  40  1   1   1   1   -4142
'-4142   -4142   -4142   40  40  40  40  40  40  40  40  -4142   -4142
'-4142   -4142   3   3   23  3   3   3   3   -4142   -4142   -4142   -4142
'-4142   3   3   3   23  3   3   23  3   3   3   -4142   -4142
'3   3   3   3   23  23  23  23  3   3   3   3   -4142
'40  40  3   23  6   23  23  6   23  3   40  40  -4142
'40  40  40  23  23  23  23  23  23  40  40  40  -4142
'40  40  23  23  23  23  23  23  23  23  40  40  -4142
'-4142   -4142   23  23  23  -4142   -4142   23  23  23  -4142   -4142   -4142
'-4142   12  12  12  -4142   -4142   -4142   -4142   12  12  12  -4142   -4142
'12  12  12  12  -4142   -4142   -4142   -4142   12  12  12  12  -4142


End Sub

