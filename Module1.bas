Attribute VB_Name = "Module1"
Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = (-4)

Dim PrevProc As Long
Dim m_hWnd As Long

Public Sub Hook(hWnd As Long)
    PrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook()
    SetWindowLong m_hWnd, GWL_WNDPROC, PrevProc
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        If (wParam \ 65536) < 0 Then
            'WHEEL WAS ROLLED BACKWARDS
            Debug.Print "BACKWARDS!"
        Else
            'WHEEL WAS ROLLED FORWARDS
            Debug.Print "FORWARDS!"
        End If
    End If
    m_hWnd = hWnd
End Function

