Attribute VB_Name = "Module2"
'@@@@@@@@@ Developed by Ravindra Deuskar @@@@@@@@@@@@@@@@@@@@
Option Explicit
'###############
    'Subclass flash activex control
    'trap all messages pass to original window
    'procedure except right mouse messages
'###############

Public glPrevWndProc As Long
Public FHW As Long
Public Const GWL_WNDPROC = (-4)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_KEYDOWN = &H100
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Sub UnSubClass()
Call SetWindowLong(FHW, GWL_WNDPROC, glPrevWndProc)
FHW = 0
Form1.Caption = "...."
End Sub

Public Function MyWindowProc(ByVal HW As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_RBUTTONDOWN Then
        Form1.Caption = "Right Button Down"
        Form1.PopupMenu Form1.mnuPop
        Exit Function
    ElseIf uMsg = WM_RBUTTONUP Then
        Form1.Caption = "Right Button Up"
        Exit Function
    ElseIf uMsg = WM_KEYDOWN Then
         Form1.Caption = "Key Down"
    End If
    MyWindowProc = CallWindowProc(glPrevWndProc, HW, uMsg, wParam, lParam)
End Function

Public Function SubClass() As Long
    SubClass = SetWindowLong(FHW, GWL_WNDPROC, AddressOf MyWindowProc)
    Form1.Caption = "I have disabled right click nenu"
End Function
