Attribute VB_Name = "modMouseWheel"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
        ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, _
        ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
        ByVal hHook As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
        ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef destination As Any, ByVal source As LongPtr, ByVal length As LongPtr)
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal className As String, ByVal windowName As String) As LongPtr
    Private Declare PtrSafe Function GetClientRect Lib "user32" ( _
        ByVal hwnd As LongPtr, ByRef rectangle As RECTAPI) As Long
    Private Declare PtrSafe Function ClientToScreen Lib "user32" ( _
        ByVal hwnd As LongPtr, ByRef point As POINTAPI) As Long
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
        ByVal moduleName As LongPtr) As LongPtr
#Else
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
        ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
        ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" ( _
        ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef destination As Any, ByVal source As Long, ByVal length As Long)
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal className As String, ByVal windowName As String) As Long
    Private Declare Function GetClientRect Lib "user32" ( _
        ByVal hwnd As Long, ByRef rectangle As RECTAPI) As Long
    Private Declare Function ClientToScreen Lib "user32" ( _
        ByVal hwnd As Long, ByRef point As POINTAPI) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
        ByVal moduleName As Long) As Long
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECTAPI
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
#If VBA7 Then
    extraInfo As LongPtr
#Else
    extraInfo As Long
#End If
End Type

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A

#If VBA7 Then
    Private mHook As LongPtr
#Else
    Private mHook As Long
#End If
Private mWheelForm As Object

Public Sub EnableMouseWheel(ByVal targetForm As Object)
    On Error Resume Next
    DisableMouseWheel
    Set mWheelForm = targetForm
#If VBA7 Then
    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseWheelHook, _
                            GetModuleHandle(0), 0)
#Else
    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseWheelHook, _
                            GetModuleHandle(0), 0)
#End If
    If mHook = 0 Then Set mWheelForm = Nothing
    On Error GoTo 0
End Sub

Public Sub DisableMouseWheel()
    On Error Resume Next
    If mHook <> 0 Then UnhookWindowsHookEx mHook
    mHook = 0
    Set mWheelForm = Nothing
    On Error GoTo 0
End Sub

#If VBA7 Then
Private Function MouseWheelHook(ByVal nCode As Long, ByVal wParam As LongPtr, _
                                ByVal lParam As LongPtr) As LongPtr
#Else
Private Function MouseWheelHook(ByVal nCode As Long, ByVal wParam As Long, _
                                ByVal lParam As Long) As Long
#End If
    Dim info As MSLLHOOKSTRUCT, delta As Long, handled As Boolean
    On Error Resume Next
    If nCode >= 0 And wParam = WM_MOUSEWHEEL And Not mWheelForm Is Nothing Then
        CopyMemory info, lParam, LenB(info)
        delta = (info.mouseData And &HFFFF0000) \ &H10000
        handled = CBool(mWheelForm.HandleMouseWheel(info.pt.x, info.pt.y, delta))
        If handled Then
            MouseWheelHook = 1
            Exit Function
        End If
    End If
    MouseWheelHook = CallNextHookEx(mHook, nCode, wParam, lParam)
    On Error GoTo 0
End Function

Public Function MousePointInForm(ByVal targetForm As Object, ByVal screenX As Long, _
                                 ByVal screenY As Long, ByRef formX As Single, _
                                 ByRef formY As Single) As Boolean
    Dim bounds As RECTAPI, origin As POINTAPI
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    On Error GoTo Unsupported
    hwnd = FindWindow(vbNullString, CStr(targetForm.Caption))
    If hwnd = 0 Or GetClientRect(hwnd, bounds) = 0 Then Exit Function
    If ClientToScreen(hwnd, origin) = 0 Then Exit Function
    If screenX < origin.x Or screenX > origin.x + bounds.right Then Exit Function
    If screenY < origin.y Or screenY > origin.y + bounds.bottom Then Exit Function
    formX = CSng(screenX - origin.x) * targetForm.InsideWidth / CSng(bounds.right)
    formY = CSng(screenY - origin.y) * targetForm.InsideHeight / CSng(bounds.bottom)
    MousePointInForm = True
Unsupported:
End Function

Public Sub ScrollListByWheel(ByVal listControl As Object, ByVal delta As Long)
    Dim newTop As Long
    On Error Resume Next
    If listControl.ListCount = 0 Then Exit Sub
    newTop = listControl.TopIndex - Sgn(delta) * 3
    If newTop < 0 Then newTop = 0
    If newTop > listControl.ListCount - 1 Then newTop = listControl.ListCount - 1
    listControl.TopIndex = newTop
    On Error GoTo 0
End Sub

Public Sub ScrollFrameByWheel(ByVal frameControl As Object, ByVal delta As Long)
    Dim newTop As Single, maximum As Single
    On Error Resume Next
    maximum = frameControl.ScrollHeight - frameControl.InsideHeight
    If maximum < 0 Then maximum = 0
    newTop = frameControl.ScrollTop - Sgn(delta) * 28
    If newTop < 0 Then newTop = 0
    If newTop > maximum Then newTop = maximum
    frameControl.ScrollTop = newTop
    On Error GoTo 0
End Sub
