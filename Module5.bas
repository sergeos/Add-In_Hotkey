Attribute VB_Name = "Module5"
Option Explicit

Public OldWindowProc As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Const MOD_ALT = &H1
Public Const VK_F10 = &H79

Public Const HOTKEY_ID = 1
Private Const HWND_MESSAGE              As Long = -3


Dim hWnd        As Long

' Look for the WM_HOTKEY message.
Public Function NewWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_NCDESTROY = &H82
    Const WM_HOTKEY = &H312

    ' If we're being destroyed,
    ' restore the original WindowProc and
    ' unregister the hotkey.
    If Msg = WM_NCDESTROY Then
        SetWindowLong hWnd, GWL_WNDPROC, OldWindowProc
        UnregisterHotKey hWnd, HOTKEY_ID
    End If

    ' See if this is the WM_HOTKEY message.
    If Msg = WM_HOTKEY Then Module5.Hotkey

    ' Process the message normally.
    NewWindowProc = CallWindowProc(OldWindowProc, hWnd, Msg, wParam, lParam)
End Function

Private Sub Hotkey()
    txtTimes.Text = txtTimes.Text & Time & vbCrLf
    txtTimes.SelStart = Len(txtTimes.Text)

    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    Me.SetFocus
End Sub

Public Sub InitHK()
    
    hWnd = CreateWindowEx(0, "myhWND", 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
    If hWnd = 0 Then Exit Sub

''--------------------------
    ' Register the hotkey.
    If RegisterHotKey(hWnd, HOTKEY_ID, MOD_ALT, VK_F10) = 0 Then
        MsgBox "Error registering hotkey."
        Exit Sub
        'Unload Me
    End If

    ' Subclass the TextBox to watch for
    ' WM_HOTKEY messages.
    OldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWindowProc)

End Sub
