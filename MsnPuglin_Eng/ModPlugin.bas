Attribute VB_Name = "ModPlugin"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Public Type NOTIFYICONDATA 'SysTray Pointer
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'SysTray Vars
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIM_MODIFY = &H1

Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONUP = &H202

Dim TrayIcon As NOTIFYICONDATA
'*************

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Const WM_COPYDATA = &H4A

Public Const WM_COMMAND As Long = &H111&
Public Const WM_USER As Long = &H400&
Public WA_hwnd As Long

Public Const WA_GETSTATUS = 104

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'SysTray Api
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_COMMAND
            Debug.Print "comando de winamp"
        Case WM_USER
            Debug.Print "comando de winamp"
    End Select
   
   'Devolver a la cola de mensajes
   WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Function GetSongName() As String
    On Error Resume Next
    Dim sSong$
    Dim strTitle As String
    strTitle = String(2048, " ")
    
    GetWindowText WA_hwnd, strTitle, Len(strTitle)
    sSong = Left$(strTitle, InStr(strTitle, "- Winamp") - 1)
    GetSongName = Mid$(sSong, InStr(sSong, ".") + 1)
End Function

Function IsWinampActive() As Boolean
    WA_hwnd = FindWindow("Winamp v1.x", vbNullString)
    IsWinampActive = CBool(WA_hwnd)
End Function

Function AddToTray(frm As Form, ToolTip As String, Icon)
On Error Resume Next
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = frm.hwnd
    TrayIcon.szTip = ToolTip & vbNullChar
    TrayIcon.hIcon = Icon
    TrayIcon.uID = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    
    Shell_NotifyIcon NIM_ADD, TrayIcon

End Function

Function RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Function

Public Function GetY()
    Dim Point As POINTAPI, RetVal As Long
    RetVal = GetCursorPos(Point)
    GetY = Point.Y
End Function
