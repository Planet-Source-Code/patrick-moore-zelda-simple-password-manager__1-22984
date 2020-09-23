Attribute VB_Name = "SysTrayModule"
'      Need to add to form using this:
'      PS: Also remember to "RemoveFromTray" when your form unloads
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Dim Message As Long
'   On Error Resume Next
'    Message = x / Screen.TwipsPerPixelX
'    Select Case Message
'        'Your Choice:
'        Case WM_RBUTTONUP
'            PopupMenu [Menu]
'        Case WM_RBUTTONDOWN
'            PopupMenu [Menu]
'    End Select
'End Sub

Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202

Global TrayIcon As NOTIFYICONDATA

Public Sub AddToTray(frm As Form, ToolTip As String, Icon)
On Error Resume Next
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frm.hwnd
TrayIcon.szTip = ToolTip & vbNullChar
TrayIcon.hIcon = Icon
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
TrayIcon.uCallbackMessage = WM_MOUSEMOVE

Shell_NotifyIcon NIM_ADD, TrayIcon

End Sub

Public Sub RemoveFromTray()
Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
