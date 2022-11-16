Attribute VB_Name = "Module1"
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal Message As Long, Data As NotifyIconData) As Boolean

Private Type NotifyIconData
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200

Private nidTrayIcon As NotifyIconData

Public lngCountdown As Long
Public lngPeriod As Long

'4 ticks per second, to match the 250 ms timer interval
'This period allows tray icon to flash at 2 Hz
Public Const TICKS_PER_SECOND = 4

Sub AddTrayIcon()
    nidTrayIcon.cbSize = Len(nidTrayIcon)
    nidTrayIcon.hWnd = frmConfig.hWnd
    nidTrayIcon.uID = vbNull
    nidTrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nidTrayIcon.uCallBackMessage = WM_MOUSEMOVE 'per https://support.microsoft.com/en-us/kb/162613
    nidTrayIcon.hIcon = frmConfig.picIcon.Picture
    nidTrayIcon.szTip = "DeskTimer - Left-click to reset, right-click for menu" & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
End Sub

Sub RefreshTrayIcon()
    'Add the image from picIcon into an imagelist control
    frmConfig.imgList.ListImages.Add 1, , frmConfig.picIcon.Image
    'Extract icon from that image to use as new tray icon
    nidTrayIcon.hIcon = frmConfig.imgList.ListImages(1).ExtractIcon
    Call Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
    'Remove the image from the imagelist, no longer required
    frmConfig.imgList.ListImages.Remove 1
End Sub

Sub RemoveTrayIcon()
    Call Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
End Sub

Sub LoadSettings()
    lngPeriod = CInt(GetSetting("DeskTimer", "Settings", "Period", "10"))
End Sub

Sub SaveSettings()
    SaveSetting "DeskTimer", "Settings", "Period", CStr(lngPeriod)
End Sub

Sub Main()
    Load frmConfig
    LoadSettings
    lngCountdown = (lngPeriod * 60 + 1) * TICKS_PER_SECOND
    AddTrayIcon
End Sub
