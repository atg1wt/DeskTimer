VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DeskTimer Configuration"
   ClientHeight    =   1935
   ClientLeft      =   5325
   ClientTop       =   4875
   ClientWidth     =   7455
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin ComctlLib.Slider sldMinutes 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1085
      _Version        =   327682
      Max             =   90
      TickFrequency   =   5
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   240
      Left            =   7080
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer tmrTick 
      Interval        =   250
      Left            =   6240
      Top             =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   6720
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label lblMins 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Version 0.6 ~ 4th May 2016"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuConfig 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Sub ChangeSliderCaption()
    Dim strTemp As String
    strTemp = CStr(sldMinutes.Value) & " min"
    If sldMinutes.Value <> 1 Then strTemp = strTemp & "s"
    lblMins.Caption = strTemp
End Sub

Private Sub cmdOK_Click()
    frmConfig.Hide
    lngPeriod = sldMinutes.Value
    'If new timer period is shorter than current countdown, reduce current countdown accordingly
    If lngCountdown > (lngPeriod * 60 * TICKS_PER_SECOND) Then lngCountdown = (lngPeriod * 60 + 1) * TICKS_PER_SECOND
    SaveSettings
End Sub

Private Sub cmdCancel_Click()
    frmConfig.Hide
End Sub

Private Sub mnuConfig_Click()
    sldMinutes.Value = lngPeriod
    ChangeSliderCaption
    frmConfig.Show
End Sub

Private Sub mnuExit_Click()
    RemoveTrayIcon
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MouseMove used to capture messages from tray icon as per https://support.microsoft.com/en-us/kb/162613
    'Due to the notification message format, the X argument received here holds the mouse action.
    'Value is delivered according to current ScaleMode; this has been set to pixels so that an unscaled value is received.
    Select Case X
    Case WM_RBUTTONDOWN
        Call frmConfig.PopupMenu(mnuPopup)
    Case WM_LBUTTONDOWN
        lngCountdown = (lngPeriod * 60 + 1) * TICKS_PER_SECOND
    End Select
End Sub

Private Sub sldMinutes_Scroll()
    ChangeSliderCaption
End Sub

Private Sub tmrTick_Timer()
    'Count down
    lngCountdown = lngCountdown - 1
    'Prevent negative overflow, in case timer is left flashing "00" for a very long time
    If lngCountdown = -TICKS_PER_SECOND Then lngCountdown = 0
    'Tray icon appearance depends on time left
    If lngCountdown >= 60 * TICKS_PER_SECOND Then
        'More than a minute left, show minutes in green text
        picIcon.BackColor = &H0
        picIcon.ForeColor = &H80FF80
        strTrayText = Format(Int(lngCountdown / TICKS_PER_SECOND / 60), "00")
    ElseIf lngCountdown > 0 Then
        'Less than a minute left, show seconds in orange text
        picIcon.BackColor = &H0
        picIcon.ForeColor = &HA0FF&
        strTrayText = Format(Int(lngCountdown / TICKS_PER_SECOND), "00")
    Else
        'Time up, flash "00" in yellow/black text
        If lngCountdown Mod 2 Then
            picIcon.BackColor = &H80FFFF
            picIcon.ForeColor = &H0
        Else
            picIcon.BackColor = &H0
            picIcon.ForeColor = &H80FFFF
        End If
        strTrayText = "00"
    End If
    'Update tray icon image
    picIcon.Cls
    picIcon.CurrentX = (16 - picIcon.TextWidth(strTrayText)) / 2
    picIcon.CurrentY = 1
    picIcon.Print strTrayText
    RefreshTrayIcon
End Sub

