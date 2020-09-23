Attribute VB_Name = "modTrayIcon"
'API declaration
Private Declare Function SetForegroundWindow Lib "user32" _
(ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


'user defined type required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

'constants used, some not utilized
Private Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Private Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
 


'type of icon displayed in the balloon tooltip.
Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum
'type of tooltip to display
Public Enum ttStyle
    ttBalloon
    ttStandard
End Enum


'type definitions
Private Type POINT_TYPE
  X As Long
  Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




'variable declarations used across all subs and functions.
Private blLive As Boolean
Private nid As NOTIFYICONDATA

'sub used by api timer to make balloon disappear.
Public Sub TrayUpdate(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
        ByVal dwTime As Long)
    Dim pos As POINT_TYPE
    Dim lCheck As Boolean
    lCheck = whereAmI(pos, frmMain)
    If lCheck Then
        'if the mouse is over me do nothing
    Else
        'if mouse is not over set the tip to blank so it will disappear
        With nid
            .szInfo = "" & Chr(0)
            .szInfoTitle = "" & Chr(0)
        End With
        KillTimer hWnd, 1 'kill the timer so we don't come back.
        Shell_NotifyIcon NIM_MODIFY, nid 'modify to blank, balloon disappears
        blLive = False 'set to false so we can detect mouse move event again.
    End If
End Sub

'function to figure out if mouse is over tray icon
Public Function whereAmI(pos As POINT_TYPE, frm As Form) As Boolean
    Dim retval As Long
    Dim frmRect As RECT
'>----------------------check to see if the mouse is over me----------------------<
    retval = GetCursorPos(pos)
    retval = GetWindowRect(frm.hWnd, frmRect)
    whereAmI = False
     If pos.X > frmRect.Left And pos.X < frmRect.Right Then
        If pos.Y > frmRect.Top And pos.Y < frmRect.Bottom Then
            whereAmI = True
       End If
    End If
'>----------------------check to see if the mouse is over me----------------------<

End Function

'sub used to initially load the icon to the tray and set it's tt style etc.
Public Sub Do_Tray(frm As Form, sTip As String, sTitle As String, ttIcon As ttIconType, tStyle As ttStyle)
Attribute Do_Tray.VB_Description = "Pass the form that will be put in the tray, it's initial tooltip(used to notify that it is here.)"
    Dim lTipStyle As Long
    lTipStyle = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    If tStyle = ttBalloon Then lTipStyle = lTipStyle Or NIF_INFO ' if balloon add the NIF_INFO flag.
    With nid
            .cbSize = Len(nid)
            .hWnd = frm.hWnd
            .uID = vbNull
            .uFlags = lTipStyle 'set the tip style
            .uCallbackMessage = WM_MOUSEMOVE 'listen for mouse move
            .hIcon = frm.Icon
            If tStyle <> ttBalloon Then 'if standard set the tip
                .szTip = sTip & Chr(0)
            End If
            .szInfo = sTip & Chr(0) ' if set to blank and balloon style, nothing is shown.
            .szInfoTitle = sTitle & Chr(0)
            .dwInfoFlags = ttIcon 'set the icon type for the balloon
        End With
    Shell_NotifyIcon NIM_ADD, nid 'stick that puppy in the tray
    frm.Hide
End Sub

'to be used by the on mouse over event of the form.
Public Sub TrayMouseMoveBalloonTip(frm As Form, sTip As String, sTitle As String, ttIcon As ttIconType, tStyle As ttStyle, iTimeout As Integer)
Attribute TrayMouseMoveBalloonTip.VB_Description = "Called from the forms mouse move event, pass the X value through"
    Dim lResult As Long
    Dim lTipStyle As Long
    lTipStyle = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    If tStyle = ttBalloon Then lTipStyle = lTipStyle Or NIF_INFO
    Dim lMsg As Long
    If frm.ScaleMode = vbPixels Then
        lMsg = X
    Else
        lMsg = X / Screen.TwipsPerPixelX
    End If
    If blLive Then Exit Sub 'if true we have already fired the mousemove event.
      With nid
            .uFlags = lTipStyle 'set the tip style
            If tStyle <> ttBalloon Then 'if standard set the tip
                .szTip = sTip & Chr(0)
            End If
            .szInfo = sTip & Chr(0) ' if set to blank and balloon style, nothing is shown.
            .szInfoTitle = "Information" & Chr(0)
            .dwInfoFlags = ttIcon 'set the icon type for the balloon
      End With
    Shell_NotifyIcon NIM_MODIFY, nid 'modify the existing tray icon
    blLive = True 'we have been here
    ' after specified time fire the event to hide tooltip if not hovering over icon.
    lResult = SetTimer(frm.hWnd, 1, (iTimeout * 1000), AddressOf TrayUpdate)

End Sub

'destroy the tray icon.
Public Sub KillTrayApp()
Attribute KillTrayApp.VB_Description = "Call from the forms QueryUnload event"
    'get rid of that sucker.
    Shell_NotifyIcon NIM_DELETE, nid
End Sub
