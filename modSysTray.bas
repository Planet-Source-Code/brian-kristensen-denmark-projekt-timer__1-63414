Attribute VB_Name = "modSysTray"
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_DELETE = &H2
Private Const NIM_ADD = 0

Private Const WM_USER = &H400
Private Const WM_HOOK = WM_USER + 1

Private Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = (-20)

Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SW_RESTORE = 9


Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

Private Const WS_EX_APPWINDOW = &H40000

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Private mFrm As VB.Form
Private mMenu As VB.Menu
Private mlDefProc As Long
Private mbLoadedInTray As Boolean

Private Function SetTrayData(ByRef oFrm As VB.Form, Optional ByVal sTitle As String) As NOTIFYICONDATA
      Dim TD As NOTIFYICONDATA
10    On Error GoTo SetTrayData_Error
20        TD.hwnd = oFrm.hwnd
30        TD.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
40        TD.hIcon = oFrm.Icon.Handle
50        TD.uID = 125&
60        TD.uCallbackMessage = WM_USER + 1
70        TD.cbSize = Len(TD)
80        If Len(sTitle) = 0 Then
90            If Len(oFrm.Tag) > 0 Then
100               If IsNumeric(oFrm.Tag) Then
110                   sTitle = oFrm.Caption
120               Else
130                   sTitle = oFrm.Tag
140               End If
150           Else
160               sTitle = oFrm.Caption
170           End If
180       End If
190       TD.szTip = sTitle & Chr$(0)
200       SetTrayData = TD
210   Exit Function
SetTrayData_Error:
220       LogError "Log in Function SetTrayData of Module modSysTray"
230       Resume Next ' cdlCancel
          'Err.Raise vbObjectError + Err.Number, "ProjectTimer - SetTrayData [" & Erl & "]", Err.Description
End Function

Public Function AddToTray(ByRef oFrm As VB.Form, ByRef oMenu As VB.Menu, Optional ByVal sTitle As String) As Boolean
Dim lReturn As Long
Dim TD As NOTIFYICONDATA
    
    On Error GoTo errTrap
    AddToTray = True
    If Not mbLoadedInTray Then
    
        TD = SetTrayData(oFrm, sTitle)
        
        lReturn = Shell_NotifyIcon(NIM_ADD, TD)
        
        Set mFrm = oFrm
        Set mMenu = oMenu
    
        mlDefProc = SetWindowLong(oFrm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
        mbLoadedInTray = True
    End If

Exit Function
errTrap:
    Debug.Assert False
    AddToTray = False
    Err.Raise Err.Number, "clsSysTray.AddToTray", Err.Description
End Function

Public Function RemoveFromTray(ByRef oFrm As VB.Form, Optional ByVal sTitle As String) As Boolean
Dim lReturn As Long
Dim TD As NOTIFYICONDATA
    On Error GoTo errTrap
    RemoveFromTray = True
    mlDefProc = 0
    TD = SetTrayData(oFrm, sTitle)
    lReturn = Shell_NotifyIcon(NIM_DELETE, TD)
    mbLoadedInTray = False
Exit Function
errTrap:
    Debug.Assert False
    RemoveFromTray = False
    Err.Raise Err.Number, "clsSysTray.RemoveFromTray", Err.Description
End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
10    On Error GoTo WindowProc_Error
          'If mlDefProc <> 0 Then
20            Select Case hwnd
                  Case mFrm.hwnd
30                    Select Case uMSG
                          Case WM_HOOK
40                            Select Case lParam
                                  Case WM_LBUTTONDOWN
                                      'MsgBox "Left Clicked..!!"
50                                Case WM_RBUTTONDOWN
                                      'MsgBox "Right Clicked..!!"
60                                    mFrm.PopupMenu mMenu
70                            End Select
80                        Case Else
90                            WindowProc = CallWindowProc(mlDefProc, hwnd, uMSG, wParam, lParam)
100                           Exit Function
110                   End Select
120               Case Else
130                   WindowProc = CallWindowProc(mlDefProc, hwnd, uMSG, wParam, lParam)
140                   Exit Function
150           End Select
          'End If
160   Exit Function
WindowProc_Error:
170       LogError "Log in Function WindowProc of Module modSysTray"
180       Resume Next ' cdlCancel
          'Err.Raise vbObjectError + Err.Number, "ProjectTimer - WindowProc [" & Erl & "]", Err.Description
End Function

Public Sub DoShowInTaskbar(ByRef oFrm As VB.Form, ByVal Value As Boolean)
10    On Error GoTo DoShowInTaskbar_Error
         ' Set WS_EX_APPWINDOW On or Off as requested.
         ' Toggling this value requires that we also toggle
         ' visibility, flipping the bit while invisible,
         ' forcing the taskbar to update on reshow.
         ' Using LockWindowUpdate prevents some flicker.
20       Call LockWindowUpdate(oFrm.hwnd)
30       Call ShowWindow(oFrm.hwnd, vbHide)
40       Call FlipBitEx(oFrm, WS_EX_APPWINDOW, Value)
50       Call ShowWindow(oFrm.hwnd, vbNormalFocus)
60       Call LockWindowUpdate(0&)
70    Exit Sub
DoShowInTaskbar_Error:
80        LogError "Log in Sub DoShowInTaskbar of Module modSysTray"
90        Resume Next ' cdlCancel
          'Err.Raise vbObjectError + Err.Number, "ProjectTimer - DoShowInTaskbar [" & Erl & "]", Err.Description
End Sub

Private Function FlipBitEx(ByRef oFrm As VB.Form, ByVal Bit As Long, ByVal Value As Boolean) As Boolean
         Dim nStyleEx As Long
   
10    On Error GoTo FlipBitEx_Error
         ' Retrieve current style bits.
20       nStyleEx = GetWindowLong(oFrm.hwnd, GWL_EXSTYLE)
   
         ' Attempt to set requested bit On or Off,
         ' and redraw.
30       If Value Then
40          nStyleEx = nStyleEx Or Bit
50       Else
60          nStyleEx = nStyleEx And Not Bit
70       End If
80       Call SetWindowLong(oFrm.hwnd, GWL_EXSTYLE, nStyleEx)
90       Call Redraw(oFrm)
   
         ' Return success code.
100      FlipBitEx = (nStyleEx = GetWindowLong(oFrm.hwnd, GWL_EXSTYLE))
110   Exit Function
FlipBitEx_Error:
120       LogError "Log in Function FlipBitEx of Module modSysTray"
130       Resume Next ' cdlCancel
          'Err.Raise vbObjectError + Err.Number, "ProjectTimer - FlipBitEx [" & Erl & "]", Err.Description
End Function

Public Sub Redraw(ByRef oFrm As VB.Form)
         ' Redraw window with new style.
         Const swpFlags As Long = _
            SWP_FRAMECHANGED Or SWP_NOMOVE Or _
            SWP_NOZORDER Or SWP_NOSIZE
10    On Error GoTo Redraw_Error
20       SetWindowPos oFrm.hwnd, 0, 0, 0, 0, 0, swpFlags
30    Exit Sub
Redraw_Error:
40        LogError "Log in Sub Redraw of Module modSysTray"
50        Resume Next ' cdlCancel
          'Err.Raise vbObjectError + Err.Number, "ProjectTimer - Redraw [" & Erl & "]", Err.Description
End Sub

Public Function DisplayApplication(AppTitle As String) As Boolean
Dim lRetVal As Long
Dim lHwnd As Long
    lHwnd = FindWindow(vbNullString, AppTitle$)
    If lHwnd <> 0 Then
        Call SetForegroundWindow(lHwnd)
        If IsIconic(lHwnd) Then
            lRetVal = ShowWindow(lHwnd, SW_RESTORE)
        End If
    End If
End Function
