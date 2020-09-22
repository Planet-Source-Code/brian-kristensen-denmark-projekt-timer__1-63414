VERSION 5.00
Begin VB.Form frmProjTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Timer"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmProjTimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTray 
      Height          =   375
      Left            =   5280
      Picture         =   "frmProjTimer.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Hide to tray"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtJob 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Job description"
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdReadLog 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmProjTimer.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "View current project file"
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   1440
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "&OnTop"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Set on top of other windows"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Height          =   375
      Left            =   5880
      Picture         =   "frmProjTimer.frx":138F6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Abort timer"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      Picture         =   "frmProjTimer.frx":1A148
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Stop and save to file"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Cancel          =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      Picture         =   "frmProjTimer.frx":2099A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pause"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      Picture         =   "frmProjTimer.frx":271EC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Start timer"
      Top             =   960
      Width           =   375
   End
   Begin VB.ComboBox cboProject 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Project Name"
      Top             =   120
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Not started"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label lblJob 
      Caption         =   "&Job descr.:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblProject 
      Caption         =   "&Project:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuStatus 
         Caption         =   "- Kører ikke"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Åben timer"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Luk"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop og gem"
      End
   End
End
Attribute VB_Name = "frmProjTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const msAppName = "Project Timer"

Private mdStartTimer As Double
Private mdPauseTime As Double
Private msLastTempLine As String
Private mdUsedPause As Double
Private mdLastAlive As Double


Private Sub SetTopMost(Optional ByVal bOnTop As Boolean = True)
    Const swpFlags = SWP_NOMOVE Or SWP_NOSIZE
On Error GoTo SetTopMost_Error
    If bOnTop Then
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, swpFlags)
    Else
        Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, swpFlags)
    End If
Exit Sub
SetTopMost_Error:
    LogError "Log in Sub SetTopMost of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - SetTopMost [" & Erl & "]", Err.Description
End Sub

Private Sub cboProject_Change()
On Error GoTo cboProject_Change_Error
    If Len(cboProject.Text) > 0 And Len(txtJob.Text) > 0 Then
        cmdStart.Enabled = True
    Else
        cmdStart.Enabled = False
    End If
Exit Sub
cboProject_Change_Error:
    LogError "Log in Sub cboProject_Change of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cboProject_Change [" & Erl & "]", Err.Description
End Sub

Private Sub cboProject_Click()
On Error GoTo cboProject_Click_Error
    If cboProject.ListIndex > -1 Then
        cmdReadLog.Enabled = True
    Else
        cmdReadLog.Enabled = False
    End If
Exit Sub
cboProject_Click_Error:
    LogError "Log in Sub cboProject_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cboProject_Click [" & Erl & "]", Err.Description
End Sub

Private Sub chkOnTop_Click()
On Error GoTo chkOnTop_Click_Error
    SetTopMost chkOnTop.Value = 1
Exit Sub
chkOnTop_Click_Error:
    LogError "Log in Sub chkOnTop_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - chkOnTop_Click [" & Erl & "]", Err.Description
End Sub

Private Sub cmdClose_Click()
On Error GoTo cmdClose_Click_Error
    Unload Me
Exit Sub
cmdClose_Click_Error:
    LogError "Log in Sub cmdClose_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdClose_Click [" & Erl & "]", Err.Description
End Sub

Private Function PrintString(ByVal sText As String) As String
On Error GoTo PrintString_Error
    PrintString = """"
    PrintString = PrintString & Replace(sText, """", "''")
    PrintString = PrintString & """"
Exit Function
PrintString_Error:
    LogError "Log in Function PrintString of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - PrintString [" & Erl & "]", Err.Description
End Function
Private Function PrintDate(ByVal dDate As Double) As String
On Error GoTo PrintDate_Error
    PrintDate = Year(dDate) & "-" & Right("0" & Month(dDate), 2) & "-" & Right("0" & Day(dDate), 2)
Exit Function
PrintDate_Error:
    LogError "Log in Function PrintDate of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - PrintDate [" & Erl & "]", Err.Description
End Function
Private Function PrintTime(ByVal dTime As Double) As String
On Error GoTo errTrap
    PrintTime = Format(dTime, "hh:nn:ss")
Exit Function
errTrap:
    Debug.Assert False
End Function
Private Function PrintDateTime(ByVal dDateTime As Double) As String
On Error GoTo PrintDateTime_Error
    PrintDateTime = PrintDate(dDateTime) & " " & PrintTime(dDateTime)
Exit Function
PrintDateTime_Error:
    LogError "Log in Function PrintDateTime of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - PrintDateTime [" & Erl & "]", Err.Description
End Function
Private Function PrintUsedTime() As String
Dim lMinute As Long
Dim lHours As Long
Dim sMin2 As String
Dim dTime As Double
On Error GoTo PrintUsedTime_Error
    dTime = Now - mdStartTimer
    lHours = Format(dTime, "hh")
    lMinute = Format(dTime, "nn")
    If lMinute < 16 Then
        sMin2 = "25"
    ElseIf lMinute < 31 Then
        sMin2 = "50"
    ElseIf lMinute < 46 Then
        sMin2 = "75"
    Else
        sMin2 = "00"
        lHours = lHours + 1
    End If
    
    PrintUsedTime = Format(lHours, "00") & "," & sMin2
Exit Function
PrintUsedTime_Error:
    LogError "Log in Function PrintUsedTime of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - PrintUsedTime [" & Erl & "]", Err.Description
End Function

Private Function PrintUsedTime2() As String
On Error GoTo PrintUsedTime2_Error
    PrintUsedTime2 = Format(Now - mdStartTimer, "hh:nn:ss")
Exit Function
PrintUsedTime2_Error:
    LogError "Log in Function PrintUsedTime2 of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - PrintUsedTime2 [" & Erl & "]", Err.Description
End Function

Private Function GetLogLine() As String
On Error GoTo GetLogLine_Error
    GetLogLine = PrintString(cboProject.Text)
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & PrintString(txtJob.Text)
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & PrintDateTime(mdStartTimer - mdUsedPause)
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & PrintDateTime(Now)
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & Format(mdUsedPause, "hh:nn:ss")
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & PrintUsedTime2
    GetLogLine = GetLogLine & ";"
    GetLogLine = GetLogLine & PrintUsedTime
Exit Function
GetLogLine_Error:
    LogError "Log in Function GetLogLine of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - GetLogLine [" & Erl & "]", Err.Description
End Function
Private Sub WriteTmpFile()
Dim lFile As Long
Dim sLine As String
On Error GoTo WriteTmpFile_Error
    sLine = GetLogLine
    If sLine <> msLastTempLine Then
        lFile = FreeFile
        Open msAppPath & "logging.tmp" For Output As #lFile
        Print #lFile, sLine
        Close lFile
        msLastTempLine = sLine
    End If
Exit Sub
WriteTmpFile_Error:
    LogError "Log in Sub WriteTmpFile of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - WriteTmpFile [" & Erl & "]", Err.Description
End Sub

Private Sub cmdPause_Click()
On Error GoTo cmdPause_Click_Error
    mdPauseTime = Now
    cmdStart.Enabled = True
    cmdPause.Enabled = False
    mnuPause.Enabled = cmdPause.Enabled
    mnuStart.Enabled = cmdStart.Enabled
    mnuStop.Enabled = cmdStop.Enabled
    mnuClose.Enabled = cmdClose.Enabled
Exit Sub
cmdPause_Click_Error:
    LogError "Log in Sub cmdPause_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdPause_Click [" & Erl & "]", Err.Description
End Sub

Private Sub cmdReadLog_Click()
On Error GoTo cmdReadLog_Click_Error
    OpenProject cboProject.Text
Exit Sub
cmdReadLog_Click_Error:
    LogError "Log in Sub cmdReadLog_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdReadLog_Click [" & Erl & "]", Err.Description
End Sub

Private Sub cmdStart_Click()
On Error GoTo cmdStart_Click_Error
    cmdStart.Enabled = False
    If mdPauseTime <> 0 Then
        mdStartTimer = Now - (mdPauseTime - mdStartTimer)
        mdUsedPause = mdUsedPause + (Now - mdPauseTime)
        mdPauseTime = 0
    Else
        mdStartTimer = Now
        mdPauseTime = 0
        mdUsedPause = 0
    End If
    WriteTmpFile
    cmdStop.Enabled = True
    cmdClose.Enabled = False
    txtJob.Enabled = False
    cboProject.Enabled = False
    cmdReadLog.Enabled = True
    cmdPause.Enabled = True
    mnuPause.Enabled = cmdPause.Enabled
    mnuStart.Enabled = cmdStart.Enabled
    mnuStop.Enabled = cmdStop.Enabled
    mnuClose.Enabled = cmdClose.Enabled
Exit Sub
cmdStart_Click_Error:
    LogError "Log in Sub cmdStart_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdStart_Click [" & Erl & "]", Err.Description
    
End Sub

Private Sub cmdStop_Click()
Dim sFile As String
Dim sProject As String
On Error GoTo cmdStop_Click_Error
    sProject = cboProject.Text
    If mdPauseTime <> 0 Then
        mdStartTimer = Now - (mdPauseTime - mdStartTimer)
        mdUsedPause = mdUsedPause + (Now - mdPauseTime)
        mdPauseTime = Now
    End If
    If AppendLineToProject(GetLogLine, cboProject.Text) Then
        cmdStop.Enabled = False
        Kill msAppPath & "logging.tmp"
        cmdStart.Enabled = True
        cmdClose.Enabled = True
        cmdPause.Enabled = False
        txtJob.Enabled = True
        cboProject.Enabled = True
                
        sFile = Dir(msAppPath & "*.csv")
        cboProject.Clear
        Do While Len(sFile) > 0
            If Left(sFile, 5) = "view_" Then
            ElseIf Right(sFile, 8) = ".log.csv" Then
            Else
                cboProject.AddItem Left$(sFile, Len(sFile) - 4)
            End If
            sFile = Dir
        Loop
        cboProject.Text = sProject
        lblStatus.Caption = "STOPPED AND SAVED          USED TIME: " & PrintUsedTime2
        
        mdStartTimer = 0 'Now
        mdPauseTime = 0
        mdUsedPause = 0
        Me.Caption = "Project Timer"
    End If
    mnuPause.Enabled = cmdPause.Enabled
    mnuStart.Enabled = cmdStart.Enabled
    mnuStop.Enabled = cmdStop.Enabled
    mnuClose.Enabled = cmdClose.Enabled
Exit Sub
cmdStop_Click_Error:
    LogError "Log in Sub cmdStop_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdStop_Click [" & Erl & "]", Err.Description
End Sub

Private Sub cmdTray_Click()
Dim sText As String
On Error GoTo cmdTray_Click_Error
    Timer1.Enabled = False
    Timer1.Interval = 0
    DoEvents
    sText = Me.Caption
    Me.Caption = ": " & cboProject.Text & " / " & txtJob.Text & " :"
    AddToTray Me, mnuHidden
    mnuClose.Enabled = cmdClose.Enabled
    Me.Visible = False
    mnuPause.Enabled = cmdPause.Enabled
    mnuStart.Enabled = cmdStart.Enabled
    mnuStop.Enabled = cmdStop.Enabled
    mnuClose.Enabled = cmdClose.Enabled
    Me.Caption = sText
    DoEvents
    Timer1.Enabled = True
    Timer1.Interval = 1000
Exit Sub
cmdTray_Click_Error:
    LogError "Log in Sub cmdTray_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - cmdTray_Click [" & Erl & "]", Err.Description
End Sub

Private Sub Form_Load()
Dim lFile As Long
Dim sLine As String
Dim sProject As String
Dim sTitle As String
On Error GoTo Form_Load_Error
    If App.PrevInstance Then
        On Error Resume Next
        Me.Visible = False
        sTitle = App.Title
        App.Title = ""
        DisplayApplication sTitle
        MsgBox "Programmet kører allerede!"
        Unload Me
        Exit Sub
    End If
    Me.Move Val(GetSetting(msAppName, "Formposition", "Left", CLng(Me.Left))), Val(GetSetting(msAppName, "Formposition", "Top", CLng(Me.Top)))
    chkOnTop.Value = Val(GetSetting(msAppName, "Formposition", "OnTop", 1))
    msAppPath = App.Path
    If Right(msAppPath, 1) <> "\" Then
        msAppPath = msAppPath & "\"
    End If
    If Len(Dir(msAppPath & "logging.tmp")) > 0 Then
        lFile = FreeFile
        Open msAppPath & "logging.tmp" For Input As #lFile
        Do While Not EOF(lFile)
            Line Input #lFile, sLine
            If Len(sLine) > 0 Then
                If InStr(1, sLine, """;""") > 0 Then
                    sProject = Replace(Split(sLine, ";")(0), """", "")
                    AppendLineToProject sLine, sProject
                End If
            End If
        Loop
        Close lFile
        Kill msAppPath & "logging.tmp"
    End If
    sProject = Dir(msAppPath & "*.csv")
    Do While Len(sProject) > 0
        If Left(sProject, 5) = "view_" Then
        ElseIf Right(sProject, 8) = ".log.csv" Then
        Else
            cboProject.AddItem Left$(sProject, Len(sProject) - 4)
        End If
        sProject = Dir
    Loop
    mdLastAlive = Now
Exit Sub
Form_Load_Error:
    LogError "Log in Sub Form_Load of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - Form_Load [" & Erl & "]", Err.Description
End Sub

Private Function AppendLineToProject(ByVal sLine As String, ByVal sProject As String) As Boolean
Dim lFile As Long
On Error GoTo errTrap
    AppendLineToProject = True
    lFile = FreeFile
    If Len(Dir(msAppPath & sProject & ".csv")) = 0 Then
        Open msAppPath & sProject & ".csv" For Output As #lFile
        Print #lFile, """Project"";""Job description"";""Start"";""End"";""Paused"";""Time"";""Rounded (0.25)"""
        Print #lFile, sLine
    Else
        FileSystem.SetAttr msAppPath & sProject & ".csv", vbNormal
        Open msAppPath & sProject & ".csv" For Append As #lFile
        Print #lFile, sLine
    End If
    Close lFile
    
    FileSystem.SetAttr msAppPath & sProject & ".csv", vbReadOnly
Exit Function
errTrap:
    AppendLineToProject = False
    MsgBox "Kunne ikke skrive til projektfilen." & vbCrLf & Err.Description
End Function

Private Sub OpenProject(ByVal sProject As String)
On Error GoTo errTrap
    If Len(Dir(msAppPath & "view_" & sProject & ".csv")) > 0 Then
        FileSystem.SetAttr msAppPath & "view_" & sProject & ".csv", vbNormal
        Kill msAppPath & "view_" & sProject & ".csv"
    End If

    If Len(Dir(msAppPath & sProject & ".csv")) > 0 Then
        FileCopy msAppPath & sProject & ".csv", msAppPath & "view_" & sProject & ".csv"
        If cmdStart.Enabled = False And (Len(cboProject.Text) > 0 And Len(txtJob.Text) > 0) Then
            AppendLineToProject GetLogLine, "view_" & sProject
        End If
    Else
        If cmdStart.Enabled = False And (Len(cboProject.Text) > 0 And Len(txtJob.Text) > 0) Then
            AppendLineToProject GetLogLine, "view_" & sProject
        Else
            AppendLineToProject "", "view_" & sProject
        End If
    End If
     
    FileSystem.SetAttr msAppPath & "view_" & sProject & ".csv", vbReadOnly
    
    ShellExecute Me.hwnd, "OPEN", msAppPath & "view_" & sProject & ".csv", 0, "", SW_SHOWNORMAL
Exit Sub
errTrap:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Form_MouseUp_Error
    If Button = vbRightButton Then
        PopupMenu mnuHidden
    End If
Exit Sub
Form_MouseUp_Error:
    LogError "Log in Sub Form_MouseUp of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - Form_MouseUp [" & Erl & "]", Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Form_QueryUnload_Error
    If cmdClose.Enabled = False Then
        MsgBox "Please Stop Timer before you close the program.", vbExclamation
        Cancel = True
    End If
Exit Sub
Form_QueryUnload_Error:
    LogError "Log in Sub Form_QueryUnload of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - Form_QueryUnload [" & Erl & "]", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    SaveSetting msAppName, "Formposition", "Left", CLng(Me.Left)
    SaveSetting msAppName, "Formposition", "Top", CLng(Me.Top)
    SaveSetting msAppName, "Formposition", "OnTop", chkOnTop.Value
End Sub

Private Sub lblStatus_DblClick()
Dim lTest As Long
On Error GoTo lblStatus_DblClick_Error
    lTest = 0 / 0
Exit Sub
lblStatus_DblClick_Error:
    LogError "Log in Sub lblStatus_DblClick of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - lblStatus_DblClick [" & Erl & "]", Err.Description
End Sub

Private Sub mnuClose_Click()
On Error GoTo mnuClose_Click_Error
    Unload Me
Exit Sub
mnuClose_Click_Error:
    LogError "Log in Sub mnuClose_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - mnuClose_Click [" & Erl & "]", Err.Description
End Sub

Private Sub mnuOpen_Click()
On Error GoTo mnuOpen_Click_Error
    'DoShowInTaskbar Me, True
    Timer1.Interval = 0
    Timer1.Enabled = False
    Me.Visible = True
    DoEvents
    'RemoveFromTray Me
    DoEvents
    Timer1.Interval = 1000
    Timer1.Enabled = True
Exit Sub
mnuOpen_Click_Error:
    LogError "Log in Sub mnuOpen_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - mnuOpen_Click [" & Erl & "]", Err.Description
End Sub

Private Sub mnuPause_Click()
On Error Resume Next
    If cmdPause.Enabled Then
        cmdPause_Click
    End If
End Sub

Private Sub mnuStart_Click()
On Error GoTo mnuStart_Click_Error
    If cmdStart.Enabled Then
        cmdStart_Click
    End If
Exit Sub
mnuStart_Click_Error:
    LogError "Log in Sub mnuStart_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - mnuStart_Click [" & Erl & "]", Err.Description
End Sub

Private Sub mnuStop_Click()
On Error GoTo mnuStop_Click_Error
    If cmdStop.Enabled Then
        cmdStop_Click
    End If
Exit Sub
mnuStop_Click_Error:
    LogError "Log in Sub mnuStop_Click of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - mnuStop_Click [" & Erl & "]", Err.Description
End Sub

Private Sub Timer1_Timer()
Dim dTime As Double
On Error GoTo Timer1_Timer_Error
    If Len(cboProject.Text) > 0 And Len(txtJob.Text) > 0 Then
        If cmdStart.Enabled = False Then
            If DateDiff("n", mdLastAlive, Now) > 15 Then
                mdPauseTime = mdLastAlive
                cmdStart.Enabled = True
                cmdPause.Enabled = False
                dTime = mdUsedPause + (Now - mdPauseTime)
                lblStatus.Caption = "Pause: " & CDate(mdStartTimer - mdUsedPause) & "         Used time: " & Format(mdPauseTime - mdStartTimer, "hh:nn:ss") & "      Paused: " & Format(dTime, "hh:nn:ss")
                cmdStart_Click
            Else
                WriteTmpFile
                If mdUsedPause > 0 Then
                    lblStatus.Caption = "Start: " & CDate(mdStartTimer - mdUsedPause) & "         Used time: " & PrintUsedTime2 & "      Paused: " & Format(mdUsedPause, "hh:nn:ss")
                Else
                    lblStatus.Caption = "Start: " & CDate(mdStartTimer) & "         Used time: " & PrintUsedTime2
                End If
                Me.Caption = PrintUsedTime2 & " " & cboProject.Text
            End If
            'Format(Now - mdStartTimer, "hh:nn:ss")
        ElseIf mdPauseTime > 0 Then
            dTime = mdUsedPause + (Now - mdPauseTime)
            lblStatus.Caption = "Pause: " & CDate(mdStartTimer - mdUsedPause) & "         Used time: " & Format(mdPauseTime - mdStartTimer, "hh:nn:ss") & "      Paused: " & Format(dTime, "hh:nn:ss")
        End If
    End If
    mnuStatus.Caption = "- " & lblStatus.Caption
    mdLastAlive = Now
Exit Sub
Timer1_Timer_Error:
    LogError "Log in Sub Timer1_Timer of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - Timer1_Timer [" & Erl & "]", Err.Description
End Sub

Private Sub txtJob_Change()
On Error GoTo txtJob_Change_Error
    If Len(cboProject.Text) > 0 And Len(txtJob.Text) > 0 Then
        cmdStart.Enabled = True
    Else
        cmdStart.Enabled = False
    End If
Exit Sub
txtJob_Change_Error:
    LogError "Log in Sub txtJob_Change of Form frmProjTimer"
    Resume Next ' cdlCancel
    'Err.Raise vbObjectError + Err.Number, "ProjectTimer - txtJob_Change [" & Erl & "]", Err.Description
End Sub

