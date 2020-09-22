Attribute VB_Name = "modLog"
Option Explicit

Public msAppPath As String

Public Sub LogError(ByVal sText As String)
Dim lFile As Long
Dim sErrDescription As String
Dim lErrNo As Long
Dim sErrSource As String
Dim sErrLastDll As String
Dim lErl As Long
Dim bNewFile As Boolean
Dim sLine As String
Dim sAppName As String
    sErrDescription = Err.Description
    lErrNo = Err.Number
    sErrSource = Err.Source
    sErrLastDll = Err.LastDllError
    lErl = Erl
On Error GoTo errTrap
    sAppName = App.EXEName & ".log"
    If Len(Dir(msAppPath & sAppName & ".csv")) > 0 Then
        bNewFile = False
        If FileSystem.FileLen(msAppPath & sAppName & ".csv") > 1000000 Then
            If Len(Dir(msAppPath & sAppName & ".2.csv")) > 0 Then
                Kill msAppPath & sAppName & ".2.csv"
            End If
            FileSystem.FileCopy msAppPath & "Log." & sAppName & ".csv", msAppPath & sAppName & ".2.csv"
            bNewFile = True
        End If
    Else
        bNewFile = True
    End If
    lFile = FreeFile
    If bNewFile Then
        Open msAppPath & sAppName & ".csv" For Output As #lFile
        sLine = """Date"""
        sLine = sLine & ";"
        sLine = sLine & """Time"""
        sLine = sLine & ";"
        sLine = sLine & """Note"""
        sLine = sLine & ";"
        sLine = sLine & """ErrDescription"""
        sLine = sLine & ";"
        sLine = sLine & """ErrNumber"""
        sLine = sLine & ";"
        sLine = sLine & """Line"""
        sLine = sLine & ";"
        sLine = sLine & """Last DLL error"""
        sLine = sLine & ";"
        sLine = sLine & """Program"""
        Print #lFile, sLine
    Else
        Open msAppPath & sAppName & ".csv" For Append As #lFile
    End If
    sLine = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    sLine = sLine & ";"
    sLine = sLine & Format(Now, "hh:nn:ss")
    sLine = sLine & ";"
    sLine = sLine & """" & Replace(sText, """", "''") & """"
    sLine = sLine & ";"
    sLine = sLine & """" & Replace(sErrDescription, """", "''") & """"
    sLine = sLine & ";"
    sLine = sLine & lErrNo
    sLine = sLine & ";"
    sLine = sLine & lErl
    sLine = sLine & ";"
    sLine = sLine & """" & Replace(sErrLastDll, """", "''") & """"
    sLine = sLine & ";"
    sLine = sLine & sErrSource
    Print #lFile, sLine
    Debug.Print vbCrLf & sErrDescription & vbCrLf & sLine
    Debug.Assert False
    Close lFile
    Beep
Exit Sub
errTrap:
    Debug.Assert False
    Debug.Print Err.Description
    Resume Next
End Sub

