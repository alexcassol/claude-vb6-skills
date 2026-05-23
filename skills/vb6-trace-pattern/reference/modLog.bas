Attribute VB_Name = "modLog"
'--------------------------------------------------------------------------------
'    Component  : modLog
'    Project    : <YourProject>
'
'    Description: Manual stack trace and logging for Visual Basic 6.
'
'                 Procedures call EnterMethod on entry and ExitMethod on exit
'                 to maintain a manual call stack in a module-level array.
'                 LogMsg writes timestamped entries to a daily rotating file
'                 under <App.Path>\log\. TraceMethod returns the current stack
'                 for inclusion in error messages.
'
'                 Dependencies:
'                   - Microsoft Scripting Runtime (Dictionary, FileSystemObject)
'                   - A fg_InIDE() As Boolean function (returns True when
'                     running in the VB6 IDE)
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

' Internal call stack
Private mstrMethods() As String

' Public configuration
Public gstrLogBasePath As String
Public gblnRegisterLog As Boolean

' Log retention in days
Private Const DEF_LOG_DAYS As Long = 30


'---------------------------------------------------------------------------------------
' Procedure : RegisterLog
' Purpose   : Initializes logging on first use. Creates the log directory and
'             decides whether logging is active for this session. Called
'             internally by LogMsg.
'             Activation signals:
'                 1. fg_InIDE() returns True (always log in dev)
'                 2. debug.txt exists in App.Path (sysadmin toggle)
'                 3. gblnRegisterLog already set True by application code
'---------------------------------------------------------------------------------------
Private Sub RegisterLog()
    Static blnVerified As Boolean

    If blnVerified Then Exit Sub

    gstrLogBasePath = App.Path & "\log\"
    If Dir(gstrLogBasePath, vbDirectory) = "" Then
        MkDir gstrLogBasePath
    End If

    If Dir(App.Path & "\debug.txt", vbArchive) <> "" Or fg_InIDE() Or gblnRegisterLog Then
        gblnRegisterLog = True
    Else
        gblnRegisterLog = False
    End If

    blnVerified = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogMsg
' Purpose   : Writes a single timestamped line to the daily log file.
'             If Force = True, writes regardless of gblnRegisterLog state.
'             Also triggers DumpProcesses when the message contains
'             "failed" or "automation error".
'---------------------------------------------------------------------------------------
Public Sub LogMsg(ByVal msg As String, Optional ByVal Force As Boolean = False)
    Dim intFile As Integer
    Static blnCleaned As Boolean
    Static blnDumping As Boolean

    On Error GoTo ErrHandler

    RegisterLog

    If Not gblnRegisterLog And Not Force Then Exit Sub
    If gstrLogBasePath = "" Then Exit Sub

    If Not blnCleaned Then
        CleanupOldLogs gstrLogBasePath
        blnCleaned = True
    End If

    intFile = FreeFile
    Open gstrLogBasePath & "app-" & Format(Date, "yyyy-mm-dd") & ".txt" For Append As #intFile

    Print #intFile, Format(Now, "hh:nn:ss") & " < " & msg & " >"

ErrHandler:
    On Error Resume Next
    Close #intFile

    ' Auto-dump processes on automation errors to aid diagnostics
    If Not blnDumping Then
        If InStr(1, msg, "failed", vbTextCompare) > 0 Or _
           InStr(1, msg, "automation error", vbTextCompare) > 0 Then
            blnDumping = True
            DumpProcesses
            blnDumping = False
        End If
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnterMethod
' Purpose   : Pushes a method onto the manual call stack and logs the entry.
'             Optional parameter values are serialized into the log message.
'---------------------------------------------------------------------------------------
Public Sub EnterMethod(ByVal Obj As String, ByVal Method As String, ParamArray Parameters())
    Dim i As Long
    Dim strParams As String

    On Error Resume Next

    ReDim Preserve mstrMethods(UBound(mstrMethods) + 1)
    If Err.Number <> 0 Then
        ReDim mstrMethods(0)
        ReDim Preserve mstrMethods(UBound(mstrMethods) + 1)
    End If
    Err.Clear
    On Error GoTo 0

    mstrMethods(UBound(mstrMethods)) = Obj & "::" & Method

    ' Serialize parameters if provided
    On Error GoTo SkipParams

    strParams = ""
    For i = 0 To UBound(Parameters)
        If IsObject(Parameters(i)) Then
            If Parameters(i) Is Nothing Then
                strParams = strParams & "Nothing"
            Else
                strParams = strParams & "Object:" & TypeName(Parameters(i))
            End If
        ElseIf IsEmpty(Parameters(i)) Then
            strParams = strParams & "Empty"
        ElseIf IsNull(Parameters(i)) Then
            strParams = strParams & "Null"
        ElseIf IsArray(Parameters(i)) Then
            strParams = strParams & "Array[" & (UBound(Parameters(i)) - LBound(Parameters(i)) + 1) & "]"
        Else
            strParams = strParams & CStr(Parameters(i))
        End If

        If i < UBound(Parameters) Then
            strParams = strParams & ", "
        End If
    Next i

    If strParams <> "" Then
        mstrMethods(UBound(mstrMethods)) = mstrMethods(UBound(mstrMethods)) & "( " & strParams & " )"
    End If

    Err.Clear
    On Error GoTo 0

SkipParams:
    LogMsg "Enter: " & mstrMethods(UBound(mstrMethods))
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExitMethod
' Purpose   : Pops a method from the manual call stack and logs the exit.
'             Two modes:
'               - Specific (Obj + Method): pops one frame matching that pair
'               - Bulk (Obj only):         pops every top frame whose object
'                                          matches (safety-net for missed exits)
'             Prefer the specific form. The bulk form indicates a missed
'             ExitMethod somewhere in the call chain.
'---------------------------------------------------------------------------------------
Public Sub ExitMethod(ByVal Obj As String, Optional ByVal Method As String = "")
    Dim i As Long

    On Error GoTo ErrHandler

    If Method = "" Then
        ' Bulk mode: pop everything at the top whose object matches
        For i = UBound(mstrMethods) To 0 Step -1
            If Left$(mstrMethods(i), Len(Obj & "::")) = Obj & "::" Then
                LogMsg "Exit: " & mstrMethods(i)
                ReDim Preserve mstrMethods(UBound(mstrMethods) - 1)
            Else
                Exit For
            End If
        Next i
    Else
        ' Specific mode: pop one frame
        If UBound(mstrMethods) > 0 Then
            ReDim Preserve mstrMethods(UBound(mstrMethods) - 1)
        End If
        LogMsg "Exit: " & Obj & "::" & Method
    End If

    Exit Sub

ErrHandler:
    LogMsg "Error in ExitMethod: " & Err.Description & " (" & Obj & "::" & Method & ")"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TraceMethod
' Purpose   : Returns the current call stack as a formatted string, suitable
'             for inclusion in error messages.
'---------------------------------------------------------------------------------------
Public Function TraceMethod() As String
    On Error Resume Next
    TraceMethod = "Trace: " & Join(mstrMethods, vbCrLf & "   ")
End Function


'---------------------------------------------------------------------------------------
' Procedure : SerializeMethod
' Purpose   : Standalone serializer for ad-hoc logging of a parameter list.
'             Useful in error handlers for capturing arguments without
'             instrumenting the procedure with EnterMethod.
'---------------------------------------------------------------------------------------
Public Function SerializeMethod(ParamArray Parameters()) As String
    Dim i As Long

    On Error Resume Next

    For i = 0 To UBound(Parameters)
        SerializeMethod = SerializeMethod & CStr(Parameters(i)) & "|"
    Next i

    LogMsg "Serialize: " & SerializeMethod
End Function


'---------------------------------------------------------------------------------------
' Procedure : CleanupOldLogs
' Purpose   : Deletes log files older than DEF_LOG_DAYS in the given directory.
'             Called once per session on first log entry.
'---------------------------------------------------------------------------------------
Private Sub CleanupOldLogs(ByVal sDir As String)
    Dim objFso As New FileSystemObject
    Dim objFile As File

    On Error GoTo ErrHandler

    For Each objFile In objFso.GetFolder(sDir).Files
        If DateAdd("d", -DEF_LOG_DAYS, Date) > objFile.DateLastModified Then
            objFile.Delete True
        End If
    Next

ErrHandler:
    ' Best-effort cleanup; ignore errors
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DumpProcesses
' Purpose   : Lists every running process on the machine via WMI and writes
'             the dump to the log. Invoked automatically by LogMsg when an
'             automation-error pattern is detected. Useful for diagnosing
'             COM conflicts, antivirus interference, or stale app instances.
'---------------------------------------------------------------------------------------
Public Sub DumpProcesses()
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim strLog As String

    On Error Resume Next

    strComputer = "."

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Process", , 48)

    For Each objItem In colItems
        strLog = ""
        strLog = strLog & vbNewLine & "********************************************************************"
        strLog = strLog & vbNewLine & "Caption: " & objItem.Caption
        strLog = strLog & vbNewLine & "CommandLine: " & objItem.CommandLine
        strLog = strLog & vbNewLine & "ExecutablePath: " & objItem.ExecutablePath
        strLog = strLog & vbNewLine & "ProcessId: " & objItem.ProcessId
        strLog = strLog & vbNewLine & "ParentProcessId: " & objItem.ParentProcessId
        strLog = strLog & vbNewLine & "CreationDate: " & objItem.CreationDate
        strLog = strLog & vbNewLine & "ThreadCount: " & objItem.ThreadCount
        strLog = strLog & vbNewLine & "HandleCount: " & objItem.HandleCount
        strLog = strLog & vbNewLine & "WorkingSetSize: " & objItem.WorkingSetSize
        strLog = strLog & vbNewLine & "PageFileUsage: " & objItem.PageFileUsage
        strLog = strLog & vbNewLine & "VirtualSize: " & objItem.VirtualSize
        strLog = strLog & vbNewLine & "Priority: " & objItem.Priority
        strLog = strLog & vbNewLine & "SessionId: " & objItem.SessionId
        strLog = strLog & vbNewLine & "Status: " & objItem.Status
        strLog = strLog & vbNewLine & "********************************************************************"
        strLog = strLog & vbNewLine

        LogMsg strLog, True
    Next

    Err.Clear
End Sub
