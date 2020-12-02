Attribute VB_Name = "mLog"
Option Explicit
Option Private Module

'Use LogEvent to determine which events are written to the log
Public Enum LogEvent
    LogEventOpen = 1
    LogEventClose
    LogEventGTe
    LogEventGTw
End Enum
Public Enum GetTextFileFor
    GetTextFileForInput = 1
    GetTextFileForAppend
    GetTextFileForOutput
End Enum

Function fGetLogFile(ByRef iGetFor As GetTextFileFor) As Integer
    Dim sFilename As String
    Dim iFile As Integer
    Dim sPS As String
    
    sPS = Application.PathSeparator
    sFilename = ThisWorkbook.Path & sPS & "Google Trends Information Extraction Tool.log"
    iFile = FreeFile
    
    On Error Resume Next
    If iGetFor = GetTextFileForInput Then
        Open sFilename For Input As iFile
    ElseIf iGetFor = GetTextFileForAppend Then
        Open sFilename For Append As iFile
    ElseIf iGetFor = GetTextFileForOutput Then
        Open sFilename For Output As iFile
    End If
    If Err.Number = 0 Then
        fGetLogFile = iFile
    Else
        If bShowCompletionMsgBoxes Then
            fGetLogFile = 0
            With Err
                If .Number = 70 Then
                'The file is already opened by another process
                    MsgBox "The log file could not be accessed because it is open in another application.", vbInformation + vbOKOnly, "Log file reserved by another application"
                ElseIf .Number = 75 Then
                'the file cannot be created (path/file) access error
                    MsgBox "A file or path error occurred in accessing the log file." & vbCrLf & "No log entry will be written.", vbInformation + vbOKOnly, "Log File/Path access error"
                ElseIf .Number = 55 Then
                'The file is already open
                    MsgBox "The log file was opened, but not closed successfully, and cannot be accessed." & vbCrLf & "No log entry will be written.", vbInformation + vbOKOnly, "Log file already open"
                End If
            End With
        End If
        Err.Clear
    End If
End Function
 
Sub UpdateLogEntry(ByVal sLogEntryOld As String, ByVal sLogEntryNew As String)
    Dim iFile As Integer
    Dim sFileText As String
    
    'Get the file for input
    iFile = fGetLogFile(GetTextFileForInput)
    
    If iFile = 0 Then Exit Sub  'If no file is returned, then there is no sense continuing the write

    sFileText = Input(LOF(iFile), iFile)
    
    Close #iFile
    
    sFileText = Replace(sFileText, sLogEntryOld, sLogEntryNew, , , vbTextCompare)
    
    'Get the file again, this time for output
    iFile = fGetLogFile(GetTextFileForOutput)
    
    Print #iFile, sFileText
    
    Close #iFile
    
End Sub
 
Sub WriteLogEntry(ByVal iLogEvent As LogEvent _
                , Optional ByVal sLogEntry As String)
    Dim iFile As Integer
    'Dim sLogEntry As String
    
    'Get the file
    iFile = fGetLogFile(GetTextFileForAppend)
    
    If iFile = 0 Then Exit Sub  'If no file is returned, then there is no sense continuing the write
    
    'If a file has been returned, determine what to write
    Select Case iLogEvent
    Case LogEventOpen
        sLogEntry = fCreateOpenCloseString("open")
    Case LogEventClose
        sLogEntry = fCreateOpenCloseString("close")
'    Case LogEventGTeStart
'        sLogEntry = fCreateGTeString
'    Case LogEventGTeEnd
'
'    Case LogEventGTwStart
'        sLogEntry = fCreateGTwString
'    Case LogEventGTwEnd
'
    End Select
    
    'Once the write string has been created, write it
    Print #iFile, sLogEntry
    Close #iFile
End Sub
Private Function fCreateOpenCloseString(sOpenClose As String) As String
    On Error Resume Next
    fCreateOpenCloseString = "Application (" & fCurrentVersionNumber & ") " & Left(sOpenClose, 4) & "ed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    If Err.Number <> 0 Then
        Err.Clear
        fCreateOpenCloseString = vbNullString
    End If
End Function
'
'Private Function fCreateGTeStartString() As String
'
'End Function
'
'Private Function fCreateGTeEndString() As String
'
'End Function
'Private Function fCreateGTwStartString() As String
'
'End Function
'
'Private Function fCreateGTwEndString() As String
'
'End Function
'
