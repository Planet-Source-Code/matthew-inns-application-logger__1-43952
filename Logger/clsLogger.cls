VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================================================
'=                                                                                        =
'=          Module: clsLogger                                                             =
'=            Type: Class Module                                                          =
'=                                                                                        =
'=       Developer: Matthew Inns                                                          =
'=            Date: 11-Mar-2003                                                           =
'=                                                                                        =
'=     Description: Contains the routines for managing log entries from the application.  =
'=                  The logging is done via the WriteFile API, the quickest method I have =
'=                  found so far; the file stays open unitl the StopLogging method is     =
'=                  called. www.zilpher.com for Excel version                             =
'=                                                                                        =
'=   Subprocedures: None                                                                  =
'=       Functions: AppendToLog - Adds the enumerated string to the end of the log file   =
'=                  IUPrint - The exposed logging method                                  =
'=                  BeginLogging - Opens the logfile for output, either creating new or   =
'=                  opening the old one for appending                                     =
'=                  StopLogging - Flushes file buffers and closes file                    =
'=                  GetNewFileName - Decides what to call the new log file                =
'=      Properties: CurrentProcedure - Stores the active procedure                        =
'=                  LogFile - The name of the current logfile                             =
'=                  LogLevel - The log level ceiling                                      =
'=                  MaxFileSize - The maximum size of the log file in bytes               =
'=                  MaxFileIterations - The number of log files to create before starting =
'=                  to overwrite them                                                     =
'=                  FileIteration - The current file number                               =
'=    DLL Declares: CreateFile - Alias CreateFileA                                        =
'=                  SetFilePointer                                                        =
'=                  CloseHandle                                                           =
'=                  WriteFile                                                             =
'=                  FlushFileBuffers                                                      =
'=                                                                                        =
'==========================================================================================

'superb documentation addin courtesy of Rob Bruce - www.analytical-dynamics.com

Option Explicit


'our API calls for handling the logfile
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal strFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

'some constants used in the above calls
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_NEW = 1
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_BEGIN = 0

'the handle of the logfile we'll be using
Dim lngLogFile As Long

'local variables for the properties
Dim pvtCurrentProcedure As String
Dim pvtLogFile As String
Dim pvtLogLevel As Byte
Dim pvtMaxFileSize As Long
Dim pvtMaxFileIterations As Long
Dim pvtFileIteration As Long

Private Function AppendToLog(ByVal strMsg As String) As Boolean
'===================================================================
'=   Procedure: AppendToLog                                        =
'=        Type: Private Function                                   =
'=                                                                 =
'=     Purpose: Appends the strMsg to the LogFile via the          =
'=              WriteFile function of Kernel32                     =
'=  Parameters: ByVal strMsg - String -                            =
'=     Returns: Boolean -                                          =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |07-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Dim lngTemp As Long
Dim bytMsgArray() As Byte

'WriteFile requires a byte array so we need to convert the strMsg string
ReDim bytMsgArray(0 To Len(strMsg) - 1) As Byte

For lngTemp = 1 To Len(strMsg)
    bytMsgArray(lngTemp - 1) = Asc(Mid$(strMsg, lngTemp, 1))
Next

'if the adding of this string will tke us over our permitted file size limit we should get a new file
If FileLen(LogFile) + UBound(bytMsgArray) > MaxFileSize And MaxFileSize > 0 Then
    StopLogging 'clean out the pipes..
    GetNewFileName 'get a new name or delete current log file if no iterations are set
    BeginLogging 'start logging again
End If


'write the string to the log file
If (WriteFile(lngLogFile, bytMsgArray(0), Len(strMsg), lngTemp, ByVal 0&) <> 0) Then
    'the data was written correctly
    AppendToLog = True
End If

End Function

Function IUPrint(Level As Byte, Msg As String, ParamArray Variables() As Variant) As Boolean
'===================================================================
'=   Procedure: IUPrint                                            =
'=        Type: Function                                           =
'=                                                                 =
'=     Purpose: The method used to write a log entry               =
'=  Parameters: Level - Byte - The logging level trigger for the   =
'=              message  Msg - String - The message to log         =
'=              ParamArray Variables() - Variant - anything else   =
'=              you want to log                                    =
'=     Returns: Boolean -                                          =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |07-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Dim i As Long

'compare logging level to Level and bail if neccessary
'when IUPrint is called the message is given a default level to log at
'if loglevel is set to 5 then all messages passed with a level of 5 or less
'will be logged, where as a message with a level of 6 or more won't
If LogLevel < Level Then Exit Function

'we should be logging so set up the msg to append
'the message will end up as:
'20030307112254;ProcedureName;Message defined by you;any;variables;passed;
Msg = CurrentProcedure & ";" & Msg & ";" 'add the procedure name to the string

'get all the paramarray into the string too
For i = 0 To UBound(Variables)
    Msg = Msg & Variables(i) & ";"
Next i

'and finally put a date/time stamp onto it and a vbcrlf so we get lines...
Msg = Format(Now(), "yyyymmddhhmmss") & Chr(58) & Msg & vbCrLf

'now off to the log
AppendToLog Msg

End Function

Function BeginLogging() As Boolean
'===================================================================
'=   Procedure: BeginLogging                                       =
'=        Type: Function                                           =
'=                                                                 =
'=     Purpose: Creates or opens the log file as necessary, leaves =
'=              a public handle for the other API functions        =
'=  Parameters: None                                               =
'=     Returns: Boolean -                                          =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |07-Mar-03| Matthew Inns  | Created                     =
'===================================================================

Dim lngFileLen As Long

lngFileLen = 0


'must have a logfile name
If Len(LogFile) = 0 Then
    BeginLogging = False
    Exit Function
End If

'an idea for you if you wish to automate the log file name
'If Len(LogFile) = 0 Then
'    If MaxFileIterations = 0 Then
'        LogFile = App.Path & "\" & Left(App.ProductName, InStr(1, App.ProductName, ".") - 1) & ".log"
'    Else
'        LogFile = App.Path & "\" & Left(App.ProductName, InStr(1, App.ProductName, ".") - 1) & ".001"
'    End If
'End If

'sort the extension
If InStr(1, LogFile, ".") = 0 Then
'as we are not using iterations to keep historical logfiles we can just use the .log extension
    If MaxFileIterations = 0 Then
        LogFile = LogFile & ".log"
    Else
'we are using iterations to keep historical logfiles so we need to work out which iteration to use
        GetNewFileName
    End If
End If

'if the log file already exists, we need to append to the end of it, so find out how long
'it is in bytes first
If (Len(Dir(LogFile))) Then lngFileLen = FileLen(LogFile)

'open the log file or create it if necessary
lngLogFile = CreateFile(LogFile, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, _
    IIf(Len(Dir(LogFile)), OPEN_EXISTING, CREATE_NEW), _
    FILE_ATTRIBUTE_NORMAL, 0&)

'if the logfile failed to open there's no need to carry on
If (lngLogFile = INVALID_HANDLE_VALUE) Then
    BeginLogging = False
    Exit Function
End If

'put the file pointer at the end of file (if we didn't create it)
'I prefer to write to the end of the file, if you prefer to write at the beginning
'skip this and be sure to send a vbcrlf first to create the line - note, I haven't tested that
If (lngFileLen <> 0) Then
    If (SetFilePointer(lngLogFile, lngFileLen, ByVal 0&, FILE_BEGIN) = &HFFFFFFFF) Then
        'exit sub if the pointer did not set correctly
        BeginLogging = False
        CloseHandle (lngLogFile)
        Exit Function
    End If
End If

BeginLogging = True

End Function

Function StopLogging() As Boolean
'===================================================================
'=   Procedure: StopLogging                                        =
'=        Type: Function                                           =
'=                                                                 =
'=     Purpose: Cleans up and closes the logfile                   =
'=  Parameters: None                                               =
'=     Returns: Boolean -                                          =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
FlushFileBuffers (lngLogFile)
CloseHandle (lngLogFile)
End Function

'===================================================================
'=   Procedure: CurrentProcedure                                   =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: Each procedure will set this property allowing the =
'=              log to state what procedure logged what            =
'=  Parameters: None                                               =
'=     Returns: String -                                           =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Public Property Get CurrentProcedure() As String
    CurrentProcedure = pvtCurrentProcedure
End Property
Public Property Let CurrentProcedure(ByVal CurrentProcedure As String)
    pvtCurrentProcedure = CurrentProcedure
End Property

'===================================================================
'=   Procedure: LogFile                                            =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: The log file we are working with                   =
'=  Parameters: None                                               =
'=     Returns: String -                                           =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Public Property Get LogFile() As String
    LogFile = pvtLogFile
End Property
Public Property Let LogFile(ByVal LogFile As String)
    pvtLogFile = LogFile
End Property

'===================================================================
'=   Procedure: LogLevel                                           =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: The level to log at, any IUPrint methods called    =
'=              with a log level higher than this will get ignored =
'=  Parameters: None                                               =
'=     Returns: Long -                                             =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Public Property Get LogLevel() As Long
    LogLevel = pvtLogLevel
End Property
Public Property Let LogLevel(ByVal LogLevel As Long)
    pvtLogLevel = LogLevel
End Property

'===================================================================
'=   Procedure: MaxFileSize                                        =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: The maximum size of the logfile before it begins a =
'=              new iteration                                      =
'=  Parameters: None                                               =
'=     Returns: Long -                                             =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Public Property Get MaxFileSize() As Long
    MaxFileSize = pvtMaxFileSize
End Property
Public Property Let MaxFileSize(ByVal MaxFileSize As Long)
    pvtMaxFileSize = MaxFileSize
End Property

'===================================================================
'=   Procedure: MaxFileIterations                                  =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: If this is > 0 log files will get a 001 extension, =
'=              incrementing until MaxFileIterations, then they    =
'=              will restart at 001 again                          =
'=  Parameters: None                                               =
'=     Returns: Long -                                             =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Public Property Get MaxFileIterations() As Long
    MaxFileIterations = pvtMaxFileIterations
End Property
Public Property Let MaxFileIterations(ByVal MaxFileIterations As Long)
    pvtMaxFileIterations = MaxFileIterations
End Property

'===================================================================
'=   Procedure: FileIteration                                      =
'=        Type: Property                                           =
'=                                                                 =
'=     Purpose: The current iteration of the file                  =
'=  Parameters: None                                               =
'=     Returns: Long -                                             =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================
Private Property Get FileIteration() As Long
    FileIteration = pvtFileIteration
End Property
Private Property Let FileIteration(ByVal FileIteration As Long)
    pvtFileIteration = FileIteration
End Property

Private Function GetNewFileName() As Boolean
'===================================================================
'=   Procedure: GetNewFileName                                     =
'=        Type: Function                                           =
'=                                                                 =
'=     Purpose: Bumps up the file name if required                 =
'=  Parameters: None                                               =
'=     Returns: Boolean -                                          =
'=                                                                 =
'= Version:  Date:     Developer:      Action:                     =
'=---------|---------|---------------|-----------------------------=
'=  1.0.0  |10-Mar-03| Matthew Inns  | Created                     =
'===================================================================

'if we are using iterations and we aren't already at the limit, increment by one
If MaxFileIterations > 0 And Not FileIteration = MaxFileIterations Then
    FileIteration = FileIteration + 1
    
'if we are using iterations and we are already at the limit, reset to one
ElseIf MaxFileIterations > 0 And FileIteration = MaxFileIterations Then
    FileIteration = 1

'otherwise we aren't using iterations so it's a .log name meaning we can
'kill the file and get on with creating a newone
Else
    'delete file and exit
    On Error Resume Next
    Kill LogFile
    On Error GoTo 0
    GetNewFileName = True
    Exit Function
End If

'now sort the name out
If InStr(1, LogFile, ".") = 0 Then
    LogFile = LogFile & "." & Format(FileIteration, "000")
Else
    LogFile = Left(LogFile, InStr(1, LogFile, ".")) & Format(FileIteration, "000")
End If
GetNewFileName = True

End Function

