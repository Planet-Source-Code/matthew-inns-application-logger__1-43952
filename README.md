<div align="center">

## Application Logger


</div>

### Description

This class provides a logging object that you can create and destroy as you like, the properties it has are:

CurrentProcedure

LogFile

LogLevel

MaxFileIterations

MaxFileSize

The methods are:

BeginLogging

StopLogging

IUPrint

If you create a global clsLogger object, in each procedure you can pass .CurrentProcedure the name of the procedure as a string. Anytime you want to log some data, use the IUPrint method to write out to the file. The logfile is kept open for the duration of your application to save time, the WriteFile API is used for this purpose also.

The IUPrint method takes the LogLevel of the message, the message itself and a parameter array of other data you'd like to write to the file. The levels work like this:

clslogger.LogLevel is set at 3, an IUPrint is processed passing a level of 5, IUPrint will not write it out to the file

clslogger.LogLevel is set at 3, an IUPrint is processed passing a level of 2, IUPrint will write it out to the file.

This way, you can up or down the level of logging to conserve log file size.

LogFile is the name of the file to be logged to. It will have a .Log extension, unless you make use of the File Iteration feature. File Iteration allows you to save to files until they reach a certain size(MaxFileSize), it will then change the extension to 002, 003 etc., until MaxFileIterations is met, then it will revert back to 001 again.

LogFile is the name of the initial logfile, without an extension.
 
### More Info
 
The properties of the class can be set in code but it is more flexible to use a database or registry in conjunction with the command line.

Here's an example of invoking the class:

Option Explicit

Dim Logger As New clsLogger

'our public object, if you want a global logger, add it to a module as a global object ie:

'Global Logger as New clsLogger

Private Sub Form_Load()

'set up the logger, you could use an ini file, or the command line in conjunction

'with either a database or registry, personally I use a database and the command line for loglevel

With Logger

.LogFile = "C:\zilpher"

.LogLevel = 2 'ignore inline messages of a level higher than 2

.MaxFileIterations = 3 'build max of 3 files

.MaxFileSize = 1200000 'max size of each log file in bytes

End With

If Logger.BeginLogging = False Then 'start the logger

Set Logger = Nothing 'kill it if it failed

MsgBox "Logging failed"

Exit Sub

End If

Logger.IUPrint 1, "Application started"

End Sub

Private Sub Form_Unload(Cancel As Integer)

Logger.CurrentProcedure = "Form_Unload"

Logger.IUPrint 1, "Exiting"

'stop the logger and clean up

Logger.StopLogging

Set Logger = Nothing

End Sub

Private Sub Form_Click()

Dim i As Long

'set the name of the current procedure - very important,

'get this wrong and it gets really confusing when reading the logs!

Logger.CurrentProcedure = "TestLogger"

'something to write out

'I've just used a loop to write some data out, not very realistic

'you'd more likely have lines written out after a dynamic SQL statement was built

'or any other point in your app

For i = 0 To 5000

'write out to the file, using i as an example of using the paramarray here

Logger.IUPrint 2, "Your message goes here", i, i - 1, i - 2, i - 3

Next i

'this message will not be written as it's level is higher than loglevel

Logger.IUPrint 3, "A message that doesn't get written"

End Sub


<span>             |<span>
---                |---
**Submitted On**   |2003-03-12 13:30:08
**By**             |[Matthew Inns](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-inns.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Applicatio1558263122003\.zip](https://github.com/Planet-Source-Code/matthew-inns-application-logger__1-43952/archive/master.zip)

### API Declarations

```
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal strFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
```





