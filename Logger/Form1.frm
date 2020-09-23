VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
