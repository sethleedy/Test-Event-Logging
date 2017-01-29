
Option Explicit On
'Option Strict On ' readFromLog sub is using late binding

Imports System
Imports System.Diagnostics ' Allows access to the EventLog Object
Imports System.Threading

Module Module1
    Dim myLog As New EventLog
    Public newData As String = ""

    ' Run this once before using the additional code. Creates the objects to use
    Public Sub startUp()

        ' Create a new log.
        '
        ' A Security Exception will fire if I'm not running with administrator privileges when checking EventLog.SourceExists("MyNewSource")
        ' This is why the program is set to elevate when executed, http://stackoverflow.com/questions/3080725/debug-a-program-that-needs-administrator-rights-under-windows-7
        ' https://msdn.microsoft.com/en-us/library/6s7642se(v=vs.110).aspx
        If Not EventLog.SourceExists("MyNewSource") Then
            EventLog.CreateEventSource("MyNewSource", "MyNewLog")
        End If

        ' Set a few options
        With myLog
            .Source = "MyNewSource" ' Look for this name within the Windows Event Logs. In Windows 10 it was located under Applications and Services Logs
            .Log = "MyNewLog"
            .EnableRaisingEvents = True
        End With

        ' Add an event to fire when an Event Eentry is written. Shows up in Form2, Form3 - textbox1
        AddHandler myLog.EntryWritten, AddressOf OnEntryWritten


    End Sub

    ' Purpose of this module, to write something to the event logs
    Public Sub writeToLog(msg As String, Optional ByVal errorType As String = "information") ' https://msdn.microsoft.com/en-us/library/system.diagnostics.eventlogentrytype(v=vs.110).aspx

        ' Second parameter options:
        '   error, warning, information, failureaudit, successaudit
        ' default: information

        If msg <> "" Then
            Select Case errorType
                Case "information"
                    myLog.WriteEntry(msg, EventLogEntryType.Information)
                Case "error"
                    myLog.WriteEntry(msg, EventLogEntryType.Error)
                Case "warning"
                    myLog.WriteEntry(msg, EventLogEntryType.Warning)
                Case "failureaudit"
                    myLog.WriteEntry(msg, EventLogEntryType.FailureAudit)
                Case "successaudit"
                    myLog.WriteEntry(msg, EventLogEntryType.SuccessAudit)
            End Select
        End If
    End Sub

    Public Sub readFromLog()

        ' Dim List to hold the data and then we can reverse the list so the newest events are at the top
        Dim dataList As New List(Of String)
        Dim tmpStr As String

        For Each entry In myLog.Entries
            With entry
                tmpStr = entry.TimeWritten + ": " + entry.Message + " - " + entry.Source + vbCr
                dataList.Add(tmpStr)
            End With
        Next

        Form3.TextBox1.Text = String.Join(vbCrLf, dataList.ToArray.Reverse)

    End Sub

    ' Handy function to display the log entry
    ' Fired from the startUp sub with the "AddHandler myLog.EntryWritten, AddressOf OnEntryWritten"
    Public Sub OnEntryWritten(ByVal [source] As Object, ByVal e As EntryWrittenEventArgs)

        ' This collects the data into a variable for pickup by the Form1 timer function
        ' This is needed because the AddHandler makes another thread which cannot change the UI thread directly.
        ' This allows an update to the UI after the background thread is finished.

        ' Add a date/time stamp to make it wasy to read on the textbox tack on the error message that was passed to Windows Event Logging
        newData = DateTime.Now.ToString("F") + ": " + e.Entry.Message

    End Sub

    ' This will clear and then delete the log we created.
    Public Sub deleteLog()

        ' Delete your new log.

        ' First check if it still exists
        If EventLog.SourceExists("MyNewSource") Then
            myLog.Clear()
            EventLog.Delete("MyNewLog") ' This is the recommendation of the VB.Net IDE. Different from referring article - myLog.Delete("MyNewLog")
        End If
    End Sub

End Module
