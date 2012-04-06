'\\ --[PrintJobEventArgs]---------------------------------------------
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ------------------------------------------------------------------

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintJobEventArgs
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class wrapper for the event arguments used in the change events 
''' relating to individual jobs in a printer queue..
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Prints the user name of a job when it is added 
''' <code>
'''   Private WithEvents pmon As New PrinterMonitorComponent
'''
'''   pmon.AddPrinter("Microsoft Office Document Image Writer")
'''   pmon.AddPrinter("HP Laserjet 5")
''' 
'''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
'''
'''    With CType(e, PrintJobEventArgs)
'''        Trace.WriteLine( .PrintJob.Username )
'''     End With
'''
'''  End Sub
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class PrintJobEventArgs
    Inherits EventArgs
    Implements IDisposable

#Region "Private member variables"
    Private _Job As PrintJob
    Private _EventTime As DateTime
    Private _EventType As PrintJobEventTypes
#End Region

#Region "Public enumerated types"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Enum PrintJobEventTypes
        JobAddedEvent = 1
        JobDeletedEvent = 2
        JobSetEvent = 3
        JobWrittenEvent = 4
    End Enum

#End Region

#Region "Public Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The print job for which this event occured
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PrintJob() As PrintJob
        Get
            Return _Job
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The date and time at which the print job event occured
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property EventTime() As DateTime
        Get
            Return _EventTime
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of job event that occured
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property EventType() As PrintJobEventTypes
        Get
            Return _EventType
        End Get
    End Property
#End Region

#Region "Public Constructors"
    Friend Sub New(ByVal Job As PrintJob, ByVal Time As DateTime, ByVal EventType As PrintJobEventTypes)
        _Job = Job
        _EventTime = Time
        _EventType = EventType
    End Sub

    Friend Sub New(ByVal Job As PrintJob, ByVal EventType As PrintJobEventTypes)
        Me.New(Job, DateTime.Now, EventType)
    End Sub
#End Region

#Region "Public Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Frees up any system resources used by this job event notification
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Overloads Sub Dispose() Implements IDisposable.Dispose
        _Job.Dispose()
    End Sub

    Public Overloads Function Equals(ByVal PrintJobEventArgs As PrintJobEventArgs) As Boolean
        If Not PrintJobEventArgs Is Nothing Then
            If PrintJobEventArgs.EventType = _EventType AndAlso PrintJobEventArgs.PrintJob.JobId = _Job.JobId AndAlso PrintJobEventArgs.PrintJob.PrinterName = _Job.PrinterName Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

#End Region

End Class
