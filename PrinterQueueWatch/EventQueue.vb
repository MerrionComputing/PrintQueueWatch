'\\ --[EventQueue]-----------------------------------------------------
'\\ Queues printer events and printjob events to be 
'\\ raised asynchronously thus allowing the printer monitor
'\\ thread to get on with its business while existing events are 
'\\ being processed.
'\\ (c) 2003 Merrion Computing Ltd
'\\ -------------------------------------------------------------------
Imports System.Threading

Friend Class EventQueue
    Implements IDisposable

#Region "PrinterEventQueue class"
    Private Class PrinterEventQueue
        Inherits Collections.Concurrent.ConcurrentQueue(Of EventArgs)

        Public Overloads Function Contains(ByVal JobEventArgs As PrintJobEventArgs) As Boolean

            '\\ Duplicate JobWritten events can be ignored
            If JobEventArgs.EventType = PrintJobEventArgs.PrintJobEventTypes.JobWrittenEvent Then
                Exit Function
            End If

            '\\ Because an EventQueue pertains to only one printer we can compare on job id and event type
            Dim oItem As Object
            Try
                For Each oItem In Me
                    If TypeOf oItem Is PrintJobEventArgs Then
                        With DirectCast(oItem, PrintJobEventArgs)
                            If JobEventArgs.EventType = .EventType AndAlso JobEventArgs.PrintJob.JobId = .PrintJob.JobId AndAlso JobEventArgs.PrintJob.QueuedTime = .PrintJob.QueuedTime Then
                                Return True
                                Exit Function
                            End If
                        End With
                    End If
                Next
            Catch es As Exception
                Return False
                Exit Function
            End Try
        End Function

    End Class
#End Region

#Region "Private member variables"

    Private _EventQueue As New PrinterEventQueue()
    Private _EventQueueWorker As Thread

    Private _JobEvent As PrinterMonitorComponent.JobEvent
    Private _PrinterEvent As PrinterMonitorComponent.PrinterEvent

    Private _WaitHandle As AutoResetEvent
    Private Shared _Cancelled As Boolean


#End Region

#Region "Public interface"
#Region "AddJobEvent"
    Public Sub AddJobEvent(ByVal JobEventArgs As PrintJobEventArgs)
        '\\ Job events must be unique
        If Not _EventQueue.Contains(JobEventArgs) Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Job event enqueued : " & JobEventArgs.EventType.ToString, Me.GetType.ToString)
            End If
            '\\ If the job didn't populate properly, try again
            With JobEventArgs.PrintJob
                If Not .Populated Then
                    Call .Refresh()
                End If
            End With
            _EventQueue.Enqueue(JobEventArgs)
        End If
    End Sub
#End Region

#Region "AddPrinterEvent"
    Public Sub AddPrinterEvent(ByVal PrinterEventArgs As PrinterEventArgs)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Printer event enqueued", Me.GetType.ToString)
        End If
        _EventQueue.Enqueue(PrinterEventArgs)
    End Sub
#End Region

#Region "Awaken"
    '\\ --[Awaken]--------------------------------------------------------
    '\\ Wakes the thread which dequeues and processes the events
    '\\ ------------------------------------------------------------------
    Public Sub Awaken()
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Event queue awakened - " & _EventQueue.Count.ToString & " events ", Me.GetType.ToString)
        End If
        If Not _WaitHandle Is Nothing AndAlso EventsPending Then
            _WaitHandle.Set()
        End If
    End Sub
#End Region

#Region "EventsPending"
    ''' <summary>
    ''' Returns True if there are any printer or print job events queued that should be processed 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EventsPending() As Boolean
        Get
            Dim eventsPendingProcessing As Boolean = (_EventQueue.Count > 0)
            Return eventsPendingProcessing
        End Get
    End Property
#End Region

#Region "OnEndInvokeJobEvent"
    Public Sub OnEndInvokeJobEvent(ByVal ar As IAsyncResult)
        _JobEvent.EndInvoke(ar)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Job event returned", Me.GetType.ToString)
        End If
    End Sub
#End Region

#Region "OnEndInvokePriterEvent"
    Public Sub OnEndInvokePrinterEvetnt(ByVal ar As IAsyncResult)
        _PrinterEvent.EndInvoke(ar)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Printer event returned", Me.GetType.ToString)
        End If
    End Sub
#End Region

#End Region

#Region "Private methods"

    Private Sub StartThread()
        If _WaitHandle Is Nothing Then
            _WaitHandle = New AutoResetEvent(False)
        End If

        While Not _Cancelled
            Try
                '\\ Wait for the WaitHandle to trigger
                If _WaitHandle.WaitOne(500, False) Then
                    '\\ if the wait handle has not timed out
                    Call ProcessQueue()
                End If
            Catch eTA As ThreadAbortException
                '\\ The thread was aborted prematurely
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                    Trace.WriteLine("EventQueue loop aborted prematurely", Me.GetType.ToString)
                End If
                _Cancelled = True
            End Try
        End While
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("EventQueue loop ended", Me.GetType.ToString)
        End If

    End Sub

    Private Sub ProcessQueue()
        Dim oEvent As EventArgs = Nothing
        Dim ar As IAsyncResult

        _WaitHandle.Reset()
        

        While _EventQueue.Count > 0
            If _EventQueue.TryDequeue(oEvent) Then
                If TypeOf (oEvent) Is PrintJobEventArgs Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                        Trace.WriteLine("Job event dequeued : " & DirectCast(oEvent, PrintJobEventArgs).EventType.ToString, Me.GetType.ToString)
                    End If
                    ar = _JobEvent.BeginInvoke(DirectCast(oEvent, PrintJobEventArgs), AddressOf OnEndInvokeJobEvent, Nothing)
                ElseIf TypeOf (oEvent) Is PrinterEventArgs Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                        Trace.WriteLine("Printer event dequeued", Me.GetType.ToString)
                    End If
                    ar = _PrinterEvent.BeginInvoke(DirectCast(oEvent, PrinterEventArgs), AddressOf OnEndInvokePrinterEvetnt, Nothing)
                End If
            End If
        End While

        '\\ If events have arrived since we started dequeueing them then triger the wait handle again to deal with them
        If EventsPending Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Events pending while dequeueing ", Me.GetType.ToString)
            End If
            _WaitHandle.Set()
        End If
    End Sub
#End Region

#Region "Public constructor"
    Public Sub New(ByVal JobEvent As PrinterMonitorComponent.JobEvent, _
                     ByVal PrinterEvent As PrinterMonitorComponent.PrinterEvent)

        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New()", Me.GetType.ToString)
        End If

        _JobEvent = JobEvent
        _PrinterEvent = PrinterEvent

        _EventQueueWorker = New Thread(AddressOf Me.StartThread)
        Try
            _EventQueueWorker.SetApartmentState(ApartmentState.STA)
        Catch e As Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Failed to set Single Threaded Apartment for : " & _EventQueueWorker.Name, Me.GetType.ToString)
            End If
        End Try
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("StartWatching created new EventQueueWorker", Me.GetType.ToString)
        End If
        _EventQueueWorker.Start()
    End Sub
#End Region

#Region "IDisposable interface"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Dispose()", Me.GetType.ToString)
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Call Shutdown()
            If (Not (_WaitHandle Is Nothing)) Then
                _WaitHandle = Nothing
            End If
        End If


    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#Region "Shutdown"
    Public Sub Shutdown()

        _Cancelled = True
        If Not _WaitHandle Is Nothing Then
            _WaitHandle.Set()
        End If
        If Not _EventQueueWorker Is Nothing Then
            _EventQueueWorker.Join()
            _EventQueueWorker = Nothing
        End If
        _WaitHandle = Nothing
        _EventQueueWorker = Nothing
        _JobEvent = Nothing
        _PrinterEvent = Nothing
    End Sub

#End Region
#End Region

End Class

