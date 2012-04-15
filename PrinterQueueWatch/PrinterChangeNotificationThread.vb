'\\ --[PrinterChangeNotificationThread]-----------------------------
'\\ Thread worker to operate a single FindFirstPrinterChange / 
'\\ FindNextPrinterChangeNotification loop for a given printer so
'\\ that a single PrinterMonitorComponent may monitor a number of
'\\ different printer queues.
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------------------
Imports System.ComponentModel
Imports System.Threading
Imports PrinterQueueWatch.SpoolerApi
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

Friend Class PrinterChangeNotificationThread
    Implements IDisposable

#Region "Private constants"

    Private Const INFINITE_THREAD_TIMEOUT As Integer = &HFFFFFFFF
    Private Const PRINTER_NOTIFY_OPTIONS_REFRESH As Integer = &H1

#End Region

#Region "Private members"

    Private _Thread As Thread
    Private _PrinterHandle As Int32

    Private _ThreadTimeout As Integer = INFINITE_THREAD_TIMEOUT

    Private _Cancelled As Boolean = False
    Private _WaitHandle As AutoResetEvent

    Private _MonitorLevel As PrinterMonitorComponent.MonitorJobEventInformationLevels
    Private _WatchFlags As Integer

    Private _PrinterNotifyOptions As PrinterNotifyOptions
    Private _PrinterInformation As PrinterInformation

    Private _PauseAllNewPrintJobs As Boolean

    Private Shared _NotificationLock As New Object
#End Region

#Region "Public methods"

    '\\ Start a watching thread
    Public Sub StartWatching() 'ByVal PrinterHandle As Int32)

        _Thread = New Thread(AddressOf Me.StartThread)
        Try
            _Thread.SetApartmentState(ApartmentState.STA)
        Catch e As Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Failed to set Single Threaded Apartment for : " & _Thread.Name, Me.GetType.ToString)
            End If
        End Try
        _Thread.Name = MyClass.GetType.Name & ":" & _PrinterHandle.ToString
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("StartWatching created new thread: " & _Thread.Name, Me.GetType.ToString)
        End If
        _Thread.Start()

    End Sub

    '\\ Cancel the watching thread
    Public Sub CancelWatching()
        _Cancelled = True
        If Not (_WaitHandle Is Nothing) Then
            If Not (_WaitHandle.SafeWaitHandle.IsInvalid OrElse _WaitHandle.SafeWaitHandle.IsClosed) Then
                _WaitHandle.Set()
            End If
        End If
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("CancelWatching()", Me.GetType.ToString)
        End If

    End Sub

    Public Sub StopWatching()
        Call CancelWatching()


        '\\ And free the printer change notification handle (if it exists)
        If Not _WaitHandle Is Nothing Then
            If Not (_WaitHandle.SafeWaitHandle.IsInvalid) OrElse (_WaitHandle.SafeWaitHandle.IsClosed) Then
                '\\ Stop monitoring the printer
                Try
                    If FindClosePrinterChangeNotification(_WaitHandle.SafeWaitHandle) Then
                        _WaitHandle.Close()
                        _WaitHandle = Nothing
                    Else
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("FindClosePrinterChangeNotification() failed", Me.GetType.ToString)
                        End If
                    End If
                Catch ex As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(ex.ToString, Me.GetType.ToString)
                    End If
                End Try
            End If
        End If

        '\\ Wait for the monitoring thread to terminate
        If Not _Thread Is Nothing Then
            If _Thread.IsAlive Then
                _Thread.Join()
            End If
        End If

        GC.KeepAlive(_WaitHandle)
        '\\ And explicitly unset it
        _Thread = Nothing

    End Sub

    Public Overrides Function ToString() As String
        If Not _WaitHandle Is Nothing Then
            Return Me.GetType.ToString & " handle is valid"
        Else
            Return Me.GetType.ToString & " handle invalid "
        End If
    End Function
#End Region

#Region "Public interface"
#Region "PauseAllNewPrintJobs"
    Public Property PauseAllNewPrintJobs() As Boolean
        Get
            Return _PauseAllNewPrintJobs
        End Get
        Set(ByVal Value As Boolean)
            _PauseAllNewPrintJobs = Value
        End Set
    End Property
#End Region
#End Region

#Region "Public constructors"

    Public Sub New(ByVal PrinterHandle As Int32, ByVal ThreadTimeout As Integer, ByVal MonitorLevel As PrinterMonitorComponent.MonitorJobEventInformationLevels, ByVal WatchFlags As Integer, ByRef PrinterInformation As PrinterInformation)

        '\\ Save a local copy of the information passed in...
        _PrinterHandle = PrinterHandle
        _ThreadTimeout = ThreadTimeout
        _MonitorLevel = MonitorLevel
        _WatchFlags = WatchFlags
        _PrinterInformation = PrinterInformation

    End Sub

    Public Sub New()

    End Sub

#End Region

#Region "Private methods"
    '\\ StartThread - The entry point for this thread
    Private Sub StartThread()

        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("StartThread() of printer handle :" & _PrinterHandle.ToString, Me.GetType.ToString)
        End If

        If _PrinterHandle = 0 Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("StartThread(): _PrinterHandle not set", Me.GetType.ToString)
            End If
            _Cancelled = True
            Exit Sub
        End If

        '\\ Initialise the printer change notification
        Dim mhWait As Microsoft.Win32.SafeHandles.SafeWaitHandle = Nothing

        If _WatchFlags = 0 Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceWarning Then
                Trace.WriteLine("StartWatch: No watch flags set - defaulting to PRINTER_CHANGE_JOB or PRINTER_CHANGE_PRINTER ", Me.GetType.ToString)
            End If
            _WatchFlags = CInt(PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_JOB Or PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER)
        End If

        '\\ Specify what we want to be notified about...
        If _MonitorLevel = PrinterMonitorComponent.MonitorJobEventInformationLevels.MaximumJobInformation Then
            _PrinterNotifyOptions = New PrinterNotifyOptions(False)
        Else
            _PrinterNotifyOptions = New PrinterNotifyOptions(True)
        End If

        If _MonitorLevel = PrinterMonitorComponent.MonitorJobEventInformationLevels.MaximumJobInformation Or _MonitorLevel = PrinterMonitorComponent.MonitorJobEventInformationLevels.MinimumJobInformation Then
            Try
                mhWait = FindFirstPrinterChangeNotification( _
                                                            _PrinterHandle, _
                                                            _WatchFlags, _
                                                            0, _
                                                           _PrinterNotifyOptions)

            Catch e As Win32Exception
                ' An operating system error has been trapped
                If PrinterMonitorComponent.ComponentTraceSwitch.Level >= TraceLevel.Error Then
                    Trace.WriteLine(e.Message & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                End If
            Catch e2 As Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine(e2.Message & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                End If
            Finally
                If Not (mhWait Is Nothing) AndAlso mhWait.IsInvalid Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("StartWatch: FindFirstPrinterChangeNotification failed - handle: " & mhWait.ToString & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                    End If
                    Throw New Win32Exception
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("StartWatch: FindFirstPrinterChangeNotification succeeded - handle: " & mhWait.ToString & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                    End If
                End If
            End Try

        ElseIf _MonitorLevel = PrinterMonitorComponent.MonitorJobEventInformationLevels.NoJobInformation Then
            Try
                mhWait = FindFirstPrinterChangeNotification( _
                                                            _PrinterHandle, _
                                                            _WatchFlags, _
                                                            0, _
                                                            0)

            Catch e As Win32Exception
                '\\ An operating system error has been trapped and returned
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine(e.Message & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                End If
            Catch e2 As Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine(e2.Message & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                End If
            Finally
                If mhWait.IsInvalid Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("StartWatch: FindFirstPrinterChangeNotification failed - handle: " & mhWait.ToString & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                    End If
                    Throw New Win32Exception
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceInfo Then
                        Trace.WriteLine("StartWatch: FindFirstPrinterChangeNotification succeeded - handle: " & mhWait.ToString & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                    End If
                End If
            End Try
        End If

        If mhWait.IsInvalid OrElse mhWait.IsClosed Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("StartWatch: FindFirstPrinterChangeNotification failed for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
            End If
            _Cancelled = True
            Exit Sub
        Else
            If _WaitHandle Is Nothing Then
                _WaitHandle = New AutoResetEvent(False)
            End If
            _WaitHandle.SafeWaitHandle = mhWait
        End If

        While Not _Cancelled
            Try
                '\\ Wait for the WaitHandle to trigger
                If _WaitHandle.WaitOne(-1, False) Then
                    '\\ if the wait handle has not timed out
                    If Not _Cancelled Then
                        Call DecodePrinterChangeInformation()
                    Else
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("StartThread: Cancelled monitoring raised an event " & _PrinterHandle.ToString, Me.GetType.ToString)
                        End If
                    End If
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                        Trace.WriteLine("Wait handle timed out", Me.GetType.ToString)
                    End If
                End If
            Catch eTAE As ThreadAbortException
                '\\ This occurs if the thread was aborted from an external source
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                    Trace.WriteLine("PrinterNotificationThread aborted", Me.GetType.ToString)
                End If
                '\\ Therefore stop watching
                _Cancelled = True
            End Try
        End While

        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("PrinterNotificationThread loop ended", Me.GetType.ToString)
        End If

    End Sub

    ''' <summary>
    ''' When the spooler notifies the monitoring thread that there is a printer change event 
    ''' signalled, this sub decodes that change event and posts it on the event queue
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DecodePrinterChangeInformation()
        Dim mpdChangeFlags As Integer
        Dim mlpPrinter As Int32
        Dim pInfo As PrinterNotifyInfo
        Dim piEventFlags As PrinterEventFlagDecoder

        If _WaitHandle Is Nothing OrElse _WaitHandle.SafeWaitHandle.IsClosed OrElse _WaitHandle.SafeWaitHandle.IsInvalid Then
            Exit Sub
        End If

        '\\ Prevent this code being re-entrant...
        SyncLock _NotificationLock
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("DecodePrinterChangeInformation() for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
            End If

            If _WaitHandle Is Nothing OrElse _WaitHandle.SafeWaitHandle.IsClosed OrElse _WaitHandle.SafeWaitHandle.IsInvalid Then
                Exit Sub
            End If

            '\\ A printer change notification has occured.....
            Try
                If FindNextPrinterChangeNotification(_WaitHandle.SafeWaitHandle, _
                                                    mpdChangeFlags, _
                                                    _PrinterNotifyOptions, _
                                                   mlpPrinter) Then


                    If mlpPrinter <> 0 Then
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                            Trace.WriteLine("FindNextPrinterChangeNotification returned a pointer to PRINTER_NOTIFY_INFO :" & mlpPrinter.ToString & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                        End If

                        pInfo = New PrinterNotifyInfo(_PrinterHandle, mlpPrinter, _PrinterInformation.PrintJobs)

                        piEventFlags = New PrinterEventFlagDecoder(pInfo.Flags)

                        '\\ If the flags indicate that there was insufficient space to store all
                        '\\ the changes we need to ask again
                        If Not piEventFlags.IsInfoComplete Then
                            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                                Trace.WriteLine("FindNextPrinterChangeNotification returned incomplete PRINTER_NOTIFY_INFO" & " for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                            End If

                            _PrinterNotifyOptions.dwFlags = _PrinterNotifyOptions.dwFlags Or PRINTER_NOTIFY_OPTIONS_REFRESH

                            If Not FindNextPrinterChangeNotification(_WaitHandle.SafeWaitHandle, _
                                                         mpdChangeFlags, _
                                                         _PrinterNotifyOptions, _
                                                          mlpPrinter) Then
                                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                                    Trace.WriteLine("FindNextPrinterChangeNotification failed to PRINTER_NOTIFY_OPTIONS_REFRESH for printer handle: " & _PrinterHandle.ToString, Me.GetType.ToString)
                                End If
                                Throw New Win32Exception
                            Else
                                _PrinterNotifyOptions.dwFlags = _PrinterNotifyOptions.dwFlags And (Not PRINTER_NOTIFY_OPTIONS_REFRESH)
                                If mlpPrinter <> 0 Then
                                    pInfo = New PrinterNotifyInfo(_PrinterHandle, mlpPrinter, _PrinterInformation.PrintJobs)
                                End If
                            End If
                        End If


                        '\\ Raise the appropriate event...
                        piEventFlags = New PrinterEventFlagDecoder(mpdChangeFlags)
                        Dim nIndex As Integer
                        Dim thisJob As PrintJob

                        If piEventFlags.ChangesOccured Then
                            For nIndex = 0 To pInfo.PrintJobs.Count - 1
                                thisJob = _PrinterInformation.PrintJobs.ItemByJobId(CType(pInfo.PrintJobs(nIndex), Integer))
                                If Not thisJob Is Nothing Then
                                    If thisJob.JobId <> 0 Then
                                        If piEventFlags.JobAdded Then
                                            If _PauseAllNewPrintJobs Then
                                                thisJob.Paused = True
                                            End If
                                            Dim e As New PrintJobEventArgs(thisJob, PrintJobEventArgs.PrintJobEventTypes.JobAddedEvent)
                                            _PrinterInformation.EventQueue.AddJobEvent(e)
                                        End If
                                        If piEventFlags.JobSet Then
                                            Dim e As New PrintJobEventArgs(thisJob, PrintJobEventArgs.PrintJobEventTypes.JobSetEvent)
                                            _PrinterInformation.EventQueue.AddJobEvent(e)
                                        End If
                                        If piEventFlags.JobWritten Then
                                            Dim e As New PrintJobEventArgs(thisJob, PrintJobEventArgs.PrintJobEventTypes.JobWrittenEvent)
                                            _PrinterInformation.EventQueue.AddJobEvent(e)
                                        End If
                                        If piEventFlags.JobDeleted Then
                                            Dim e As New PrintJobEventArgs(thisJob, PrintJobEventArgs.PrintJobEventTypes.JobDeletedEvent)
                                            _PrinterInformation.EventQueue.AddJobEvent(e)
                                            With _PrinterInformation.PrintJobs
                                                '\\ Remove this item from the printjobs collection
                                                .JobPendingDeletion = CType(pInfo.PrintJobs(nIndex), Integer)
                                            End With
                                        End If
                                    End If
                                End If
                            Next nIndex

                            If piEventFlags.PrinterChange > 0 Then
                                '\\ If the printer info changed throw that event
                                Dim pe As New PrinterEventArgs(pInfo.PrinterInfoChangeFlags, _PrinterInformation)
                                _PrinterInformation.EventQueue.AddPrinterEvent(pe)
                            End If

                            ' -- SERVER_EVENTS ----------------------------------------------
                            ' Not implemented in this release
#If SERVER_EVENTS = 1 Then
                        If piEventFlags.DriverChange > 0 Then
                            'TODO: Pass on the appropriate driver event
                            Dim pde As New PrintServerDriverEventArgs(piEventFlags.DriverChange)

                        End If

                        If piEventFlags.FormChange > 0 Then
                            'TODO: Pass on the appropriate form change event
                            Dim pfe As New PrintServerFormEventArgs(piEventFlags.FormChange)

                        End If

                        If piEventFlags.PortChange > 0 Then
                            'TODO: Pass on the appropriate port change event
                            Dim ppe As New PrintServerPortEventArgs(piEventFlags.PortChange)

                        End If

                        If piEventFlags.ProcessorChange > 0 Then
                            'TODO: Pass on the appropriate processor change event
                            Dim ppce As New PrintServerProcessorEventArgs(piEventFlags.ProcessorChange)
                        End If
#End If

                        End If
                        piEventFlags = Nothing
                    Else
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                            Trace.WriteLine("FindNextPrinterChangeNotification did not return a pointer to PRINTER_NOTIFY_INFO - the change flag was:" & _WatchFlags.ToString, Me.GetType.ToString)
                        End If
                        piEventFlags = New PrinterEventFlagDecoder(mpdChangeFlags)
                        If piEventFlags.PrinterChange > 0 Then
                            '\\ If the printer info changed throw that event
                            Dim pe As New PrinterEventArgs(0, _PrinterInformation)
                            _PrinterInformation.EventQueue.AddPrinterEvent(pe)
                        End If
                    End If

                    _PrinterInformation.EventQueue.Awaken()
                Else
                    '\\ Failed to get the next printer change notification set up...
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("FindNextPrinterChangeNotification failed", Me.GetType.ToString)
                    End If
                    CancelWatching()
                End If
            Catch exP As Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("FindNextPrinterChangeNotification failed : " & exP.ToString, Me.GetType.ToString)
                End If
            End Try
        End SyncLock
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
            Call StopWatching()
            If Not _PrinterNotifyOptions Is Nothing Then
                _PrinterNotifyOptions.Dispose()
                _PrinterNotifyOptions = Nothing
            End If

        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region

End Class
