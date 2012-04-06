'\\ --[PrinterMonitorComponent]-------------------------
'\\ Top level component to allow plug-in monitoring of
'\\ the windows print spool for a given printer 
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Threading
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports System.Reflection
Imports System.IO

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterMonitorComponent
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Top level component to allow plug-in monitoring of the windows print spool
''' for one or more <see cref="PrinterInformation">printers</see>
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<ToolboxBitmap(GetType(PrinterQueueWatch.PrinterMonitorComponent), "toolboximage.bmp"), _
 System.Security.SuppressUnmanagedCodeSecurity()> _
Public Class PrinterMonitorComponent
    Inherits System.ComponentModel.Component

#Region "Tracing"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the level of trace information output by the PrinterMonitorComponent
    ''' </summary>
    ''' <remarks>
    ''' You can alter the trace switch by adding the switch to the application.exe.config
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ComponentTraceSwitch As New TraceSwitch("PrinterMonitorComponent", "Printer Monitor Component Tracing")
#End Region

#Region "Localisation"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Provides multi-culture support for the component 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ComponentLocalisationResourceManager As New Resources.ResourceManager("PrinterMonitorComponent", System.Reflection.Assembly.Load("PrinterQueueWatch.Resources"))
#End Region

#Region "Public enumerated types"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to configure how much information is returned with a print job event
    ''' </summary>
    ''' <remarks>
    ''' Typically this should be set to MaximumJobInformation except in cases of
    ''' very low bandwidth networks
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Enum MonitorJobEventInformationLevels
        MaximumJobInformation = 1
        MinimumJobInformation = 2
        NoJobInformation = 3
    End Enum

#End Region

#Region "Private constants"
    Private Const DEFAULT_THREAD_TIMEOUT As Integer = 1000
    Private Const INFINITE_THREAD_TIMEOUT As Integer = &HFFFFFFFF
    Private Const PRINTER_NOTIFY_OPTIONS_REFRESH As Integer = &H1
#End Region

#Region "Private Member Variables"

    '\\ Printer handle - returned by the OpenPrinter API call
    Private mhPrinter As IntPtr
    Private msDeviceName As String

    '\\ A combination of PrinterChangeNotificationGeneralFlags that describe what to monitor
    Private _WatchFlags As Integer = PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_JOB Or PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER

    Private _MonitorJobInfoLevel As MonitorJobEventInformationLevels = MonitorJobEventInformationLevels.MaximumJobInformation

    Private _ThreadTimeout As Integer = DEFAULT_THREAD_TIMEOUT

    Private piOut As PrinterInformation

    Private _MonitoredPrinters As MonitoredPrinters

    Private _SpoolMonitoringDisabled As Boolean = False '\\ Switch off spool monitoring if a communications error occurs
#End Region

#Region "Public events"

#Region "Job events"

#Region "Event Delegates"
    <Serializable()> _
    Public Delegate Sub PrintJobEventHandler( _
                  ByVal sender As Object, _
                  ByVal e As PrintJobEventArgs)

    <Serializable()> _
    Public Delegate Sub PrinterEventHandler( _
                      ByVal sender As Object, _
                      ByVal e As PrinterEventArgs)
#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a job is added to one of the print spool queues being monitored
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event JobAdded As PrintJobEventHandler

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a job is removed from one of the print spool queues being monitored
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event JobDeleted As PrintJobEventHandler

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a job is written to on one of the print spool queues being monitored
    ''' </summary>
    ''' <remarks>
    ''' This event is fired when an application writes to the spool file or when a spool
    ''' file writes to the print device
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event JobWritten As PrintJobEventHandler

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a job's properties are changed in one of the print spool queues being monitored
    ''' </summary>
    ''' <remarks>
    ''' Be careful altering the print job in response to this event as you might get an endless loop
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event JobSet As PrintJobEventHandler

#Region "PrintJobSpoolfileParsed"
#If SPOOL_MONITORING_ENABLED Then
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a request to parse a print job spool file completes
    ''' </summary>
    ''' <remarks>
    ''' Spool file parsing is asynchronous and non blocking therefore this event may 
    ''' occur some time after the request was sent
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event PrintJobSpoolfileParsed As EventHandler
#End If
#End Region

    Protected Sub OnJobEvent(ByVal e As PrintJobEventArgs)
        If e.EventType = PrintJobEventArgs.PrintJobEventTypes.JobAddedEvent Then
            RaiseEvent JobAdded(Me, e)

        ElseIf e.EventType = PrintJobEventArgs.PrintJobEventTypes.JobSetEvent Then
            RaiseEvent JobSet(Me, e)
        ElseIf e.EventType = PrintJobEventArgs.PrintJobEventTypes.JobWrittenEvent Then
            RaiseEvent JobWritten(Me, e)
        ElseIf e.EventType = PrintJobEventArgs.PrintJobEventTypes.JobDeletedEvent Then
            RaiseEvent JobDeleted(Me, e)
        End If
    End Sub

#Region "OnSpoolfileParsed"
#If SPOOL_MONITORING_ENABLED Then
    Protected Sub OnSpoolfileParsed(ByVal e As SpoolResponseEventArgs)
        RaiseEvent PrintJobSpoolfileParsed(Me, e)
    End Sub
#End If
#End Region

#End Region

#Region "Printer events"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when the properties of a printer being monitored are changed
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event PrinterInformationChanged As PrinterEventHandler
    Protected Sub OnPrinterInformationChanged(ByVal e As PrinterEventArgs)
        RaiseEvent PrinterInformationChanged(Me, e)
    End Sub
#End Region

#End Region

#Region "Public delegates"
    '\\ Print job events
    Public Delegate Sub JobEvent(ByVal e As PrintJobEventArgs)
    '\\ Print server events
    Public Delegate Sub PrinterEvent(ByVal e As PrinterEventArgs)
#End Region

#Region "Public interface"

#Region "Monitoring"
    ''' <summary>
    ''' True if the component is monitoring any printers
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Monitoring() As Boolean
        Get
            If _MonitoredPrinters Is Nothing Then
                Return False
            Else
                Return (_MonitoredPrinters.Count > 0)
            End If
        End Get
    End Property
#End Region

#Region "DeviceName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to the name of the printer to monitor
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This property is for backward compatibility with the component versions which 
    ''' did not support monitoring multiple printers.  
    ''' Replaced by AddPrinter
    ''' </remarks>
    ''' <exception cref="System.ComponentModel.Win32Exception">
    ''' Thrown when the printer does not exist or the user has no access rights to monitor it
    ''' </exception>
    ''' <seealso cref="PrinterMonitorComponent.AddPrinter"/>
    ''' <seealso cref="PrinterMonitorComponent.RemovePrinter"/>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    <DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public Property DeviceName() As String
        Set(ByVal Value As String)

            If Value = "" Then
                Exit Property
            End If

            '\\ Set the callbacks for the printers list
            If _MonitoredPrinters Is Nothing Then
#If SPOOL_MONITORING_ENABLED Then
                _MonitoredPrinters = New MonitoredPrinters(AddressOf OnPrinterInformationChanged, AddressOf OnJobEvent, AddressOf OnSpoolfileParsed)
#Else
                _MonitoredPrinters = New MonitoredPrinters(AddressOf OnPrinterInformationChanged, AddressOf OnJobEvent)
#End If
            End If

            '\\ If the name changes, destroy and recreate the printer watch worker
            If Not Value.Equals(msDeviceName) Then
                If msDeviceName <> "" Then
                    _MonitoredPrinters.Remove(msDeviceName)
                End If
                msDeviceName = Value

                '\\ Only need to watch a printer in runtime mode...
                If MyBase.DesignMode = False Then
                    If msDeviceName <> "" Then
                        _MonitoredPrinters.Add(msDeviceName, New PrinterInformation(msDeviceName, PrinterAccessRights.PRINTER_ALL_ACCESS Or PrinterAccessRights.SERVER_ALL_ACCESS, _ThreadTimeout, _MonitorJobInfoLevel, _WatchFlags))
                    End If
                End If
            End If
        End Set
        Get
            If Not (_MonitoredPrinters Is Nothing) OrElse _MonitoredPrinters.Count > 0 Then
                Return _MonitoredPrinters.Item(0).PrinterName
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "PrintJobs"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retruns the print job collection of the printer being monitored
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' If more that one printer is being monitored this will return the print jobs 
    ''' of the first one added.  Use the overloaded version to get the named print device's print jobs
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public Overloads ReadOnly Property PrintJobs() As PrintJobCollection
        Get
            If _MonitoredPrinters.Count > 0 Then
                Return _MonitoredPrinters(0).PrintJobs
            Else
                Return New PrintJobCollection
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retruns the print job collection of the named printer being monitored
    ''' </summary>
    ''' <param name="DeviceName">The name of the printer being monitored that you want to retrieve the print jobs for</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public Overloads ReadOnly Property PrintJobs(ByVal DeviceName As String) As PrintJobCollection
        Get
            If _MonitoredPrinters.Contains(DeviceName) Then
                Return _MonitoredPrinters(DeviceName).PrintJobs
            Else
                Return New PrintJobCollection
            End If
        End Get
    End Property
#End Region

#Region "Printer Information"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Returns the printer settings for the printer being monitored
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' If more that one printer is being monitored this will return the details 
    ''' of the first one added.  Use the overloaded version to get the named print device settings
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public Overloads ReadOnly Property PrinterInformation() As PrinterInformation
        Get
            If _MonitoredPrinters.Count > 0 Then
                Return _MonitoredPrinters(0)
            Else
                Throw New ArgumentException("No printers being monitored")
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the printer settings for the named printer being monitored
    ''' </summary>
    ''' <param name="DeviceName">The name of the print device to return the information for</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
        Public Overloads ReadOnly Property PrinterInformation(ByVal DeviceName As String) As PrinterInformation
        Get
            If _MonitoredPrinters.Count > 0 Then
                Return _MonitoredPrinters(DeviceName)
            Else
                Throw New ArgumentException("Printer information not found for this device")
            End If
        End Get
    End Property
#End Region

#Region "Printer info properties"

    <Description("The number of jobs on the queued on the printer being monitored")> _
    Public ReadOnly Property JobCount() As Int32
        Get
            If mhPrinter.ToInt32 <> 0 Then
                Return Me.PrinterInformation.JobCount
            Else
                Return 0
            End If
        End Get
    End Property

#End Region

#Region "AddPrinter"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds the printer to the internal list and starts monitoring it
    ''' </summary>
    ''' <param name="DeviceName">The name of the printer to monitor.
    ''' This can be a UNC name or the name of a printer share
    ''' </param>
    ''' <remarks>
    ''' 
    ''' </remarks>
    ''' <exception cref="System.ComponentModel.Win32Exception">
    ''' Thrown when the printer does not exist or the user has no access rights to monitor it
    ''' </exception>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Adds the printer to the internal list and starts monitoring it")> _
    Public Sub AddPrinter(ByVal DeviceName As String)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("AddPrinter(" & DeviceName & ")", Me.GetType.ToString)
        End If
        If _MonitoredPrinters Is Nothing Then
#If SPOOL_MONITORING_ENABLED Then
            _MonitoredPrinters = New MonitoredPrinters(AddressOf OnPrinterInformationChanged, AddressOf OnJobEvent, AddressOf OnSpoolfileParsed)
#Else
            _MonitoredPrinters = New MonitoredPrinters(AddressOf OnPrinterInformationChanged, AddressOf OnJobEvent)
#End If
        End If
        '\\ Find out if it is a network printer
        Dim piTest As New PrinterInformation(DeviceName, PrinterAccessRights.PRINTER_ACCESS_USE Or PrinterAccessRights.READ_CONTROL, False, False)
        Try
            If piTest.IsNetworkPrinter Then
                _MonitoredPrinters.Add(DeviceName, New PrinterInformation(DeviceName, PrinterAccessRights.PRINTER_ALL_ACCESS Or PrinterAccessRights.SERVER_ALL_ACCESS Or PrinterAccessRights.READ_CONTROL, _ThreadTimeout, _MonitorJobInfoLevel, _WatchFlags))
            Else
                _MonitoredPrinters.Add(DeviceName, New PrinterInformation(DeviceName, PrinterAccessRights.PRINTER_ALL_ACCESS Or PrinterAccessRights.READ_CONTROL, _ThreadTimeout, _MonitorJobInfoLevel, _WatchFlags))
            End If
        Catch ea As Win32Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("AddPrinter(" & DeviceName & ") failed. " & ea.ToString, Me.GetType.ToString)
            End If
        End Try
        piTest.Dispose()

    End Sub
#End Region

#Region "RemovePrinter"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a printer from the internal list, stopping monitoring as appropriate
    ''' </summary>
    ''' <param name="DeviceName">The name of the printer to remove and stop monitoring</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Removes a printer from the internal list, stopping monitoring as appropriate ")> _
    Public Sub RemovePrinter(ByVal DeviceName As String)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("RemovePrinter(" & DeviceName & ")", Me.GetType.ToString)
        End If
        _MonitoredPrinters.Remove(DeviceName)
    End Sub
#End Region

#Region "Disconnect"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disconnects from all printers being monitored
    ''' </summary>
    ''' <remarks>
    ''' You should disconnect from all the printers being monitored before exiting 
    ''' your application to ensure all the resources are released cleanly
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Disconnects from all printers being monitored")> _
    Public Sub Disconnect()
        If Not _MonitoredPrinters Is Nothing Then
            Dim nCount As Integer = _MonitoredPrinters.Count
            For n As Integer = nCount To 1 Step -1
                _MonitoredPrinters.Item(n - 1).Monitored = False
            Next
            _MonitoredPrinters.Clear()
            _MonitoredPrinters = Nothing
        End If
    End Sub
#End Region

#Region "RequestSpoolParse"
#If SPOOL_MONITORING_ENABLED Then
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sent to try and get the print job parsed asynchronously by the
    ''' SpoolMonitorService (if installed)
    ''' </summary>
    ''' <param name="Printername">The printer which is printing the job you want to parse</param>
    ''' <param name="JobId">The job number of the job you want to parse</param>
    ''' <param name="GetPageCount">True if you want to get the page count from the spool file</param>
    ''' <remarks>
    ''' Parsing the spool file is a slow operation and is undertaken asynchronously
    ''' </remarks>
    ''' <seealso cref="PrinterMonitorComponent.PrintJobSpoolfileParsed"/>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub RequestSpoolParse(ByVal Printername As String, _
                                 ByVal JobId As Integer, _
                                 ByVal GetPageCount As Boolean)

        If ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("RequestSpoolParse", Me.GetType.ToString)
        End If
        If _SpoolMonitoringDisabled Then
            If ComponentTraceSwitch.TraceWarning Then
                Trace.WriteLine("RequestSpoolParse : SpoolMonitoringDisabled = True", Me.GetType.ToString)
            End If
            Exit Sub
        End If
        If Not _MonitoredPrinters Is Nothing Then
            If _MonitoredPrinters.Count > 0 Then
                Try
                    _MonitoredPrinters.RequestSpoolParse(Printername, JobId, GetPageCount)
                Catch ex As Exception
                    If ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(ex.ToString, Me.GetType.ToString)
                    End If
                    '\\ Once an error has occured talking to the spool monitor, don't try again
                    SpoolMonitoringDisabled = True
                End Try
            End If
        End If
    End Sub
#End If
#End Region

#Region "SpoolMonitoringDisabled"
#If SPOOL_MONITORING_ENABLED Then
    Public WriteOnly Property SpoolMonitoringDisabled() As Boolean
        Set(ByVal Value As Boolean)
            _SpoolMonitoringDisabled = Value
        End Set
    End Property
#End If
#End Region

#End Region

#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        '\\ Trace the version number and other info of use
        If ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("-- Printer Monitor Component ----------------- ")
            Trace.WriteLine(" Started : " & DateTime.Now.ToLongDateString)
            Trace.WriteLine(" Version : " & Application.ProductVersion)
            Trace.WriteLine(" --------------------------------------------- ")
        End If

#If SPOOL_MONITORING_ENABLED Then
        '\\ Start listening for any incoming spool responses
        Dim _Port As Integer = 8913
        Dim configSettings As New System.Configuration.AppSettingsReader
        Try
            With configSettings
                _Port = Convert.ToInt32(.GetValue("ReturnPort", GetType(System.Int32)))
            End With
            SpoolReciever = New SpoolTCPReceiver(_Port)
        Catch ex As Exception
            If ComponentTraceSwitch.TraceError Then
                Trace.WriteLine(ex.ToString, Me.GetType.ToString)
            End If
            '\\ Don't try and parse spool files if no return port specified
            SpoolMonitoringDisabled = True
        End Try
#End If

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)


        If disposing Then
            Call Disconnect()
            If Not (components Is Nothing) Then
                Try
                    components.Dispose()
                Catch ex As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("Error in Dispose " & ex.ToString, Me.GetType.ToString)
                    End If
                End Try
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

#Region "Design/pre watching interface"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to make the component monitor jobs being added to the job queue
    ''' </summary>
    ''' <value>True to make the component raise a JobAdded event
    ''' when a job is added to a printer being monitored
    ''' </value>
    ''' <remarks>
    ''' Selecting only the notifications you want to be informed about can improve performance
    ''' in low network bandwidth situations
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to make the component monitor jobs being added to the job queue"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property MonitorJobAddedEvent() As Boolean
        Get
            Return ((_WatchFlags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_ADD_JOB) <> 0)
        End Get
        Set(ByVal Value As Boolean)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                If Value Then
                    _WatchFlags = _WatchFlags Or PrinterChangeNotificationJobFlags.PRINTER_CHANGE_ADD_JOB
                Else
                    _WatchFlags = _WatchFlags And Not (PrinterChangeNotificationJobFlags.PRINTER_CHANGE_ADD_JOB)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to make the component monitor jobs being removed from the job queue
    ''' </summary>
    ''' <value>True to make the component raise a JobDeleted event
    ''' when a job is added to a printer being monitored</value>
    ''' <remarks>
    ''' Selecting only the notifications you want to be informed about can improve performance
    ''' in low network bandwidth situations
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to make the component monitor jobs being removed from the job queue"), _
      DefaultValue(GetType(Boolean), "True")> _
    Public Property MonitorJobDeletedEvent() As Boolean
        Get
            Return ((_WatchFlags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_DELETE_JOB) <> 0)
        End Get
        Set(ByVal Value As Boolean)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                If Value Then
                    _WatchFlags = _WatchFlags Or PrinterChangeNotificationJobFlags.PRINTER_CHANGE_DELETE_JOB
                Else
                    _WatchFlags = _WatchFlags And Not (PrinterChangeNotificationJobFlags.PRINTER_CHANGE_DELETE_JOB)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to make the component monitor jobs being written on the job queue
    ''' </summary>
    ''' <value>True to make the component raise a JobWritten event
    ''' when a job is written to on a printer being monitored</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to make the component monitor jobs being written on the job queue"), _
      DefaultValue(GetType(Boolean), "True")> _
    Public Property MonitorJobWrittenEvent() As Boolean
        Get
            Return ((_WatchFlags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_WRITE_JOB) <> 0)
        End Get
        Set(ByVal Value As Boolean)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                If Value Then
                    _WatchFlags = _WatchFlags Or PrinterChangeNotificationJobFlags.PRINTER_CHANGE_WRITE_JOB
                Else
                    _WatchFlags = _WatchFlags And Not (PrinterChangeNotificationJobFlags.PRINTER_CHANGE_WRITE_JOB)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to make the component monitor changes to the jobs on the job queue
    ''' </summary>
    ''' <value>True to make the component raise a JobSet event
    ''' when a job is altered on a printer being monitored</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to make the component monitor changes to the jobs on the job queue"), _
      DefaultValue(GetType(Boolean), "True")> _
     Public Property MonitorJobSetEvent() As Boolean
        Get
            Return ((_WatchFlags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_SET_JOB) <> 0)
        End Get
        Set(ByVal Value As Boolean)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                If Value Then
                    _WatchFlags = _WatchFlags Or PrinterChangeNotificationJobFlags.PRINTER_CHANGE_SET_JOB
                Else
                    _WatchFlags = _WatchFlags And Not (PrinterChangeNotificationJobFlags.PRINTER_CHANGE_SET_JOB)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to make the component monitor printer setup change events
    ''' </summary>
    ''' <value>True to make the component raise a PrinterInformationChanged event
    ''' when a printer being monitored has its settings changed</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to make the component monitor printer setup change events"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property MonitorPrinterChangeEvent() As Boolean
        Get
            Return ((_WatchFlags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER) <> 0)
        End Get
        Set(ByVal Value As Boolean)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                If Value Then
                    _WatchFlags = _WatchFlags Or PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER
                Else
                    _WatchFlags = _WatchFlags And Not (PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to fine tune the job information required for networks
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to fine tune the job information required for networks"), _
    DefaultValue(MonitorJobEventInformationLevels.MaximumJobInformation)> _
    Public Property MonitorJobEventInformationLevel() As MonitorJobEventInformationLevels
        Get
            Return _MonitorJobInfoLevel
        End Get
        Set(ByVal Value As MonitorJobEventInformationLevels)
            If Monitoring Then
                Throw New ReadOnlyException("This property cannot be set once the component is monitoring a print queue")
            Else
                _MonitorJobInfoLevel = Value
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set to tune the printer watch refresh interval
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This property is obsolete and only included for backward compatibility.
    ''' It has no effect on the operation of the component
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Category("Performance Tuning"), _
    Description("Set to tune the printer watch refresh interval"), _
    DefaultValue(DEFAULT_THREAD_TIMEOUT), _
    Obsolete("This property no longer affects the operation of the component")> _
    Public Property ThreadTimeout() As Integer
        Get
            If _ThreadTimeout = 0 Or _ThreadTimeout < -1 Then
                _ThreadTimeout = INFINITE_THREAD_TIMEOUT
            End If
            Return _ThreadTimeout
        End Get
        Set(ByVal Value As Integer)
            If ComponentTraceSwitch.TraceWarning Then
                Trace.WriteLine("Obsolete property ThreadTimeout set", Me.GetType.ToString)
            End If
            _ThreadTimeout = INFINITE_THREAD_TIMEOUT
        End Set
    End Property

#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "Private methods"

#Region "SpoolReciever.SpoolResponse"
#If SPOOL_MONITORING_ENABLED Then
    Private Sub SpoolReciever_SpoolResponse(ByVal sender As Object, ByVal e As System.EventArgs) Handles SpoolReciever.SpoolResponse
        If ComponentTraceSwitch.TraceVerbose OrElse SpoolSocket.SpoolSocketTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Spoolresponse recieved", Me.GetType.ToString)
        End If
        Call OnSpoolfileParsed(DirectCast(e, SpoolResponseEventArgs))
    End Sub
#End If
#End Region

#Region "LicenseExpiryDate"
    Private Function LicenseExpiryDate() As Date
        Try
            Dim foo As [Assembly] = [Assembly].GetAssembly(GetType(PrinterQueueWatch.PrinterMonitorComponent))
            Dim fi As New FileInfo(foo.GetLoadedModules(False)(0).FullyQualifiedName)
            Return fi.LastWriteTime.AddMonths(6)
        Catch ex As Exception
            Return #1/1/2006#
        End Try
    End Function
#End Region
#End Region

End Class

#Region "PrinterEventFlagDecoder class"
Friend Class PrinterEventFlagDecoder
#Region "Private Member Variables"

    Private Const PRINTER_NOTIFY_INFO_DISCARDED As Integer = &H1

    Private mflags As Integer

#End Region

#Region "Public interface"
    ''' <summary>
    ''' Returns true if the printer notification message was complete
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsInfoComplete() As Boolean
        Get
            Return ((mflags And PRINTER_NOTIFY_INFO_DISCARDED) = 0)
        End Get
    End Property

    Public ReadOnly Property ChangesOccured() As Boolean
        Get
            Return Not (mflags = 0)
        End Get
    End Property

    Friend ReadOnly Property JobChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_JOB)
        End Get
    End Property

    Friend ReadOnly Property PrinterChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER)
        End Get
    End Property

    Friend ReadOnly Property ProcessorChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINT_PROCESSOR)
        End Get
    End Property

    Friend ReadOnly Property DriverChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PRINTER_DRIVER)
        End Get
    End Property

    Friend ReadOnly Property FormChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_FORM)
        End Get
    End Property

    Friend ReadOnly Property PortChange() As Integer
        Get
            Return (mflags And PrinterChangeNotificationGeneralFlags.PRINTER_CHANGE_PORT)
        End Get
    End Property

#Region "Job Change events"
    Public ReadOnly Property JobAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_ADD_JOB) <> 0)
        End Get
    End Property

    Public ReadOnly Property JobDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_DELETE_JOB) <> 0)
        End Get
    End Property

    Public ReadOnly Property JobWritten() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_WRITE_JOB) <> 0)
        End Get
    End Property

    Public ReadOnly Property JobSet() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationJobFlags.PRINTER_CHANGE_SET_JOB) <> 0)
        End Get
    End Property
#End Region

#Region "Printer Change events"
    ''' <summary>
    ''' A printer was added to the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPrinterFlags.PRINTER_CHANGE_ADD_PRINTER) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A printer was removed from the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPrinterFlags.PRINTER_CHANGE_DELETE_PRINTER) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' The settings of a printer on the server being monitored were changed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterSet() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPrinterFlags.PRINTER_CHANGE_SET_PRINTER) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A printer connection error occured on the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterConnectionFailed() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPrinterFlags.PRINTER_CHANGE_FAILED_CONNECTION_PRINTER) <> 0)
        End Get
    End Property
#End Region

#Region "Server change events"
#Region "Server Port Events"
    ''' <summary>
    ''' A new printer port was added to the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerPortAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPortFlags.PRINTER_CHANGE_ADD_PORT) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' The settings for one of the ports on the server being monitored changed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerPortSet() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPortFlags.PRINTER_CHANGE_CONFIGURE_PORT) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A port was removed from the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerPortDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationPortFlags.PRINTER_CHANGE_DELETE_PORT) <> 0)
        End Get
    End Property
#End Region

#Region "Server form events"
    ''' <summary>
    ''' A from was added to the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerFormAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationFormFlags.PRINTER_CHANGE_ADD_FORM) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A form was removed from the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerFormDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationFormFlags.PRINTER_CHANGE_DELETE_FORM) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' The properties of a form on the server being monitored were changed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerFormSet() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationFormFlags.PRINTER_CHANGE_SET_FORM) <> 0)
        End Get
    End Property
#End Region

#Region "Server processor events"
    ''' <summary>
    ''' A print processor was added to the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerProcessorAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationProcessorFlags.PRINTER_CHANGE_ADD_PRINT_PROCESSOR) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A print processor was removed from the printer being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerProcessorDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationProcessorFlags.PRINTER_CHANGE_DELETE_PRINT_PROCESSOR) <> 0)
        End Get
    End Property
#End Region

#Region "Server driver events"
    ''' <summary>
    ''' A new printer driver was added to the printer being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerDriverAdded() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationDriverFlags.PRINTER_CHANGE_ADD_PRINTER_DRIVER) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' A printer driver was uninstalled from the server being monitored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerDriverDeleted() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationDriverFlags.PRINTER_CHANGE_DELETE_PRINTER_DRIVER) <> 0)
        End Get
    End Property

    ''' <summary>
    ''' The settings for a printer driver installed on the server being monitored were changed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerDriverSet() As Boolean
        Get
            Return ((mflags And PrinterChangeNotificationDriverFlags.PRINTER_CHANGE_SET_PRINTER_DRIVER) <> 0)
        End Get
    End Property
#End Region

#End Region
#End Region

#Region "Public constructor"
    Public Sub New(ByVal flags As Int32)
        mflags = flags
    End Sub
#End Region

End Class
#End Region

#Region "MonitoredPrinters collection class"
''' <summary>
''' A type safe collection of PrinterInformation objects representing 
''' the printers being monitored by any given PrinterMonitorComponent
''' Unique key is Printer.DeviceName
''' </summary>
''' <remarks></remarks>
Friend Class MonitoredPrinters
    Implements IDisposable

#Region "Private member variables"
    Private _PrinterList As New Generic.SortedList(Of String, PrinterInformation)
    Private _JobEvent As PrinterMonitorComponent.JobEvent
    Private _PrinterEvent As PrinterMonitorComponent.PrinterEvent
    Private _SpoolfileParsers As New SortedList
#End Region

#Region "Public properties"
    Default Public Overloads ReadOnly Property Item(ByVal DeviceName As String) As PrinterInformation
        Get
            If DeviceName Is Nothing Then
                Throw New ArgumentNullException("DeviceName")
            ElseIf DeviceName = "" Then
                Throw New ArgumentException("DeviceName cannot be blank")
            Else
                Return CType(_PrinterList.Item(DeviceName), PrinterInformation)
            End If
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As PrinterInformation
        Get
            Return CType(_PrinterList.Values(Index), PrinterInformation)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return _PrinterList.Count
        End Get
    End Property

#End Region

#Region "Public methods"
#Region "Add"
    Public Sub Add(ByVal DeviceName As String, ByVal PrinterInformation As PrinterInformation)
        If Not _PrinterList.ContainsKey(DeviceName) Then
            _PrinterList.Add(DeviceName, PrinterInformation)
            With Item(DeviceName)
                '\\ Make the PrinterInformation class know that this component is it's event target
                Call .InitialiseEventQueue(_JobEvent, _PrinterEvent)
                '\\ Make the printerinformation class start monitoring
                .Monitored = True
            End With
        End If
    End Sub
#End Region
#Region "Remove"
    Public Sub Remove(ByVal DeviceName As String)
        If _PrinterList.ContainsKey(DeviceName) Then
            DirectCast(_PrinterList.Item(DeviceName), PrinterInformation).Monitored = False
            Call RemoveAt(_PrinterList.IndexOfKey(DeviceName))
        End If
        If _SpoolfileParsers.ContainsKey(DeviceName) Then
            _SpoolfileParsers.RemoveAt(_SpoolfileParsers.IndexOfKey(DeviceName))
        End If
    End Sub

    Public Sub RemoveAt(ByVal Index As Integer)
        _PrinterList.Values(Index).Monitored = False
        _PrinterList.RemoveAt(Index)
    End Sub
#End Region
#Region "Contains"
    Public Function Contains(ByVal Devicename As String) As Boolean
        Return _PrinterList.ContainsKey(Devicename)
    End Function
#End Region
#Region "Clear"
    Public Sub Clear()

        For Each p As PrinterInformation In _PrinterList.Values
            Try
                p.Dispose()
            Catch e As Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("Error in Dispose of " & p.PrinterName & ":: " & e.ToString, Me.GetType.ToString)
                End If
            End Try
        Next
        _PrinterList.Clear()
    End Sub
#End Region

#Region "IDisposable interface implementation"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Dispose()", Me.GetType.ToString)
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Call Clear()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region

#End Region

#Region "Public constructor"
    Public Sub New(ByVal PrinterEventCallback As PrinterMonitorComponent.PrinterEvent, _
               ByVal JobEventCallback As PrinterMonitorComponent.JobEvent)
        _PrinterEvent = PrinterEventCallback
        _JobEvent = JobEventCallback
    End Sub
#End Region
End Class
#End Region


