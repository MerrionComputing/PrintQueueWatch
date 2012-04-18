'\\ --[PrinterInformation]--------------------------------
'\\ Class wrapper for the windows API calls and constants
'\\ relating to getting and setting printer info..
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports System.ComponentModel
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports PrinterQueueWatch.PrinterMonitoringExceptions
Imports System.Runtime.InteropServices

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterInformation
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class which holds the settings for a printer 
''' </summary>
''' <remarks>
''' These settings can apply to physical printers and also to virtual print devices
''' </remarks>
''' <history>
''' 	[Duncan]	20/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrinterInformation
    Implements IDisposable

#Region "Private member variables"

    Private mhPrinter As IntPtr

    '\\ PRINTER_INFO_ structures
    Private mPrinter_Info_2 As New PRINTER_INFO_2
    Private mPrinter_Info_3 As New PRINTER_INFO_3

    Private bHandleOwnedByMe As Boolean

    Private _PrintJobs As PrintJobCollection
    Private _NotificationThread As PrinterChangeNotificationThread
    Private _Monitored As Boolean
    Private _MonitorLevel As PrinterMonitorComponent.MonitorJobEventInformationLevels
    Private _ThreadTimeout As Integer
    Private _WatchFlags As Integer

    Private _JobEvent As PrinterMonitorComponent.JobEvent
    Private _PrinterEvent As PrinterMonitorComponent.PrinterEvent

    Private _TimeWindow As TimeWindow

#End Region

#Region "Friend member variables"

    Private _EventQueue As EventQueue

    Friend ReadOnly Property EventQueue() As EventQueue
        Get
            Return _EventQueue
        End Get
    End Property
#End Region

#Region "IDisposable implementation"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Monitored = False
            If Not _NotificationThread Is Nothing Then
                _NotificationThread.Dispose()
            End If
            If Not _EventQueue Is Nothing Then
                _EventQueue.Dispose()
            End If
            If Not _JobEvent Is Nothing Then
                _JobEvent = Nothing
            End If
            If Not _PrinterEvent Is Nothing Then
                _PrinterEvent = Nothing
            End If
            If bHandleOwnedByMe Then
                Try
                    If Not ClosePrinter(mhPrinter) Then
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("Error in PrinterInformation:Dispose", Me.GetType.ToString)
                        End If
                    Else
                        bHandleOwnedByMe = False
                    End If
                Catch ex32 As Win32Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("Error in PrinterInformation:Dispose " & ex32.ToString, Me.GetType.ToString)
                    End If
                Catch ex As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("Error in PrinterInformation:Dispose " & ex.ToString, Me.GetType.ToString)
                    End If
                End Try

            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region

#Region "Public interface"

#Region "PRINTER_INFO_2 properties"

#Region "ServerName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the server on which this printer is installed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value may be blank if the printer is attached to the current machine
    ''' </remarks>
    ''' <example>Prints the name of the server that the named printer is installed on
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.ServerName)
    '''</code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name of the physical server this printer is attached to")> _
    Public Overridable ReadOnly Property ServerName() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pServerName Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pServerName
            End If
        End Get
    End Property
#End Region

#Region "PrinterName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The unique name by which the printer is known
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The unique name of the printer itself")> _
    Public Overridable ReadOnly Property PrinterName() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pPrinterName Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pPrinterName
            End If
        End Get
    End Property
#End Region

#Region "ShareName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' identifies the sharepoint for the printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This will only be set if the Shared property is set to True
    ''' </remarks>
    ''' <example>Prints the name of the share (if any) that the named printer is shared out on
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.ShareName)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name that the printer is shared under (if it is shared)")> _
    Public Overridable ReadOnly Property ShareName() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pShareName Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pShareName
            End If
        End Get
    End Property
#End Region

#Region "PortName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the port the printer is connected to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the port that the named printer is installed on
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.PortName)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name of the port the printer is connected to")> _
    Public Overridable ReadOnly Property PortName() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pPortName Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pPortName
            End If
        End Get
    End Property
#End Region

#Region "DriverName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the printer driver software used by this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the driver that the named printer is using
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.DriverName)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name of the printer driver software used by this printer")> _
    Public Overridable ReadOnly Property DriverName() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pDriverName Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pDriverName
            End If
        End Get
    End Property
#End Region

#Region "Comment"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The administrator defined comment for this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This can be useful for giving extra information about a printer to the user
    ''' </remarks>
    ''' <example>Changes the comment assigned for this printer
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    pi.Comment = "Monitored by PUMA"
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The administrator defined comment for this printer")> _
    Public Overridable Property Comment() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pComment Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pComment
            End If
        End Get
        Set(ByVal Value As String)
            If Value <> mPrinter_Info_2.pComment Then
                mPrinter_Info_2.pComment = Value
                Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
            End If
        End Set
    End Property
#End Region

#Region "Location"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The administrator defined location for this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the location that the named printer is installed on
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.Location)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The administrator defined location for this printer")> _
    Public Overridable Property Location() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pLocation Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pLocation
            End If
        End Get
        Set(ByVal Value As String)
            If Value <> mPrinter_Info_2.pLocation Then
                mPrinter_Info_2.pLocation = Value
                Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
            End If
        End Set
    End Property
#End Region

#Region "SeperatorFilename"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the file (if any) that is printed to seperate print jobs on this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the print job seperator that the named printer using
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.SeperatorFilename)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name of the file (if any) that is printed to seperate print jobs on this printer")> _
    Public Overridable Property SeperatorFilename() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pSeperatorFilename Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pSeperatorFilename
            End If
        End Get
        Set(ByVal Value As String)
            If Value <> mPrinter_Info_2.pSeperatorFilename Then
                mPrinter_Info_2.pSeperatorFilename = Value
                Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
            End If
        End Set
    End Property
#End Region

#Region "PrintProcessor"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print processor associated to this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the print processor that the named printer using
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.PrintProcessor)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The name of the print processor associated to this printer")> _
    Public Overridable ReadOnly Property PrintProcessor() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pPrintProcessor Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pPrintProcessor
            End If
        End Get
    End Property
#End Region

#Region "DefaultDataType"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default spool data type (e.g. RAW, EMF etc.) used by this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' If this value is set to RAW the printer will spool data in a printer control language 
    ''' (such as PCL or PostScript)
    ''' </remarks>
    ''' <example>Prints the name of the default data type that the named printer using
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.DefaultDataType)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The default spool data type (e.g. RAW, EMF etc.) used by this printer spooler")> _
    Public Overridable ReadOnly Property DefaultDataType() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pDataType Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pDataType
            End If
        End Get
    End Property
#End Region

#Region "Parameters"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Additional parameter used when printing on this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The possible values and meanings of these extra parameters depend on the 
    ''' printer driver being used
    ''' </remarks>
    ''' <example>Prints the extra parameters (if any) that the named printer using
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.Parameters)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("Additional parameter used when printing on this printer")> _
    Public Overridable ReadOnly Property Parameters() As String
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.pParameters Is Nothing Then
                Return ""
            Else
                Return mPrinter_Info_2.pParameters
            End If
        End Get
    End Property
#End Region

#Region "Attributes related"
#Region "IsDefault"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if this printer is the default printer on this machine
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if this printer is the default printer on this machine")> _
    Public Overridable ReadOnly Property IsDefault() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_DEFAULT, Boolean)
        End Get
    End Property
#End Region

#Region "IsShared"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if this printer is a shared device
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if this printer is a shared device")> _
    Public Overridable ReadOnly Property IsShared() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_SHARED, Boolean)
        End Get
    End Property
#End Region

#Region "IsNetworkPrinter"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if this is a network printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if this is a network printer")> _
    Public Overridable ReadOnly Property IsNetworkPrinter() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_NETWORK, Boolean)
        End Get
    End Property
#End Region

#Region "IsLocalPrinter"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if this printer is local to this machine
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if this printer is local to this machine")> _
    Public Overridable ReadOnly Property IsLocalPrinter() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_LOCAL, Boolean)
        End Get
    End Property
#End Region

#End Region

#Region "Priority"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default priority of print jobs sent to this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Priority can range from 1 (lowest) to 99 (highest).  
    ''' Attempting to set the value outside the range will be reset to the nearest range bounds
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The default priority of print jobs sent to this printer")> _
    Public Overridable Property Priority() As Int32
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return mPrinter_Info_2.Priority
        End Get
        Set(ByVal Value As Int32)
            If Value <> mPrinter_Info_2.Priority Then
                If Value < 1 Then
                    mPrinter_Info_2.Priority = 1
                ElseIf Value > 99 Then
                    mPrinter_Info_2.Priority = 99
                Else
                    mPrinter_Info_2.Priority = Value
                    Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
                End If
            End If
        End Set
    End Property
#End Region

#Region "Status related"

#Region "IsReady"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is ready to print
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is ready to print")> _
    Public Overridable ReadOnly Property IsReady() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return (mPrinter_Info_2.Status = 0)
        End Get
    End Property
#End Region

#Region "IsDoorOpen"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because a door or papaer tray is open
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because a door or papaer tray is open")> _
    Public Overridable ReadOnly Property IsDoorOpen() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_DOOR_OPEN) = PrinterStatuses.PRINTER_STATUS_DOOR_OPEN)
        End Get
    End Property
#End Region

#Region "IsInError"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer has an error
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer has an error")> _
    Public Overridable ReadOnly Property IsInError() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_ERROR) = PrinterStatuses.PRINTER_STATUS_ERROR)
        End Get
    End Property
#End Region

#Region "IsInitialising"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is initialising
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is initialising")> _
    Public Overridable ReadOnly Property IsInitialising() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_INITIALIZING) = PrinterStatuses.PRINTER_STATUS_INITIALIZING)
        End Get
    End Property
#End Region

#Region "IsAwaitingManualFeed"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled awaiting a manual paper feed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled awaiting a manual paper feed")> _
    Public Overridable ReadOnly Property IsAwaitingManualFeed() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_MANUAL_FEED) = PrinterStatuses.PRINTER_STATUS_MANUAL_FEED)
        End Get
    End Property
#End Region

#Region "IsOutOfToner"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because it is out of toner or ink
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it is out of toner or ink")> _
    Public Overridable ReadOnly Property IsOutOfToner() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_NO_TONER) = PrinterStatuses.PRINTER_STATUS_NO_TONER)
        End Get
    End Property
#End Region

#Region "IsTonerLow"
    ''' <summary>
    ''' True if the printer is stalled because it is low on toner or ink
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it is low on toner or ink")> _
    Public Overridable ReadOnly Property IsTonerLow() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_TONER_LOW) = PrinterStatuses.PRINTER_STATUS_NO_TONER)
        End Get
    End Property
#End Region

#Region "IsUnavailable"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is currently unnavailable
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is currently unnavailable")> _
    Public Overridable ReadOnly Property IsUnavailable() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_NOT_AVAILABLE) = PrinterStatuses.PRINTER_STATUS_NOT_AVAILABLE)
        End Get
    End Property
#End Region

#Region "IsOffline"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is offline
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is offline")> _
    Public Overridable ReadOnly Property IsOffline() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_OFFLINE) = PrinterStatuses.PRINTER_STATUS_OFFLINE)
        End Get
    End Property
#End Region

#Region "IsOutOfMemory"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because it has run out of memory
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it has run out of memory")> _
    Public Overridable ReadOnly Property IsOutOfmemory() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_OUT_OF_MEMORY) = PrinterStatuses.PRINTER_STATUS_OUT_OF_MEMORY)
        End Get
    End Property
#End Region

#Region "IsOutputBinFull"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because it's output tray is full
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it's output tray is full")> _
    Public Overridable ReadOnly Property IsOutputBinFull() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_OUTPUT_BIN_FULL) = PrinterStatuses.PRINTER_STATUS_OUTPUT_BIN_FULL)
        End Get
    End Property
#End Region

#Region "IsPaperJammed"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because it has a paper jam
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it has a paper jam")> _
    Public Overridable ReadOnly Property IsPaperJammed() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_PAPER_JAM) = PrinterStatuses.PRINTER_STATUS_PAPER_JAM)
        End Get
    End Property
#End Region

#Region "IsOutOfPaper"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled because it is out of paper
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled because it is out of paper")> _
    Public Overridable ReadOnly Property IsOutOfPaper() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_PAPER_OUT) = PrinterStatuses.PRINTER_STATUS_PAPER_OUT)
        End Get
    End Property
#End Region

#Region "Paused"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is paused
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is paused")> _
    Public Overridable ReadOnly Property Paused() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_PAUSED) = PrinterStatuses.PRINTER_STATUS_PAUSED)
        End Get
    End Property
#End Region

#Region "IsDeletingJob"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is deleting a job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is deleting a job")> _
    Public Overridable ReadOnly Property IsDeletingJob() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_PENDING_DELETION) = PrinterStatuses.PRINTER_STATUS_PENDING_DELETION)
        End Get
    End Property
#End Region

#Region "IsInPowerSave"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is in power saving mode
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is in power saving mode")> _
    Public Overridable ReadOnly Property IsInPowerSave() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_POWER_SAVE) = PrinterStatuses.PRINTER_STATUS_POWER_SAVE)
        End Get
    End Property
#End Region

#Region "IsPrinting"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is currently printing a job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is currently printing a job")> _
    Public Overridable ReadOnly Property IsPrinting() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_PRINTING) = PrinterStatuses.PRINTER_STATUS_PRINTING)
        End Get
    End Property
#End Region

#Region "IsWaitingOnUserIntervention"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is stalled awaiting manual intervention
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is stalled awaiting manual intervention")> _
    Public Overridable ReadOnly Property IsWaitingOnUserIntervention() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_USER_INTERVENTION) = PrinterStatuses.PRINTER_STATUS_USER_INTERVENTION)
        End Get
    End Property
#End Region

#Region "IsWarmingUp"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer is warming up to be ready to print
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The status is updated when a job is sent to the printer so may not match the true state of the printer 
    ''' if there are no jobs in the queue
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("True if the printer is warming up to be ready to print")> _
    Public Overridable ReadOnly Property IsWarmingUp() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return ((mPrinter_Info_2.Status And PrinterStatuses.PRINTER_STATUS_WARMING_UP) = PrinterStatuses.PRINTER_STATUS_WARMING_UP)
        End Get
    End Property
#End Region
#End Region

#Region "JobCount"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of print jobs queued to be printed by this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The number of jobs waiting on this printer")> _
    Public Overridable ReadOnly Property JobCount() As Int32
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return mPrinter_Info_2.JobCount
        End Get
    End Property
#End Region

#Region "AveragePagesPerMonth"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The average throughput of this printer in pages per month
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The average throughput of this printer in pages per month")> _
    Public Overridable ReadOnly Property AveragePagesPerMonth() As Int32
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return mPrinter_Info_2.AveragePPM
        End Get
    End Property
#End Region

#Region "TimeWindow"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The time window within which jobs can be scheduled against this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Obsolete("Use Availablity instead."), Diagnostics.MonitoringDescription("The time window within which jobs can be scheduled against this printer")> _
    Public Property TimeWindow() As TimeWindow
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            _TimeWindow = New TimeWindow(mPrinter_Info_2.StartTime, mPrinter_Info_2.UntilTime)
            Return _TimeWindow
        End Get
        Set(ByVal Value As TimeWindow)
            mPrinter_Info_2.StartTime = Value.StartTimeMinutes
            mPrinter_Info_2.UntilTime = Value.EndTimeMinutes
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "DefaultPaperSource"
    ''' <summary>
    ''' The default tray used by this printer 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>This value can be overriden for each individual print job
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	11/02/2006	Created
    ''' </history>
    <Diagnostics.MonitoringDescription("The default paper tray for this printer")> _
    Public Property DefaultPaperSource() As Drawing.Printing.PaperSourceKind
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.DeviceMode.dmDefaultSource, Drawing.Printing.PaperSourceKind)
        End Get
        Set(ByVal value As Drawing.Printing.PaperSourceKind)
            mPrinter_Info_2.DeviceMode.dmDefaultSource = CShort(value)
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "Copies"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of copies of each print job to produce 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value can be overridden for each print job
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	11/02/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Copies() As Integer
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            If mPrinter_Info_2.DeviceMode.dmCopies > 0 Then
                Return mPrinter_Info_2.DeviceMode.dmCopies
            Else
                Return 1
            End If
        End Get
    End Property
#End Region

#Region "Landscape"
    ''' <summary>
    ''' True if the printer orientation is set to Landscape
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' This value can be overriden by the individual print job's orientation
    ''' </remarks>
    Public Property Landscape() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return (mPrinter_Info_2.DeviceMode.dmOrientation = DeviceOrientations.DMORIENT_LANDSCAPE)
        End Get
        Set(ByVal value As Boolean)
            If value Then
                mPrinter_Info_2.DeviceMode.dmOrientation = CShort(DeviceOrientations.DMORIENT_LANDSCAPE)
            Else
                mPrinter_Info_2.DeviceMode.dmOrientation = CShort(DeviceOrientations.DMORIENT_PORTRAIT)
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "Colour"
    ''' <summary>
    ''' True if a colour printer is set to print in colour, false for monochrome
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Colour() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return (mPrinter_Info_2.DeviceMode.dmColor = DeviceColourModes.DMCOLOR_COLOR)
        End Get
        Set(ByVal value As Boolean)
            If value Then
                mPrinter_Info_2.DeviceMode.dmColor = CShort(DeviceColourModes.DMCOLOR_COLOR)
            Else
                mPrinter_Info_2.DeviceMode.dmColor = CShort(DeviceColourModes.DMCOLOR_MONOCHROME)
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "Collate"
    ''' <summary>
    ''' Specifies whether collation should be used when printing multiple copies.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Not all printers support collation.  Those that don't will ignore this setting
    ''' </remarks>
    Public Property Collate() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return (mPrinter_Info_2.DeviceMode.dmCollate = DeviceCollateSettings.DMCOLLATE_TRUE)
        End Get
        Set(ByVal value As Boolean)
            If value Then
                mPrinter_Info_2.DeviceMode.dmCollate = CShort(DeviceCollateSettings.DMCOLLATE_TRUE)
            Else
                mPrinter_Info_2.DeviceMode.dmCollate = CShort(DeviceCollateSettings.DMCOLLATE_FALSE)
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "PrintQuality"
    Public Property PrintQuality() As DeviceModeResolutions
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.DeviceMode.dmPrintQuality, DeviceModeResolutions)
        End Get
        Set(ByVal value As DeviceModeResolutions)
            mPrinter_Info_2.DeviceMode.dmPrintQuality = CShort(value)
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#Region "Scale"
    ''' <summary>
    ''' The scale (percentage) to print the page at
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Property Scale() As Short
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return mPrinter_Info_2.DeviceMode.dmScale
        End Get
        Set(ByVal value As Short)
            If value <= 0 OrElse value > 100 Then
                Throw New ArgumentException("Scale cannot be less than or equal to zero nor greater than 100")
            End If
            mPrinter_Info_2.DeviceMode.dmScale = value
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#End Region

#Region "PRINTER_INFO_3 properties"

#Region "SecurityDescriptorPointer"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property SecurityDescriptorPointer() As Integer
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel3)
            Return mPrinter_Info_3.pSecurityDescriptor
        End Get
    End Property
#End Region

#End Region

#Region "Public Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pauses the printer 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Pauses the named printer
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, False)
    '''    pi.PausePrinting
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub PausePrinting()
        '\\ Do not attempt to pause and already paused printer
        If Not Paused Then
            Try
                If Not SetPrinter(mhPrinter, 0, IntPtr.Zero, PrinterControlCommands.PRINTER_CONTROL_PAUSE) Then
                    Throw New Win32Exception()
                End If
            Catch e As Win32Exception
                If e.NativeErrorCode = SpoolerWin32ErrorCodes.ERROR_ACCESS_DENIED Then
                    Throw New InsufficentPrinterAccessRightsException(My.Resources.pem_NoPause, e)
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("SetPrinter (Pause) failed", Me.GetType.ToString)
                    End If
                    Throw
                End If
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Restart a printer that has been paused
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Resumes printing on the named printing
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    pi.ResumePrinting
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ResumePrinting()
        '\\ Do not attempt to resume if the printer is not paused
        If Paused Then
            Try
                If Not SetPrinter(mhPrinter, 0, IntPtr.Zero, PrinterControlCommands.PRINTER_CONTROL_RESUME) Then
                    Throw New Win32Exception()
                End If
            Catch e As Win32Exception
                If e.NativeErrorCode = SpoolerWin32ErrorCodes.ERROR_ACCESS_DENIED Then
                    Throw New InsufficentPrinterAccessRightsException(My.Resources.pem_NoResume, e)
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("SetPrinter (Resume) failed", Me.GetType.ToString)
                    End If
                    Throw
                End If
            End Try
        End If
    End Sub

#End Region

#Region "PrintJobs collection"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The collection of PrintJobs queued for printing on this printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the documents in the print queue
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    For Each pj As PrintJob In pi.PrintJobs
    '''       Trace.WriteLine(pj.Document)
    '''    Next pj
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("The collection of PrintJobs queued for printing on this printer")> _
    Public ReadOnly Property PrintJobs() As PrintJobCollection
        Get
            If _PrintJobs Is Nothing Then
                _PrintJobs = New PrintJobCollection(mhPrinter, JobCount)
            End If
            If _PrintJobs.JobPendingDeletion > 0 Then
                _PrintJobs.RemoveByJobId(_PrintJobs.JobPendingDeletion)
                _PrintJobs.JobPendingDeletion = 0
            End If
            Return _PrintJobs
        End Get
    End Property
#End Region

#Region "Monitored"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets whether or not events occuring on this printer are raised by the component
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Diagnostics.MonitoringDescription("Sets whether or not events occuring on this printer are raised by the component")> _
    Public Property Monitored() As Boolean
        Get
            If Not _NotificationThread Is Nothing Then
                Return _Monitored
            Else
                _Monitored = False
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            If Value <> _Monitored Then
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                    Trace.WriteLine("Monitored set to : " & Value.ToString, Me.GetType.ToString)
                End If

                If Value Then
                    If Not _NotificationThread Is Nothing Then
                        _NotificationThread.Dispose()
                    End If
                    Try
                        _NotificationThread = New PrinterChangeNotificationThread(mhPrinter, _ThreadTimeout, _MonitorLevel, _WatchFlags, Me)
                    Catch e As Exception
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine(e.Message & " creating new PrinterChangeNotificationThread for handle : " & mhPrinter.ToString, Me.GetType.ToString)
                        End If
                    End Try
                    _EventQueue = New EventQueue(_JobEvent, _PrinterEvent)
                    If Not _NotificationThread Is Nothing Then
                        _NotificationThread.StartWatching()
                    End If
                Else
                    If Not _NotificationThread Is Nothing Then
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                            Trace.WriteLine("Stop monitoring  _NotificationThread : " & _NotificationThread.ToString, Me.GetType.ToString)
                        End If
                        _NotificationThread.StopWatching()
                    End If
                End If
                _Monitored = Value
            End If
        End Set
    End Property
#End Region

#Region "PauseAllNewJobs"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' If true and the printer is being monitored, all print jobs are paused as they 
    ''' are added to the spool queue
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This is useful for print quota type applications as the print job is immediately 
    ''' paused allowing the quota program to decide whether or not to delete or resume it
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	07/01/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property PauseAllNewJobs() As Boolean
        Get
            If Not _NotificationThread Is Nothing Then
                Return _NotificationThread.PauseAllNewPrintJobs
            End If
        End Get
        Set(ByVal Value As Boolean)
            If Not _NotificationThread Is Nothing Then
                _NotificationThread.PauseAllNewPrintJobs = Value
            End If
        End Set
    End Property
#End Region

#Region "InitialiseEventQueue"
    Friend Sub InitialiseEventQueue(ByVal JobEventCallback As PrinterMonitorComponent.JobEvent, ByVal PrinterEventCallback As PrinterMonitorComponent.PrinterEvent)

        _JobEvent = JobEventCallback
        _PrinterEvent = PrinterEventCallback
        _EventQueue = New EventQueue(_JobEvent, _PrinterEvent)

    End Sub
#End Region

#Region "CanLoggedInUserAdministerPrinter"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Function CanLoggedInUserAdministerPrinter() As Boolean
        '\\ Need to get the DACL for the printer and see if the logged in user is allowed the administer role..?
        Return True
    End Function
#End Region

#Region "Printer Driver"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns information about the printer driver used by a given printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the printer driver 
    ''' <code>
    '''    Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
    '''    Trace.WriteLine(pi.PrinterDriver.Name)
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PrinterDriver() As PrinterDriver
        Get
            Return New PrinterDriver(mhPrinter)
        End Get
    End Property
#End Region

#Region "PrinterForms"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the collection of print forms installed on the printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>List the print forms on the named printer
    ''' <code>
    '''        Dim pi As New PrinterInformation("HP Laserjet 5L", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, False)
    '''
    '''        For pf As Integer = 0 To pi.PrinterForms.Count - 1
    '''            Me.ListBox1.Items.Add( pi.PrinterForms(pf).Name )
    '''        Next
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PrinterForms() As PrinterFormCollection
        Get
            Return New PrinterFormCollection(mhPrinter)
        End Get
    End Property
#End Region

#Region "Public overrides"
    Public Overrides Function ToString() As String
        Return Me.PrinterName
    End Function
#End Region

#Region "Advanced Printer Properties"

    ''' <summary>
    ''' Is the printer for 24-hour availability.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public ReadOnly Property IsAlwaysAvailable() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.StartTime.CompareTo(mPrinter_Info_2.UntilTime) = 0, Boolean)
        End Get
        'Set(ByVal value As Boolean)
        '    If Not IsAlwaysAvailable = value Then
        '        Availability = New TimeWindow(0, 0)
        '    End If
        'End Set
    End Property

    ''' <summary>
    ''' Configures the printer for limited availability. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If you send a document to a printer when it is unavailable, the document will be held (spooled) until the printer is available.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property Availability() As TimeWindow
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return New TimeWindow(mPrinter_Info_2.StartTime, mPrinter_Info_2.UntilTime)
        End Get
        Set(ByVal value As TimeWindow)
            If Not mPrinter_Info_2.StartTime = value.StartTimeMinutes Then
                mPrinter_Info_2.StartTime = value.StartTimeMinutes
            End If
            If Not mPrinter_Info_2.UntilTime = value.EndTimeMinutes Then
                mPrinter_Info_2.UntilTime = value.EndTimeMinutes
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Specifies that documents should be spooled before being printed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' Spooling is the process of first storing the document on the hard disk and then sending the document to the print device. You can continue working with your program as soon as the document is stored on the disk. The spooler sends the document to the print device in the background.
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property SpoolPrintJobs() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return Not CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_DIRECT, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be throwed
            If Not SpoolPrintJobs = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_DIRECT, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_DIRECT, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Specifies that the print device should wait to begin printing until after the last page of the document is spooled. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' The printing program is unavailable until the document has finished spooling. 
    ''' However, using this option ensures that the whole document is available to the print device.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property EnableSpoolBeforePrint() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_QUEUED, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be throwed
            If Not EnableSpoolBeforePrint = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_QUEUED, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_QUEUED, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Directs the spooler to check the printer setup and match it to the document setup before sending the document to the print device. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the information does not match, the document is held in the queue.
    ''' A mismatched document in the queue will not prevent correctly matched documents from printing.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property HoldMismatchedDocuments() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_ENABLE_DEVQ, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be throwed
            If Not HoldMismatchedDocuments = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_ENABLE_DEVQ, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_ENABLE_DEVQ, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Specifies that the spooler should favor documents that have completed spooling when deciding which document to print next, even if the completed documents are a lower priority than documents that are still spooling. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If there are no documents that have completed spooling, the spooler will favor larger spooling documents over smaller ones. 
    ''' Use this option if you want to maximize printer efficiency.
    ''' When this option is disabled, the spooler picks documents based only on priority.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property PrintSpooledDocumentsFirst() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be throwed
            If Not PrintSpooledDocumentsFirst = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Specifies that the spooler should not delete documents after they are printed. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' This allows a document to be resubmitted to the printer from the printer queue instead of from the program.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property KeepPrintedDocuments() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be throwed
            If Not KeepPrintedDocuments = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property

    ''' <summary>
    ''' Specifies whether the advanced printing feature is enabled. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' When enabled, metafile spooling is turned on and options such as Page Order, Booklet Printing, and Pages Per Sheet are available, depending on your printer. For normal printing, leave the advanced printing feature set to the default (Enabled). If compatibility problems occur, you can disable the feature. When disabled, metafile spooling is turned off and the printing options might be unavailable.
    ''' </remarks>
    ''' <history>
    ''' 	[solidcrip]	13/11/2006	Created
    ''' </history>
    Public Property EnableAdvancedPrintingFeatures() As Boolean
        Get
            RefreshPrinterInformation(PrinterInfoLevels.PrinterInfoLevel2)
            Return CType(mPrinter_Info_2.Attributes And PrinterAttributes.PRINTER_ATTRIBUTE_RAW_ONLY, Boolean)
        End Get
        Set(ByVal value As Boolean)
            'value must be different, otherwise a StackOverflowException will be thrown
            If Not EnableAdvancedPrintingFeatures = value Then
                If value Then
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Or PrinterAttributes.PRINTER_ATTRIBUTE_RAW_ONLY, Int32)
                Else
                    mPrinter_Info_2.Attributes = CType(mPrinter_Info_2.Attributes Xor PrinterAttributes.PRINTER_ATTRIBUTE_RAW_ONLY, Int32)
                End If
            End If
            Call SavePrinterInfo(PrinterInfoLevels.PrinterInfoLevel2, False)
        End Set
    End Property
#End Region

#End Region

#Region "Private methods"
    Private Sub RefreshPrinterInformation(ByVal level As PrinterInfoLevels)

        If mhPrinter.Equals(0) Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("RefreshPrinterInformation failed: Handle is invalid", Me.GetType.ToString)
            End If
            Exit Sub
        End If

        Select Case level
            Case PrinterInfoLevels.PrinterInfoLevel2
                Try
                    mPrinter_Info_2 = New PRINTER_INFO_2(mhPrinter)
                Catch e As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(e.Message & " getting PRINTER_INFO_2 for handle: " & mhPrinter.ToString, Me.GetType.ToString)
                    End If
                End Try

            Case PrinterInfoLevels.PrinterInfoLevel3
                Try
                    mPrinter_Info_3 = New PRINTER_INFO_3(mhPrinter)
                Catch e As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(e.Message & " getting PRINTER_INFO_2 for handle: " & mhPrinter.ToString, Me.GetType.ToString)
                    End If
                End Try

            Case Else
                '\\ Not yet implemented...
        End Select
    End Sub

    Private Sub SavePrinterInfo(ByVal Level As PrinterInfoLevels, ByVal ModifySecurityDescriptor As Boolean)

        Select Case Level
            Case PrinterInfoLevels.PrinterInfoLevel2
                If Not ModifySecurityDescriptor Then
                    mPrinter_Info_2.lpSecurityDescriptor = 0
                End If
                Try
                    SetPrinter(mhPrinter, PrinterInfoLevels.PrinterInfoLevel2, mPrinter_Info_2, 0)
                Catch e As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(e.Message & " setting PRINTER_INFO_2 for handle: " & mhPrinter.ToString, Me.GetType.ToString)
                    End If
                    Throw New PrinterMonitoringExceptions.InsufficentPrinterAccessRightsException(e)
                End Try


            Case PrinterInfoLevels.PrinterInfoLevel3
                Try
                    SetPrinter(mhPrinter, PrinterInfoLevels.PrinterInfoLevel3, mPrinter_Info_3, 0)
                Catch e As Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine(e.Message & " setting PRINTER_INFO_3 for handle: " & mhPrinter.ToString, Me.GetType.ToString)
                    End If
                    Throw New PrinterMonitoringExceptions.InsufficentPrinterAccessRightsException(e)
                End Try

            Case Else
                '\\ Not yet implemented...
        End Select

    End Sub
#End Region

#Region "Public constructors"

    Private Sub InitPrinterInfo(ByVal GetJobs As Boolean)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("InitPrinterInfo()", Me.GetType.ToString)
        End If
        '\\ Get the current printer information
        Try
            mPrinter_Info_2 = New PRINTER_INFO_2(mhPrinter)
        Catch e As Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine(e.Message & " creating new PRINTER_INFO_2 for handle : " & mhPrinter.ToString, Me.GetType.ToString)
            End If
            Throw
            Exit Sub
        End Try
        If GetJobs Then
            Try
                _PrintJobs = New PrintJobCollection(mhPrinter, JobCount)
            Catch e As Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine(e.Message & " creating new PRINTER_INFO_3 for handle : " & mhPrinter.ToString, Me.GetType.ToString)
                End If
                Throw
            End Try
        End If
    End Sub

    Friend Sub New(ByVal PrinterHandle As IntPtr)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & PrinterHandle.ToString & ")", Me.GetType.ToString)
        End If
        mhPrinter = PrinterHandle
        Call InitPrinterInfo(True)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new printer information class for the named printer
    ''' </summary>
    ''' <param name="DeviceName">The name of the print device</param>
    ''' <param name="DesiredAccess">The required access rights for that printer</param>
    ''' <param name="GetJobs">True to return the collection of print jobs
    ''' queued against this print device
    ''' </param>
    ''' <remarks>
    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal DeviceName As String, ByVal DesiredAccess As SpoolerApiConstantEnumerations.PrinterAccessRights, ByVal GetJobs As Boolean)
        Me.New(DeviceName, DesiredAccess, (((DesiredAccess And PrinterAccessRights.SERVER_ACCESS_ADMINISTER) Or (DesiredAccess And PrinterAccessRights.PRINTER_ACCESS_ADMINISTER)) <> 0), GetJobs)
        bHandleOwnedByMe = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new printer information class for the named printer
    ''' </summary>
    ''' <param name="DeviceName">The name of the print device</param>
    ''' <param name="DesiredAccess">The required access rights for that printer</param>
    ''' <param name="GetSecurityInfo"></param>
    ''' <param name="GetJobs">True to return the collection of print jobs
    ''' queued against this print device
    ''' </param>
    ''' <remarks>
    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal DeviceName As String, ByVal DesiredAccess As SpoolerApiConstantEnumerations.PrinterAccessRights, ByVal GetSecurityInfo As Boolean, ByVal GetJobs As Boolean)
        Dim hPrinter As IntPtr = IntPtr.Zero
        If OpenPrinter(DeviceName, hPrinter, New PRINTER_DEFAULTS(DesiredAccess)) Then
            mhPrinter = hPrinter
            Call InitPrinterInfo(GetJobs)
        ElseIf OpenPrinter(DeviceName, hPrinter, New PRINTER_DEFAULTS(PrinterAccessRights.PRINTER_ALL_ACCESS)) Then
            mhPrinter = hPrinter
            Call InitPrinterInfo(GetJobs)
        ElseIf OpenPrinter(DeviceName, hPrinter, 0) Then
            mhPrinter = hPrinter
            Call InitPrinterInfo(GetJobs)
        Else
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("OpenPrinter() failed", Me.GetType.ToString)
            End If
            Throw New Win32Exception
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new printer information class for the named printer
    ''' </summary>
    ''' <param name="DeviceName">The name of the print device</param>
    ''' <param name="DesiredAccess">The required access rights for that printer</param>
    ''' <param name="ThreadTimeout">No longer used</param>
    ''' <param name="MonitorLevel"></param>
    ''' <param name="WatchFlags"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal DeviceName As String, ByVal DesiredAccess As SpoolerApiConstantEnumerations.PrinterAccessRights, ByVal ThreadTimeout As Integer, ByVal MonitorLevel As PrinterMonitorComponent.MonitorJobEventInformationLevels, ByVal WatchFlags As Integer)
        Me.New(DeviceName, DesiredAccess, True)
        _ThreadTimeout = ThreadTimeout
        _MonitorLevel = MonitorLevel
        _WatchFlags = WatchFlags
    End Sub

    Friend Sub New(ByVal DeviceName As String, ByVal DesiredAccess As SpoolerApiConstantEnumerations.PrinterAccessRights, ByVal ThreadTimeout As Integer, ByVal MonitorLevel As PrinterMonitorComponent.MonitorJobEventInformationLevels, _
                   ByVal WatchFlags As Integer, ByVal JobEventCallback As PrinterMonitorComponent.JobEvent, ByVal PrinterEventCallback As PrinterMonitorComponent.PrinterEvent)

        Me.New(DeviceName, DesiredAccess, ThreadTimeout, MonitorLevel, WatchFlags)
        _JobEvent = JobEventCallback
        _PrinterEvent = PrinterEventCallback
        If _NotificationThread Is Nothing Then
            _NotificationThread = New PrinterChangeNotificationThread(mhPrinter, _ThreadTimeout, _MonitorLevel, _WatchFlags, Me)
        End If
        _EventQueue = New EventQueue(_JobEvent, _PrinterEvent)

    End Sub

    Friend Sub New(ByVal DeviceName As String, _
                   ByVal Description As String, _
                   ByVal Comment As String, _
                   ByVal ServerName As String, _
                   ByVal Index As Integer)

        '\\ -- Don't try to refresh this printer info...
        bHandleOwnedByMe = False
        mPrinter_Info_2 = New PRINTER_INFO_2
        With mPrinter_Info_2
            .pPrinterName = DeviceName
            .pComment = Comment
            .pLocation = Description
            .pServerName = ServerName
        End With

    End Sub
#End Region

#Region "Friend methods"
    Friend Function MakeUrl(ByVal TransportProtocol As String, _
                           ByVal Port As Integer) As String
        '\\ Return the URL  
        Dim sRet As String
        If Me.ServerName <> "" Then
            sRet = TransportProtocol & "://" & Me.ServerName & ":" & Port.ToString & "/SpoolMonitorService"
        Else
            sRet = TransportProtocol & "://localhost:" & Port.ToString & "/SpoolMonitorService"
        End If
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("MakeUrl : " & sRet, Me.GetType.ToString)
        End If
        Return sRet
    End Function

#End Region

End Class

#Region "PrinterInformationCollection"
''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterInformationCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of printer information classes
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	20/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrinterInformationCollection
    Inherits Generic.List(Of PrinterInformation)

#Region "Public interface"
    Public Overloads Sub Remove(ByVal obj As PrinterInformation)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a collection of printer information for the current machine 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    '''     [Duncan]    23/02/2008  Changed to use PRINTER_INFO_1 structure
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pPrinters As Int32
        Dim pcbProvided As Int32 = 0

        If Not EnumPrinters(EnumPrinterFlags.PRINTER_ENUM_NAME, String.Empty, 1, pPrinters, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pPrinters = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrinters(EnumPrinterFlags.PRINTER_ENUM_NAME, String.Empty, 1, pPrinters, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPrinters
            While pcReturned > 0
                Dim pi1 As New PRINTER_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), pi1)
                If Not pi1.pName Is Nothing Then
                    Me.Add(New PrinterInformation(pi1.pName, PrinterAccessRights.PRINTER_ACCESS_USE, False)) ', pi2.pLocation, pi2.pComment, pi2.pServerName, 1))
                End If
                ptNext = ptNext + Marshal.SizeOf(GetType(PRINTER_INFO_1))
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPrinters > 0 Then
            Marshal.FreeHGlobal(CType(pPrinters, IntPtr))
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a collection of printer information objects for the named machine 
    ''' </summary>
    ''' <param name="Servername">The name of the server to list the printer devices</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Servername As String)
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pPrinters As Int32
        Dim pcbProvided As Int32 = 0

        If Not EnumPrinters(EnumPrinterFlags.PRINTER_ENUM_NAME, Servername, 1, pPrinters, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pPrinters = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrinters(EnumPrinterFlags.PRINTER_ENUM_NAME, Servername, 1, pPrinters, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPrinters
            While pcReturned > 0
                Dim pi2 As New PRINTER_INFO_2
                Marshal.PtrToStructure(New IntPtr(ptNext), pi2)
                Me.Add(New PrinterInformation(pi2.pPrinterName, pi2.pLocation, pi2.pComment, pi2.pServerName, 1))
                ptNext = ptNext + Marshal.SizeOf(pi2) - Marshal.SizeOf(GetType(Int32))
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPrinters > 0 Then
            Marshal.FreeHGlobal(CType(pPrinters, IntPtr))
        End If

    End Sub
#End Region

End Class
#End Region
