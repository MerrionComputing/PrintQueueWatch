'\\ --[PrintJob]----------------------------------------
'\\ Class wrapper for the windows API calls and constants
'\\ relating to individual jobs in a printer queue..
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing.Printing
Imports System.Text

Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports PrinterQueueWatch.PrinterMonitoringExceptions
Imports System.Runtime.Remoting.Activation


#Region "PrintJob class"

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintJob
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents the properties of a single print job queued against a print device
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	20/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintJob
    Implements IDisposable

#Region "Tracing"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Print Job specific tracing switch
    ''' </summary>
    ''' <remarks>Add a trace flag named "PrintJob" in the application .config file to 
    ''' trace print job related processes
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared TraceSwitch As New TraceSwitch("PrintJob", "Printer Monitor Component Print Job Tracing")
#End Region

#Region "Private member variables"

    Private mhPrinter As IntPtr
    Private midJob As Int32

    Private bHandleOwnedByMe As Boolean

    Private mPrinterName As String
    Private mDocument As String

    Private ji1 As New JOB_INFO_1
    Private ji2 As New JOB_INFO_2

    Private _TimeWindow As New TimeWindow

    Private _PositionChanged As Boolean
    Private _Changed_Ji1 As Boolean
    Private _Changed_Ji2 As Boolean

    Private _UrlString As String

    Private _Populated As Boolean = False

#End Region

#Region "Public interface"

#Region "JobId"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The unique identifier of the print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This id is only unique for the printer
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property JobId() As Int32
        Get
            Return midJob
        End Get
    End Property
#End Region

#Region "JOB_INFO_1 properties"

#Region "PrinterName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print device that this job is queued against
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the printer whenever a new job is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PrinterName)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PrinterName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If ji1.pPrinterName Is Nothing Then
                Return ""
            Else
                Return ji1.pPrinterName
            End If
        End Get
    End Property
    Friend WriteOnly Property InitPrinterName() As String
        Set(ByVal Value As String)
            ji1.pPrinterName = Value
        End Set
    End Property
#End Region

#Region "UserName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the user that sent the print job for printing
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the user that originated the job whenever a new job is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.Username)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property UserName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If ji1.pUserName Is Nothing Then
                Return ""
            Else
                Return ji1.pUserName
            End If
        End Get
        Set(ByVal Value As String)
            ji1.pUserName = Value
            _Changed_Ji1 = True
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Username changed to " & Value, Me.GetType.ToString)
            End If
        End Set
    End Property
    Friend WriteOnly Property InitUsername() As String
        Set(ByVal Value As String)
            ji1.pUserName = Value
        End Set
    End Property
#End Region

#Region "MachineName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the workstation that sent the print job to print
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the name of the machine that originated a job whenever a new job is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.MachineName)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property MachineName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If ji1.pMachineName Is Nothing Then
                Return ""
            Else
                Return ji1.pMachineName
            End If
        End Get
    End Property
    Friend WriteOnly Property InitMachineName() As String
        Set(ByVal Value As String)
            ji1.pMachineName = Value
        End Set
    End Property
#End Region

#Region "Document"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The document name being printed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is set by the application which sends the job to be printed.  Many 
    ''' applications put the application name at the start of the document name to aid 
    ''' identification
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
    Public Overridable Property Document() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If ji1.pDocument Is Nothing Then
                Return ""
            Else
                Return ji1.pDocument
            End If
        End Get
        Set(ByVal Value As String)
            ji1.pDocument = Value
            _Changed_Ji1 = True
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                Trace.WriteLine("Document name changed to " & Value, Me.GetType.ToString)
            End If
        End Set
    End Property
    Friend WriteOnly Property InitDocument() As String
        Set(ByVal Value As String)
            ji1.pDocument = Value
        End Set
    End Property
#End Region

#Region "StatusDescription"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The description of the current status of the print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the status of a job when it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.StatusDescription)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property StatusDescription() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If (ji1.pStatus Is Nothing) OrElse (ji1.pStatus = "") Then
                Return DerivedStatusDescription()
            Else
                Return ji1.pStatus
            End If
        End Get
    End Property
    Friend WriteOnly Property InitStatusDescription() As String
        Set(ByVal Value As String)
            ji1.pStatus = Value
        End Set
    End Property
#End Region

#Region "DataType"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the data type that is used for this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This can be RAW or EMF.  if the data type is RAW then the spool file contains 
    ''' a printer control language such as PCL or PostScript
    ''' </remarks>
    ''' <example>Prints the data type of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.DataType)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property DataType() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            If ji1.pDatatype Is Nothing Then
                Return ""
            Else
                Return ji1.pDatatype
            End If
        End Get
    End Property
    Friend WriteOnly Property InitDataType() As String
        Set(ByVal Value As String)
            ji1.pDatatype = Value
        End Set
    End Property
#End Region

#Region "Status (Private)"
    Private ReadOnly Property Status() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            Return ji1.Status
        End Get
    End Property
#End Region

#Region "PagesPrinted"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of pages that have been printed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the number of pages in each job as it is deleted
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''   Private Sub pmon_JobDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobDeleted
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PagesPrinted.ToString)
    '''    End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PagesPrinted() As Int32
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            Return ji1.PagesPrinted
        End Get
    End Property
    Friend WriteOnly Property InitPagesPrinted() As Int32
        Set(ByVal Value As Int32)
            ji1.PagesPrinted = Value
        End Set
    End Property
#End Region

#Region "Position"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The position of the job in the print device job queue
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the position in the queue of each new job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.Position.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property Position() As Int32
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            Return ji1.Position
        End Get
        Set(ByVal Value As Int32)
            If Value > 0 Then
                ji1.Position = Value
                _PositionChanged = True
            End If
        End Set
    End Property
    Friend WriteOnly Property InitPosition() As Int32
        Set(ByVal Value As Int32)
            ji1.Position = Value
        End Set
    End Property
#End Region

#Region "Update"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the latest state of the print job from the print device spool queue
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Update()
        Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, True)
        Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, True)
        _PositionChanged = False
        _Changed_Ji1 = False
        _Changed_Ji2 = False
        _TimeWindow.Changed = False
    End Sub
#End Region
#Region "Commit"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update the print spool with changes made to this print job class
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Changes the user name for each new job as it is added to the monitored 
    ''' printers and commits this to the print queue
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        .PrintJob.Username = "New user name"
    '''        .PrintJob.Commit
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Commit()
        If _Changed_Ji1 Or _TimeWindow.Changed Then
            Call SaveJobInfo(JobInfoLevels.JobInfoLevel1)
            _Changed_Ji1 = False
            _TimeWindow.Changed = False
        End If
        If _Changed_Ji2 Or _TimeWindow.Changed Then
            Call SaveJobInfo(JobInfoLevels.JobInfoLevel2)
            _Changed_Ji2 = False
            _TimeWindow.Changed = False
        End If
    End Sub
#End Region

#End Region

#Region "JOB_INFO_2 derived properties"

#Region "TotalPages"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The total number of pages in this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the number of pages in each new job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.TotalPages.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property TotalPages() As Int32
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel1, False)
            Return ji2.TotalPage
        End Get
    End Property
    Friend WriteOnly Property InitTotalPages() As Int32
        Set(ByVal Value As Int32)
            If Value > 0 Then
                ji2.TotalPage = Value
            End If
        End Set
    End Property
#End Region

#Region "PaperKind"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The paper type that the job is intended to be printed on
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This could be a standard paper size (A4, A5 etc) or custom paper size if the printer 
    ''' supports this
    ''' </remarks>
    ''' <example>Prints the paper type of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PaperKind.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PaperKind() As PaperKind
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            With ji2.DeviceMode
                If .dmPaperSize > 118 Or .dmPaperSize < 0 Then
                    Return PaperKind.Custom
                Else
                    Return CType(.dmPaperSize, PaperKind)
                End If
            End With
        End Get
    End Property
#End Region

#Region "PaperWidth"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The width of the selected paper (if PaperKind is PaperKind.Custom)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is measured in millimeters
    ''' </remarks>
    ''' <example>Prints the paper width of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PaperWidth.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PaperWidth() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return ji2.DeviceMode.dmPaperWidth
        End Get
    End Property
#End Region

#Region "PaperLength"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The height of the selected paper (if PaperKind is PaperKind.Custom)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is measured in millimeters
    ''' </remarks>
    ''' <example>Prints the paper length of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PaperLength.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PaperLength() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return ji2.DeviceMode.dmPaperLength
        End Get
    End Property
#End Region

#Region "Landscape"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is to be printed in landscape mode
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <example>Prints the landacape mode of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(IIf(.PrintJob.Landsape,"Landscape", "Portrait"))
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property Landscape() As Boolean
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return (ji2.DeviceMode.dmOrientation = DeviceOrientations.DMORIENT_LANDSCAPE)
        End Get
    End Property
#End Region

#Region "Color"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is to be printed in colour
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This will be true if the setting is set to rpint in colour even if
    ''' the actual document has no colour elements 
    ''' </remarks>
    ''' <example>Prints the colour / monochrome setting of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(Iif(.PrintJob.Color,"Colour", "Monochrome"))
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property Color() As Boolean
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return (ji2.DeviceMode.dmColor = DeviceColourModes.DMCOLOR_COLOR)
        End Get
    End Property
#End Region

#Region "PaperSource"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The input source (tray or bin) requested for the print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' 
    ''' </remarks>
    ''' <example>Prints the paper source of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PaperSource.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PaperSource() As PaperSourceKind
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            With ji2.DeviceMode
                If .dmDefaultSource >= PrinterPaperBins.DMBIN_USER Or .dmDefaultSource < 0 Then
                    Return PaperSourceKind.Custom
                Else
                    Return CType(ji2.DeviceMode.dmDefaultSource, PaperSourceKind)
                End If
            End With
        End Get
    End Property
#End Region

#Region "PrinterResolutionKind"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The resolution to use for the print document
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Can be draft, low, medium or high quality or custom quality
    ''' </remarks>
    ''' <example>Prints the print resolution of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.PrinterResolutionKind)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PrinterResolutionKind() As PrinterResolutionKind
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.DeviceMode.dmPrintQuality = DeviceModeResolutions.DMRES_DRAFT Then
                Return PrinterResolutionKind.Draft
            ElseIf ji2.DeviceMode.dmPrintQuality = DeviceModeResolutions.DMRES_HIGH Then
                Return PrinterResolutionKind.High
            ElseIf ji2.DeviceMode.dmPrintQuality = DeviceModeResolutions.DMRES_LOW Then
                Return PrinterResolutionKind.Low
            ElseIf ji2.DeviceMode.dmPrintQuality = DeviceModeResolutions.DMRES_MEDIUM Then
                Return PrinterResolutionKind.Medium
            Else
                Return PrinterResolutionKind.Custom
            End If
        End Get
    End Property
#End Region

#Region "PrinterResolutionX"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The printer resolution in the horizontal dimension 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is set if PrinterResolutionKind is PrinterResolutionKind.Custom
    ''' This is measured in dots per inch
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PrinterResolutionX() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return ji2.DeviceMode.dmPrintQuality
        End Get
    End Property
#End Region

#Region "PrinterResolutionY"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The printer resolution in the vertical dimension 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This valuse is set if PrinterResolutionKind is PrinterResolutionKind.Custom.
    ''' This is measured in dots per inch
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PrinterResolutionY() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return ji2.DeviceMode.dmYResolution
        End Get
    End Property
#End Region

#Region "Copies"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of copies of each page to print
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Some applications misreport the number of copies to the spooler which will 
    ''' result in incorrect values being returned
    ''' </remarks>
    ''' <example>Prints the number of copies of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.Copies.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property Copies() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.DeviceMode.dmCopies > 0 Then
                Return ji2.DeviceMode.dmCopies
            Else
                If ji2.PagesPrinted > ji2.TotalPage Then
                    Return ji2.PagesPrinted \ ji2.TotalPage
                Else
                    Return 1
                End If
            End If
        End Get
    End Property
#End Region

#Region "NotifyUserName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The user to notify about the progress of this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This should be set to the network login name of the user 
    ''' </remarks>
    ''' <example>Changes the notify user of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        .PrintJob.NotifyUserName = "Administrator"
    '''        .PrintJob.Commit
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property NotifyUserName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.pNotifyName Is Nothing Then
                Return ""
            Else
                Return ji2.pNotifyName
            End If
        End Get
        Set(ByVal Value As String)
            ji2.pNotifyName = Value
            _Changed_Ji2 = True
        End Set
    End Property
    Friend WriteOnly Property InitNotifyUsername() As String
        Set(ByVal Value As String)
            ji2.pNotifyName = Value
        End Set
    End Property
#End Region

#Region "PrintProcessorName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print processor 
    ''' which is responsible for printing this job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property PrintProcessorName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.pPrintProcessor Is Nothing Then
                Return ""
            Else
                Return ji2.pPrintProcessor
            End If
        End Get
    End Property
    Friend WriteOnly Property InitPrintProcessorName() As String
        Set(ByVal Value As String)
            ji2.pPrintProcessor = Value
        End Set
    End Property
#End Region

#Region "Drivername"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the printer driver
    ''' that is responsible for producing this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property DriverName() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.pDriverName Is Nothing Then
                Return ""
            Else
                Return ji2.pDriverName
            End If
        End Get
    End Property
    Friend WriteOnly Property InitDrivername() As String
        Set(ByVal Value As String)
            ji2.pDriverName = Value
        End Set
    End Property
#End Region

#Region "Priority"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The priority of this print job.  Higher priority jobs will be processed ahead
    ''' of lower priority jobs
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Valid values are in the range of 1 (lowest priority) to 99 (highest priority)
    ''' </remarks>
    ''' <example>Prints the priority of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.Priority.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property Priority() As Int32
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return ji2.Priority
        End Get
        Set(ByVal Value As Int32)
            If Value > 99 Then
                '\\ Priority cannot exceed 99
                ji2.Priority = 99
            ElseIf Value < 1 Then
                '\\ Priority cannot be less than 1
                ji2.Priority = 1
            Else
                ji2.Priority = Value
            End If
            _Changed_Ji2 = True
        End Set
    End Property
    Friend WriteOnly Property InitPriority() As Int32
        Set(ByVal Value As Int32)
            ji2.Priority = Value
        End Set
    End Property
#End Region

#Region "Parameters"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Extra driver specific parameters for this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The acceptable parameters and values depend on the print driver being used to 
    ''' print this print job
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable ReadOnly Property Parameters() As String
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.pParameters Is Nothing Then
                Return ""
            Else
                Return ji2.pParameters
            End If
        End Get
    End Property
    Friend WriteOnly Property InitParameters() As String
        Set(ByVal Value As String)
            ji2.pParameters = Value
        End Set
    End Property
#End Region

#Region "Submitted"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the date and time at which the job was submitted for printing
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The time value is returned in the local time of the machine on which the PrintQueueWatch
    ''' component is installed 
    ''' </remarks>
    ''' <example>Prints the date and time submitted of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.Submitted.ToString)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Specifies the date and time at which the job was submitted for printing")> _
    Public Overridable ReadOnly Property Submitted() As DateTime
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            '\\ Submitted is held in FileTime so needs to be converted to the localised time to take account of daylight saving etc.
            Return ji2.Submitted.ToDateTime.ToLocalTime
        End Get
    End Property
#End Region

#Region "TimeWindow"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the earliest time and latest times that the job can be printed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' See the TimeWindow class for 
    ''' details of the settings of this class
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Specifies the earliest time and latest times that the job can be printed")> _
    Public Overridable Property TimeWindow() As TimeWindow
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            _TimeWindow = New TimeWindow(ji2.StartTime, ji2.UntilTime)
            Return _TimeWindow
        End Get
        Set(ByVal Value As TimeWindow)
            ji2.StartTime = Value.StartTimeMinutes
            ji2.UntilTime = Value.EndTimeMinutes
            _Changed_Ji2 = True
        End Set
    End Property
#End Region

#Region "QueuedTime"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the total time, in milliseconds, that has elapsed since the job began printing
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Specifies the total time, in milliseconds, that has elapsed since the job began printing")> _
    Public Overridable ReadOnly Property QueuedTime() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, True)
            Return ji2.Time
        End Get
    End Property
#End Region

#Region "JobSize"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the size, in bytes, of the job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' While the job is being spooled this will contain the current size of the spool file
    ''' </remarks>
    ''' <example>Prints the job size of each job as it is added to the monitored printers
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs)
    '''        Trace.WriteLine(.PrintJob.DataType)
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Specifies the size, in bytes, of the job")> _
    Public Overridable ReadOnly Property JobSize() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, True)
            Return ji2.JobSize
        End Get
    End Property
#End Region

#Region "LogicalPagesPerPhysicalPage"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of logical pages for each physical page
    ''' </summary>
    ''' <value></value>
    ''' <remarks>This value should be 1, 2, 4, 8 or 16
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	11/02/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LogicalPagesPerPhysicalPage() As Integer
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            If ji2.DeviceMode.dmNup > 0 Then
                Return ji2.DeviceMode.dmNup
            Else
                Return 1
            End If
        End Get
    End Property
#End Region

#Region "DefaultPaperSource"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The paper tray (or input bin) to use for this print job
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	11/02/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DefaultPaperSource() As Drawing.Printing.PaperSourceKind
        Get
            Call RefreshJobInfo(JobInfoLevels.JobInfoLevel2, False)
            Return CType(ji2.DeviceMode.dmDefaultSource, Drawing.Printing.PaperSourceKind)
        End Get
    End Property
#End Region

#End Region

#Region "Cancel"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels this print job and removes it from the spool queue
    ''' </summary>
    ''' <remarks>
    ''' Only the originator of a print job or a user with administrator rights on the
    ''' print device may cancel the job
    ''' </remarks>
    ''' <exception cref="InsufficentPrintJobAccessRightsException">
    ''' Thrown if the user has no access rights to delete this job
    ''' </exception>
    ''' <example>Cancels any jobs that have more than 8 copies 
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs).PrintJob
    '''        If .Copies > 8 Then
    '''           .Cancel
    '''        End If
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Cancel()

        If System.Environment.OSVersion.Version.Major < 4 Then   '\\ For systems less than windows NT 4...
            If Not SetJob(mhPrinter, midJob, 0, New IntPtr(0), PrintJobControlCommands.JOB_CONTROL_CANCEL) Then
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("SetJob (Cancel) failed", Me.GetType.ToString)
                End If
                Throw New InsufficentPrintJobAccessRightsException(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pjerr_cancel"), New Win32Exception)
            Else
                If PrintJob.TraceSwitch.TraceVerbose Then
                    Trace.WriteLine("SetJob (Cancel) succeeded", Me.GetType.ToString)
                End If
            End If
        Else
            If Not SetJob(mhPrinter, midJob, 0, New IntPtr(0), PrintJobControlCommands.JOB_CONTROL_CANCEL) Then
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("SetJob (Cancel) failed", Me.GetType.ToString)
                End If
                Throw New InsufficentPrintJobAccessRightsException(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pjerr_cancel"), New Win32Exception)
            Else
                If PrintJob.TraceSwitch.TraceVerbose Then
                    Trace.WriteLine("SetJob (Cancel) succeeded", Me.GetType.ToString)
                End If
            End If
        End If

    End Sub
#End Region

#Region "Delete"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes this print job
    ''' </summary>
    ''' <remarks>
    ''' Only the originator of a print job or a user with administrator rights on the
    ''' print device may delete it
    ''' </remarks>    
    ''' <exception cref="InsufficentPrintJobAccessRightsException">
    ''' Thrown if the user has no access rights to delete this job
    ''' </exception>
    ''' <example>Cancels any jobs that have more than 8 copies 
    ''' <code>
    '''   Private WithEvents pmon As New PrinterMonitorComponent
    '''
    '''   pmon.AddPrinter("Microsoft Office Document Image Writer")
    '''   pmon.AddPrinter("HP Laserjet 5")
    ''' 
    '''    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.JobAdded
    '''
    '''    With CType(e, PrintJobEventArgs).PrintJob
    '''        If .Copies > 8 Then
    '''           .Delete
    '''        End If
    '''     End With
    '''
    '''  End Sub
    ''' </code>
    ''' </example>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Delete()

        If Not SetJob(mhPrinter, midJob, 0, New IntPtr(0), PrintJobControlCommands.JOB_CONTROL_DELETE) Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("SetJob (Delete) failed", Me.GetType.ToString)
            End If
            Throw New InsufficentPrintJobAccessRightsException(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pjerr_delete"), New Win32Exception)
        Else
            If PrintJob.TraceSwitch.TraceVerbose Then
                Trace.WriteLine("SetJob (Delete) succeeded", Me.GetType.ToString)
            End If
        End If

    End Sub
#End Region

#Region "Print job status settings "
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the print job is paused
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <exception cref="System.ComponentModel.Win32Exception">
    ''' Thrown if the job does not exist or the user has no access rights to pause it
    ''' </exception>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Paused() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_PAUSED) = PrintJobStatuses.JOB_STATUS_PAUSED)
        End Get
        Set(ByVal Value As Boolean)
            If Not Value.Equals(Me.Paused) Then
                '\\ The paused state has changed: Call the pause or resume command as appropriate
                If Value Then
                    If Not SetJob(mhPrinter, midJob, 0, IntPtr.Zero, PrintJobControlCommands.JOB_CONTROL_PAUSE) Then
                        Throw New Win32Exception
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("SetJob (Cancel) failed", Me.GetType.ToString)
                        End If
                    Else
                        If PrintJob.TraceSwitch.TraceVerbose Then
                            Trace.WriteLine("SetJob (Pause) succeeded", Me.GetType.ToString)
                        End If
                    End If
                Else
                    If Not SetJob(mhPrinter, midJob, 0, IntPtr.Zero, PrintJobControlCommands.JOB_CONTROL_RESUME) Then
                        Throw New Win32Exception
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("SetJob (Resume) failed", Me.GetType.ToString)
                        End If
                    Else
                        If PrintJob.TraceSwitch.TraceVerbose Then
                            Trace.WriteLine("SetJob (Resume) succeeded", Me.GetType.ToString)
                        End If
                    End If
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job has been deleted
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Deleted() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_DELETED) = PrintJobStatuses.JOB_STATUS_DELETED)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is being deleted
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Deleting() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_DELETING) = PrintJobStatuses.JOB_STATUS_DELETING)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the job has been printed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This will be true once the job has been completely sent to the printer.  This 
    ''' does not mean that the physical print out has necessarily appeared.
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Printed() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_PRINTED) = PrintJobStatuses.JOB_STATUS_PRINTED)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is currently printing
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Printing() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_PRINTING) = PrintJobStatuses.JOB_STATUS_PRINTING)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if there is an error with this print job that prevents it from 
    ''' printing
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property InError() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_ERROR) = PrintJobStatuses.JOB_STATUS_ERROR)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the job is in error because the printer is off line
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Offline() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_OFFLINE) = PrintJobStatuses.JOB_STATUS_OFFLINE)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the job is in error because the printer has run out of paper
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PaperOut() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_PAPEROUT) = PrintJobStatuses.JOB_STATUS_PAPEROUT)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is in error because the printer requires user intervention
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This can be caused by a print job that requires manual paper feed
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property UserInterventionRequired() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_USER_INTERVENTION) = PrintJobStatuses.JOB_STATUS_USER_INTERVENTION)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the print job is spooling to a spool file
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Spooling() As Boolean
        Get
            Return ((Status And PrintJobStatuses.JOB_STATUS_SPOOLING) = PrintJobStatuses.JOB_STATUS_SPOOLING)
        End Get
    End Property


    Friend WriteOnly Property InitStatus() As Int32
        Set(ByVal Value As Int32)
            ji1.Status = Value
        End Set
    End Property
#End Region

#Region "Transfer"
    'TODO: Transfer fails to open the print job consistently - this will need to be 
    'raised as an MS support incident before this can be achieved
#Const TRANSFER_SUPPORTED = True

#If TRANSFER_SUPPORTED Then

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Transfers this print job to the named printer 
    ''' </summary>
    ''' <param name="NewPrinter">Name of the new printer to move the job to</param>
    ''' <param name="RemoveLocal">True to remove this copy of the job</param>
    ''' <remarks>If the DataType is RAW then the target printer may not print the job if it does not
    ''' support the printer control language that the job contains
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	05/12/2005	Created
    ''' </history>
    ''' ----------------------------------------------------------------------------- 
    Public Sub Transfer(ByVal NewPrinter As String, ByVal RemoveLocal As Boolean)



        Dim NewDoc As New DOC_INFO_1

        With NewDoc
            .DocumentName = Me.Document
            .DataType = Me.DataType
            .OutputFilename = String.Empty
        End With

        Dim phPrinter As Integer
        Dim pDefault As New PRINTER_DEFAULTS(PrinterAccessRights.PRINTER_ACCESS_USE)

        If OpenPrinter(Me.UniquePrinterObject, phPrinter, pDefault) Then
            If phPrinter <> 0 Then
                'Read this print job into a bit of memory....
                Dim ptBuf As IntPtr
                Try
                    ptBuf = Marshal.AllocHGlobal(Me.JobSize)
                Catch exMem As OutOfMemoryException
                    Throw New PrintJobTransferException("Print job is too large", exMem)
                    Exit Sub
                End Try

                '\\ Read the print job in to memory
                Dim pcbneeded As Integer
                If Not ReadPrinter(phPrinter, ptBuf, Me.JobSize, pcbneeded) Then
                    Throw New PrintJobTransferException("Failed to read the print job", New Win32Exception)
                    Exit Sub
                End If

                Dim DataFile As New PrinterDataFile(ptBuf, Me.DataType)

                'Open the target printer
                Dim phPrinterTarget As Integer
                If OpenPrinter(NewPrinter, phPrinterTarget, pDefault) Then
                    'Start the new document
                    If StartDocPrinter(phPrinterTarget, 1, NewDoc) Then
                        'Print each page in the print job...
                        For CurrentPage As Integer = 1 To DataFile.TotalPages
                            '
                            'Print this page
                            If Not WritePrinter(New IntPtr(phPrinterTarget), ptBuf, Me.JobSize, pcbneeded) Then
                                Throw New PrintJobTransferException("Failed to write the print job", New Win32Exception)
                                Exit Sub
                            End If
                            'Notify the spooler that the page is finished
                            Call EndPagePrinter(phPrinterTarget)
                        Next
                        EndDocPrinter(phPrinterTarget)
                    Else
                        Throw New PrintJobTransferException("Failed to write the print job", New Win32Exception)
                        Exit Sub
                    End If

                    Call ClosePrinter(phPrinterTarget)
                End If

                'Free this buffer again
                Marshal.FreeHGlobal(ptBuf)
            Else
                    Throw New InsufficentPrinterAccessRightsException("Could not read the print job")
            End If
        Else
            Throw New Win32Exception
        End If

    End Sub
#End If
#End Region

#Region "PrinterHandle"
    Private ReadOnly Property PrinterHandle() As Int32
        Get
            Dim pDef As New PrinterDefaults

            pDef.DesiredAccess = PrinterAccessRights.PRINTER_ACCESS_USE

            If mhPrinter.ToInt32 = 0 Then
                If Not OpenPrinter(PrinterName, mhPrinter, pDef) Then
                    Throw New Win32Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("OpenPrinter() failed", Me.GetType.ToString)
                    End If
                Else
                    bHandleOwnedByMe = True
                End If
            End If
            Return mhPrinter.ToInt32
        End Get
    End Property
#End Region

#Region "DerivedStatusDescription"
    Private Function DerivedStatusDescription() As String

        Dim TempDescription As New StringBuilder

        If Paused Then
            TempDescription.Append(My.Resources.jsd_paused)
        End If

        If InError Then
            TempDescription.Append(My.Resources.jsd_error)
            If PaperOut Then
                TempDescription.Append(My.Resources.jsd_paperout)
            End If
        End If

        If Offline Then
            TempDescription.Append(My.Resources.jsd_offline)
        End If

        If Printed Then
            TempDescription.Append(My.Resources.jsd_printed)
        End If

        If Printing Then
            TempDescription.Append(My.Resources.jsd_printing)
        End If

        Return TempDescription.ToString

    End Function
#End Region

#Region "InitJobInfo"
    Private Sub InitJobInfo()
        '\\ Get the first cut of the job info..
        _Populated = True
        Try
            ji1 = New JOB_INFO_1(mhPrinter, midJob)
        Catch e32 As Win32Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine(e32.ToString, Me.GetType.ToString)
            End If
            _Populated = False
        End Try
        Try
            ji2 = New JOB_INFO_2(mhPrinter, midJob)
        Catch e32_2 As Win32Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine(e32_2.ToString, Me.GetType.ToString)
            End If
            _Populated = False
        End Try

    End Sub
#End Region

#Region "Public constructor"
    Friend Sub New(ByVal hPrinter As IntPtr, ByVal idJob As Int32)

        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & hPrinter.ToString & "," & idJob.ToString & ")", Me.GetType.ToString)
        End If

        '\\ Take a local copy of the printer handle and job id
        mhPrinter = hPrinter
        bHandleOwnedByMe = False
        midJob = idJob
        Call InitJobInfo()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the job identified by JobId queued against the device named
    ''' </summary>
    ''' <param name="DeviceName">The name of the device the job is queued against</param>
    ''' <param name="idJob">The unique job identifier</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal DeviceName As String, ByVal idJob As Int32)
        Dim hPrinter As Integer
        bHandleOwnedByMe = True
        If OpenPrinter(DeviceName, hPrinter, 0) Then
            mhPrinter = New IntPtr(hPrinter)
            midJob = idJob
            Call InitJobInfo()
        Else
            Throw New Win32Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("OpenPrinter() failed", Me.GetType.ToString)
            End If
        End If
    End Sub

    Public Sub New()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Frees up system resources used by this PrintJob class
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Overloads Sub Dispose() Implements IDisposable.Dispose
        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("Dispose()", Me.GetType.ToString)
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If bHandleOwnedByMe Then
                If Not ClosePrinter(mhPrinter) Then
                    Throw New Win32Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("Error in PrinterInformation:Dispose")
                    End If
                End If
            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region

#Region "RefreshJobInfo"
    Private Sub RefreshJobInfo(ByVal Index As JobInfoLevels, ByVal ForceReload As Boolean)

        Dim ji1Temp As JOB_INFO_1
        Dim ji2Temp As JOB_INFO_2


        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("RefreshJobInfo(" & Index.ToString & ")", Me.GetType.ToString)
        End If

        Select Case Index
            Case JobInfoLevels.JobInfoLevel3

            Case JobInfoLevels.JobInfoLevel2
                If ForceReload OrElse (ji2 Is Nothing) Then
                    Try
                        ji2Temp = New JOB_INFO_2(mhPrinter, midJob)
                    Catch e As Win32Exception
                        Exit Sub
                    End Try
                    ji2 = ji2Temp
                End If
            Case Else
                If ForceReload OrElse (ji1 Is Nothing) Then
                    Try
                        ji1Temp = New JOB_INFO_1(mhPrinter, midJob)
                    Catch e As Win32Exception
                        Exit Sub
                    End Try
                    ji1 = ji1Temp
                End If
        End Select

    End Sub
#End Region

#Region "SaveJobInfo"
    Private Sub SaveJobInfo(ByVal Index As JobInfoLevels)
        Const JOB_POSITION_UNSPECIFIED As Integer = 0

        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("SaveJobInfo(" & Index.ToString & ")", Me.GetType.ToString)
        End If

        Select Case Index
            Case JobInfoLevels.JobInfoLevel3

            Case JobInfoLevels.JobInfoLevel2
                '\\ Update the JOB_INFO_2 structure held in the spool file
                If Not _PositionChanged Then
                    ji2.Position = JOB_POSITION_UNSPECIFIED
                End If
                If Not SetJob(mhPrinter, midJob, JobInfoLevels.JobInfoLevel2, ji2, PrintJobControlCommands.JOB_CONTROL_SETJOB) Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("SetJob() failed", Me.GetType.ToString)
                    End If
                    Throw New InsufficentPrintJobAccessRightsException(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pjerr_update"), New Win32Exception)
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                        Trace.WriteLine("Job Info (2) saved ", Me.GetType.ToString)
                    End If
                    _Changed_Ji2 = False
                End If

            Case Else
                '\\ Update the JOB_INFO_1 structure held in the spool file
                If Not _PositionChanged Then
                    ji1.Position = JOB_POSITION_UNSPECIFIED
                End If
                If Not SetJob(mhPrinter, midJob, Index, ji1, PrintJobControlCommands.JOB_CONTROL_SETJOB) Then
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("SetJob() failed", Me.GetType.ToString)
                    End If
                    Throw New InsufficentPrintJobAccessRightsException(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pjerr_update"), New Win32Exception)
                Else
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
                        Trace.WriteLine("Job Info (1) saved ", Me.GetType.ToString)
                    End If
                    _Changed_Ji1 = False
                End If
        End Select
        _PositionChanged = False

        _TimeWindow.Changed = False
    End Sub
#End Region

#End Region

#Region "Friend properties"
#Region "UrlString"
    Friend WriteOnly Property UrlString() As String
        Set(ByVal Value As String)
            _UrlString = Value
        End Set
    End Property
#End Region

#Region "Populated"
    Friend ReadOnly Property Populated() As Boolean
        Get
            Return _Populated
        End Get
    End Property
#End Region

#Region "Refresh"
    '\\--[Refresh]--------------------------------------------------
    '\\ Repopulate the PrintJob from the spooler [API] on demand
    '\\ for the case that it was not succesfully filled on creation
    '\\ ------------------------------------------------------------
    Friend Sub Refresh()
        Call InitJobInfo()
    End Sub
#End Region

#Region "UniquePrinterObject"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the unique name of this job which can be passed to get a handle to it
    ''' </summary>
    ''' <returns>[PrinterNmae], PrintJob xxxxx </returns>
    ''' <remarks>Used internally for ReadPrinter api call
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	05/12/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Friend Function UniquePrinterObject() As String

        Return Me.PrinterName & ",Job " & Me.JobId.ToString

    End Function
#End Region

#End Region

End Class

#End Region

#Region "Print Job type safe collection"

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintJobCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of print jobs
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
<System.Security.SuppressUnmanagedCodeSecurity()> _
Public Class PrintJobCollection
    Inherits Generic.SortedList(Of Integer, PrintJob)

#Region "Private member variables"
    Private bHandleOwnedByMe As Boolean
    Private hPrinter As IntPtr
#End Region

#Region "JobPendingDeletion"
    '\\ Because the delete event is asynchronous, jobs have to be removed from the
    '\\ job list one in arrears.  This public variable tells which one is to be removed
    '\\ next.
    Friend JobPendingDeletion As Int32
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If bHandleOwnedByMe Then
            If Not ClosePrinter(hPrinter) Then
                Throw New Win32Exception()
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("Error in PrinterInformation:Dispose")
                End If
            End If
        End If
    End Sub
#End Region

#Region "Public interface"

    Friend ReadOnly Property AddOrGetById(ByVal dwJobId As Int32, ByVal mhPrinter As IntPtr) As PrintJob
        Get
            Dim printJobLock As Object = New Object()
            SyncLock printJobLock
                Dim pjThis As PrintJob
                If Not ContainsJobId(dwJobId) Then
                    Try
                        pjThis = New PrintJob(mhPrinter, dwJobId)
                        Me.Add(dwJobId, pjThis)
                    Catch e As Win32Exception
                        If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                            Trace.WriteLine("AddOrGetById(" & dwJobId.ToString & ") failed - " & e.Message & "::" & e.NativeErrorCode, Me.GetType.ToString)
                        End If
                    End Try
                End If
                Return ItemByJobId(dwJobId)
            End SyncLock
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the print job by its unique Job Id
    ''' </summary>
    ''' <param name="dwJob"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ItemByJobId(ByVal dwJob As Int32) As PrintJob
        Get
            Return Me.Item(dwJob)
        End Get
    End Property

    Public Overloads Sub Add(ByVal pjIn As PrintJob)
        Me.Add(pjIn.JobId, pjIn)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns true if this collection contains the given Job Id
    ''' </summary>
    ''' <param name="pjTestId"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ContainsJobId(ByVal pjTestId As Int32) As Boolean
        Get
            Return Me.ContainsKey(pjTestId)
        End Get
    End Property

    <Description("Removes the print job identified by the given job id from the printjobs collection")> _
    Public Sub RemoveByJobId(ByVal pjId As Int32)
        Me.Remove(pjId)
    End Sub



#End Region
#Region "Public constructors"

    Private Sub InitJobList(ByVal mhPrinter As IntPtr, ByVal JobCount As Int32)
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer (in bytes)
        Dim pJobInfo As IntPtr

        '\\ Save the printer handle
        hPrinter = mhPrinter

        '\\ If the current jobcount is unknown, try 255
        If JobCount = 0 Then
            JobCount = 255
        End If

        If Not EnumJobs(mhPrinter, 0, JobCount, JobInfoLevels.JobInfoLevel1, New IntPtr(0), 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pJobInfo = Marshal.AllocHGlobal(pcbNeeded)
                Dim pcbProvided As Int32 = pcbNeeded
                Dim pcbTotalNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
                Dim pcTotalReturned As Int32 '\\ Holds the returned size of the output buffer (in bytes)
                If EnumJobs(mhPrinter, 0, JobCount, JobInfoLevels.JobInfoLevel1, pJobInfo, pcbProvided, pcbTotalNeeded, pcTotalReturned) Then
                    If pcTotalReturned > 0 Then
                        Dim item As Int32
                        Dim pnextJob As IntPtr = pJobInfo
                        For item = 0 To pcTotalReturned - 1
                            Dim jiTemp As New JOB_INFO_1(pnextJob)
                            Call Add(New PrintJob(mhPrinter, jiTemp.JobId))
                            pnextJob = New IntPtr(pnextJob.ToInt32 + 64)
                        Next
                    End If
                Else
                    Throw New Win32Exception()
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("EnumJobs() failed", Me.GetType.ToString)
                    End If
                End If
                Marshal.FreeHGlobal(pJobInfo)
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new empty PrintJobs collection
    ''' </summary>
    ''' <remarks>
    ''' This constructor is not meant for use except by the PrinterQueueWatch component
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Creates a new empty PrintJobs collection")> _
    Public Sub New()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new list and fills it with all the jobs currently on a given printer's queue by printer handle
    ''' </summary>
    ''' <param name="mhPrinter">The handle of the printer to get the print jobs for</param>
    ''' <param name="JobCount">The number of jobs to retrieve</param>
    ''' <remarks>
    ''' If JobCount is less than the number of jobs in teh queue only the first JobCount number will 
    ''' be returned
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Creates a new list and fills it with all the jobs currently on a given printer's queue by printer handle")> _
    Public Sub New(ByVal mhPrinter As IntPtr, ByVal JobCount As Int32)

        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & mhPrinter.ToString & "," & JobCount.ToString & ")", Me.GetType.ToString)
        End If
        Call InitJobList(mhPrinter, JobCount)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new list and fills it with all the jobs currently on a given printer's queue by printer device name
    ''' </summary>
    ''' <param name="DeviceName">The name of the print device to get the jobs for</param>
    ''' <param name="JobCount">The number of print jobs to return</param>
    ''' <remarks>
    ''' If JobCount is less than the number of jobs in the queue only the first JobCount number will 
    ''' be returned
    ''' </remarks>
    ''' <exception cref="System.ComponentModel.Win32Exception">
    ''' Thrown if the print device does not exist or the user has no access rights to retrieve the job queue from it
    ''' </exception>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Creates a new list and fills it with all the jobs currently on a given printer's queue by printer device name ")> _
    Public Sub New(ByVal DeviceName As String, ByVal JobCount As Int32)

        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & DeviceName & "," & JobCount.ToString & ")", Me.GetType.ToString)
        End If

        Dim hPrinter As Integer
        bHandleOwnedByMe = True
        If OpenPrinter(DeviceName, hPrinter, 0) Then
            Call InitJobList(New IntPtr(hPrinter), JobCount)
        Else
            Throw New Win32Exception
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("OpenPrinter() failed", Me.GetType.ToString)
            End If
        End If

    End Sub

#End Region

End Class

#End Region

#Region "TimeWindow class"
'\\ --[TimeWindow]--------------------------------------------------
'\\ Specifies a time window during which an event can be scheduled -
'\\ for example when a print job can be printed
'\\ (c) 2003 Merrion Computing Ltd
'\\ ----------------------------------------------------------------

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : TimeWindow
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Specifies a time window during which an event can be scheduled -
'''  for example when a print job can be printed
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class TimeWindow

#Region "Static functions"

    Public Shared Function LocalTimeToMinutesPastGMT(ByVal LocalTime As DateTime) As Integer

        Dim dtGMT As DateTime = LocalTime.ToUniversalTime
        Return (dtGMT.Hour * 60) + dtGMT.Minute

    End Function

    Public Shared Function MinutesPastGMTToLocalTime(ByVal MinutesPastGMT As Integer) As DateTime

        Return New DateTime(Date.Now.Year, Date.Now.Month, Date.Now.Day, MinutesPastGMT \ 60, MinutesPastGMT Mod 60, 0, 0).ToLocalTime

    End Function

#End Region

#Region "Private members"
    Private _StartTime As Integer 'Minutes past GMT
    Private _EndTime As Integer 'Minutes past GMT
    Private _Changed As Boolean
#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new time window with the given range
    ''' </summary>
    ''' <param name="StartTime">The start of the time range in minutes past midnight GMT</param>
    ''' <param name="EndTime">The end of the time range in minutes past midnight GMT</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal StartTime As Integer, ByVal EndTime As Integer)

        If PrintJob.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & StartTime.ToString & "," & EndTime.ToString & ")", Me.GetType.ToString)
        End If

        If StartTime > EndTime Then
            _StartTime = EndTime
            _EndTime = StartTime
        Else
            _StartTime = StartTime
            _EndTime = EndTime
        End If
        _Changed = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new time window with the given range
    ''' </summary>
    ''' <param name="StartTime">The start of the time range in local time</param>
    ''' <param name="EndTime">The end of the time range in local time</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal StartTime As DateTime, ByVal EndTime As DateTime)
        Me.New(LocalTimeToMinutesPastGMT(StartTime), LocalTimeToMinutesPastGMT(EndTime))
        _Changed = True
    End Sub

    Public Sub New()
        '\\ Initialise to an unrestricted time window
        _StartTime = 0
        _EndTime = 0
        _Changed = True
    End Sub
#End Region

#Region "Public interface"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The time of the start of this time range
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is in the system local time 
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property StartTime() As DateTime
        Get
            Return MinutesPastGMTToLocalTime(_StartTime)
        End Get
        Set(ByVal Value As DateTime)
            _StartTime = LocalTimeToMinutesPastGMT(Value)
            If _StartTime > _EndTime Then
                _EndTime = _StartTime
            End If
            _Changed = True
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The end of the time range
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This time is in local system time
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property EndTime() As DateTime
        Get
            Return MinutesPastGMTToLocalTime(_EndTime)
        End Get
        Set(ByVal Value As DateTime)
            _EndTime = LocalTimeToMinutesPastGMT(Value)
            If _EndTime < _StartTime Then
                _StartTime = _EndTime
            End If
            _Changed = True
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the time range is unrestricted
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' If the time range is unrestricted it is from midnight to midnight
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property Unrestricted() As Boolean
        Get
            Return (_StartTime = 0 And _EndTime = 0)
        End Get
        Set(ByVal Value As Boolean)
            If Value Then
                _StartTime = 0
                _EndTime = 0
            End If
            _Changed = True
        End Set
    End Property

    Friend Property Changed() As Boolean
        Get
            Return _Changed
        End Get
        Set(ByVal value As Boolean)
            _Changed = value
        End Set
    End Property
#End Region

#Region "Hidden interface"
    Friend ReadOnly Property StartTimeMinutes() As Integer
        Get
            Return _StartTime
        End Get
    End Property
    Friend ReadOnly Property EndTimeMinutes() As Integer
        Get
            Return _EndTime
        End Get
    End Property
#End Region

#Region "ToString override"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a text description of the tiem range
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Dim sOut As New System.Text.StringBuilder
        If Not PrinterMonitorComponent.ComponentLocalisationResourceManager Is Nothing Then
            With PrinterMonitorComponent.ComponentLocalisationResourceManager
                '\\ Return the data in this TimeWindow class
                sOut.Append(" ")
                If Me.Unrestricted Then
                    sOut.Append(.GetString("tw_ts_Unrestricted"))
                Else
                    sOut.Append(.GetString("tw_ts_From"))
                    sOut.Append(" ")
                    sOut.Append(Me.StartTime.ToString)
                    sOut.Append(" ")
                    sOut.Append(.GetString("tw_ts_Until"))
                    sOut.Append(" ")
                    sOut.Append(Me.EndTime.ToString)
                End If
            End With
        Else
            sOut.Append(Me.GetType.ToString)
        End If
        Return sOut.ToString
    End Function
#End Region

End Class

#End Region

#Region "Friend helper classes"
#Region "PrinterDataFile"
''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterDataFile
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' An in-memory representation of a print spool file
''' </summary>
''' <remarks>
''' This class is for internal use of the PrintQueueWatch component
''' </remarks>
''' <history>
''' 	[Duncan]	05/12/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Friend Class PrinterDataFile

#Region "Constants"
    '\\ PJL Notes :--
    '\\ Ways of storing the job language
    Private Const PJL_LANGUAGE As String = "@PJL ENTER LANGUAGE"
    Private Const PJL_LANGUAGE_POSTSCRIPT As String = "POSTSCRIPT"
    Private Const PJL_LANGUAGE_PCL As String = "PCL"
    Private Const PJL_LANGUAGE_PCLXL As String = "PCLXL"
    Private Const PJL_LANGUAGE_HPGL As String = "HPGL"
    Private Const RAW_POSTSCRIPT As String = "%!PS-Adobe"
    Private Const RAW_PCL_NUMERIC As String = "%-12345X"

    '\\ Ways of storing the number of copies
    Private Const PJL_SET_COPIES As String = "@PJL SET COPIES"
    Private Const PJL_SET_COPIES_QTY As String = "@PJL SET QTY"
    Private Const PJL_COMMENT_QTY As String = "@PJL COMMENT @@CPY"

    '\\ PostScript notes:----
    '\\ "%%Page" comment marks the start of each printed page
    '\\ "NumCopies" or "#copies" precedes number of copies
    Private Const POSTSCRIPT_COPIES As String = "#copies"
    Private Const POSTSCRIPT_NUMCOPIES As String = "NumCopies"
    '\\ HPGL notes :----
    '\\ "RP" precedes number of copies
    Private Const HPGL_COPIES As String = "RP"

#End Region

#Region "Public enumerated types"
    Public Enum SpoolFileFormats
        EMF = 1
        PCL_5 = 2
        PCL_6 = 3
        Postscript = 4
        HPGL = 5
    End Enum
#End Region

#Region "Private members"
    Private _ptBuf As IntPtr
    Private _DataType As String

    Private _TotalPages As Integer
#End Region

#Region "Public interface"

#Region "TotalPages"
    Public ReadOnly Property TotalPages() As Integer
        Get
            Return _TotalPages
        End Get
    End Property
#End Region

    'todo: get the start and size of each page in memory...


#End Region

#Region "Public constructors"
    Public Sub New(ByVal Buffer As IntPtr, ByVal DataType As String)
        _ptBuf = Buffer
        _DataType = DataType
    End Sub
#End Region

End Class
#End Region

#Region "EMF Spool File"
''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : EMF_SpoolFile
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Helper class for EMF format spool file
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	12/12/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Friend Class EMF_SpoolFile

#Region "Private helper classes"
#Region "EMFPAGE"
    Private Class EMFPAGE

#Region "Private members"
        Private _StartOffset As Integer
        Private _EndOffset As Integer
        Private _BasePtr As IntPtr

        Private _Header As EMFMETAHEADER
#End Region

#Region "Public Interface"
        Public ReadOnly Property StartOffset() As Integer
            Get
                Return _StartOffset
            End Get
        End Property

        Public ReadOnly Property EndOffset() As Integer
            Get
                Return _EndOffset
            End Get
        End Property
#End Region

#Region "Public constructor"
        Public Sub New(ByVal MemPtr As IntPtr, ByVal Start As Integer)

            _BasePtr = MemPtr
            _StartOffset = Start
            _Header = New EMFMETAHEADER(MemPtr, Start)
            _EndOffset = Start + _Header.FileSize

        End Sub
#End Region
    End Class
#End Region
#Region "EMFMETAHEADER"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode), System.Security.SuppressUnmanagedCodeSecurity()> _
     Private Class EMFMETAHEADER

#Region "Properties"
        '\\ EMR record header
        Private iType As Int32
        Private nSize As Int32
        '\\ Boundary rectangle
        Private rclBounds_Left As Int32
        Private rclBounds_Top As Int32
        Private rclBounds_Right As Int32
        Private rclBounds_Bottom As Int32
        '\\ Frame rectangle
        Private _rclFrame_Left As Int32
        Private _rclFrame_Top As Int32
        Private _rclFrame_Right As Int32
        Private _rclFrame_Bottom As Int32
        '\\ "Signature"
        Private _signature_1 As Byte
        Private _signature_2 As Byte 'E
        Private _signature_3 As Byte 'M
        Private _signature_4 As Byte 'F
        '\\ nVersion
        Private _nVersion As Int32
        Private _nBytes As Int32
        Private _nRecords As Int32
        Private _nHandles As Int32
        Private _sReversed As Int16
        Private _nDescription As Int16
        Private _offDescription As Int32
        Private _nPalEntries As Int32
        Private _szlDeviceWidth As Int32
        Private _szlDeviceHeight As Int32
        Private _szlDeviceWidthMilimeters As Int32
        Private _szlDeviceHeightMilimeters As Int32
        Private _cbPixelFormat As Int32
        Private _offPixelFormat As Int32
        Private _bOpenGL As Boolean
        Private _szlMicrometersWidth As Int32
        Private _szlMicrometersHeight As Int32
        Private _Description As String

#End Region

#Region "Public properties"
        Public ReadOnly Property BoundaryRect() As Rectangle
            Get
                Return New Rectangle(rclBounds_Left, rclBounds_Top, rclBounds_Right, rclBounds_Bottom)
            End Get
        End Property

        Public ReadOnly Property FrameRect() As Rectangle
            Get
                Return New Rectangle(_rclFrame_Left, _rclFrame_Top, _rclFrame_Right, _rclFrame_Bottom)
            End Get
        End Property

        Public ReadOnly Property Size() As Integer
            Get
                Return nSize
            End Get
        End Property

        Public ReadOnly Property RecordCount() As Integer
            Get
                Return _nRecords
            End Get
        End Property

        Public ReadOnly Property FileSize() As Integer
            Get
                Return _nBytes
            End Get
        End Property

#End Region

#Region "Public constructor"
        Public Sub New(ByVal MemPtr As IntPtr, ByVal Start As Integer)

            iType = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            nSize = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            '\\ Boundary rectangle
            rclBounds_Left = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            rclBounds_Top = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            rclBounds_Right = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            rclBounds_Bottom = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            '\\ Frame rectangle
            _rclFrame_Left = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _rclFrame_Top = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _rclFrame_Right = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _rclFrame_Bottom = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            '\\ "Signature"
            _signature_1 = Marshal.ReadByte(MemPtr, Start)
            Start += 1
            _signature_2 = Marshal.ReadByte(MemPtr, Start)
            Start += 1
            _signature_3 = Marshal.ReadByte(MemPtr, Start)
            Start += 1
            _signature_4 = Marshal.ReadByte(MemPtr, Start)
            Start += 1
            '\\ nVersion
            _nVersion = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _nBytes = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _nRecords = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _nHandles = Marshal.ReadInt16(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int16))
            _sReversed = Marshal.ReadInt16(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int16))
            _nDescription = Marshal.ReadInt16(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int16))
            _offDescription = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _nPalEntries = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _szlDeviceWidth = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _szlDeviceHeight = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _szlDeviceWidthMilimeters = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            _szlDeviceHeightMilimeters = Marshal.ReadInt32(MemPtr, Start)
            Start += Marshal.SizeOf(GetType(Int32))
            If nSize > Start Then
                _cbPixelFormat = Marshal.ReadInt32(MemPtr, Start)
                Start += Marshal.SizeOf(GetType(Int32))
                _offPixelFormat = Marshal.ReadInt32(MemPtr, Start)
                Start += Marshal.SizeOf(GetType(Int32))
                _bOpenGL = (Marshal.ReadInt32(MemPtr, Start) <> 0)
            End If
            If nSize > Start Then
                _szlMicrometersWidth = Marshal.ReadInt32(MemPtr, Start)
                Start += Marshal.SizeOf(GetType(Int32))
                _szlMicrometersHeight = Marshal.ReadInt32(MemPtr, Start)
                Start += Marshal.SizeOf(GetType(Int32))
            End If
            If _nDescription > 0 Then
                _Description = Marshal.PtrToStringAuto(New IntPtr(MemPtr.ToInt32 + Start), _nDescription)
            End If

        End Sub
#End Region

    End Class
#End Region
#Region "DEVMODE"
    Private Class DEVMODE
#Region "Private properties"
        Private dmDeviceName(64) As Char
        Private dmSpecVersion As Short
        Private dmDriverVersion As Short
        Private dmSize As Short
        Private dmDriverExtra As Short
        Private dmFields As Integer
        Private dmOrientation As Short
        Private dmPaperSize As Short
        Private dmPaperLength As Short
        Private dmPaperWidth As Short
        Private dmScale As Short
        Private dmCopies As Short
        Private dmDefaultSource As Short
        Private dmPrintQuality As Short
        Private dmColor As Short
        Private dmDuplex As Short
        Private dmYResolution As Short
        Private dmTTOption As Short
        Private dmCollate As Short
        Private dmFormName(32) As Char
        Private dmUnusedPadding As Short
        Private dmBitsPerPel As Integer
        Private dmPelsWidth As Integer
        Private dmPelsHeight As Integer
        Private dmNup As Integer
        Private dmDisplayFrequency As Integer
        Private dmICMMethod As Integer
        Private dmICMIntent As Integer
        Private dmMediaType As Integer
        Private dmDitherType As Integer
        Private dmReserved1 As Integer
        Private dmReserved2 As Integer
        Private dmPanningWidth As Integer
        Private dmPanningHeight As Integer
        Private DriverExtra() As Byte
#End Region

#Region "Public properties"
#Region "Fields"
        Private ReadOnly Property Fields() As DevModeFields
            Get
                Return New DevModeFields(dmFields)
            End Get
        End Property
#End Region
#Region "DeviceName"
        Public ReadOnly Property DeviceName() As String
            Get
                If dmDeviceName.Length = 64 Then
                    '\\ Remove the balnk chars...
                    Dim c As Char
                    For Each c In dmDeviceName

                    Next
                End If
                Return New String(dmDeviceName)
            End Get
        End Property
#End Region
#Region "FormName"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The name of the form used to print the print job
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Duncan]	17/02/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property FormName() As String
            Get
                If dmFormName Is Nothing Then
                    Return ""
                Else
                    Return New String(dmFormName)
                End If
            End Get
        End Property
#End Region

#Region "Copies"
        Public ReadOnly Property Copies() As Short
            Get
                If dmCopies < 1 Then
                    dmCopies = 1
                End If
                Return dmCopies
            End Get
        End Property
#End Region
#Region "Collate"
        Public ReadOnly Property Collate() As Boolean
            Get
                Return (dmCollate > 0)
            End Get
        End Property
#End Region

#Region "DriverVersion"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The version number of the driver used to print this job
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Duncan]	17/02/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property DriverVersion() As Integer
            Get
                Return dmDriverVersion
            End Get
        End Property
#End Region

#End Region

#Region "Public constructors"
        Public Sub New()


        End Sub
#End Region

#Region "DevModeFields"
        Private Class DevModeFields

#Region "Private properties"
            Private _dmFields As Int32
#End Region

#Region "Private enumerated types"
            <Flags()> _
        Private Enum DeviceModeFieldFlags
                DM_POSITION = &H20
                DM_COLLATE = &H8000
                DM_COLOR = &H800&
                DM_COPIES = &H100&
                DM_DEFAULTSOURCE = &H200&
                DM_DITHERTYPE = &H10000000
                DM_DUPLEX = &H1000&
                DM_FORMNAME = &H10000
                DM_ICMINTENT = &H4000000
                DM_ICMMETHOD = &H2000000
                DM_MEDIATYPE = &H8000000
                DM_ORIENTATION = &H1&
                DM_PAPERLENGTH = &H4&
                DM_PAPERSIZE = &H2&
                DM_PAPERWIDTH = &H8&
                DM_PRINTQUALITY = &H400&
                DM_RESERVED1 = &H800000
                DM_RESERVED2 = &H1000000
                DM_SCALE = &H10&
            End Enum
#End Region

#Region "Public constructor"
            Public Sub New(ByVal Flags As Int32)
                _dmFields = Flags
            End Sub
#End Region

#Region "Public interface"
#Region "Orientation"
            Public ReadOnly Property Orientation() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_ORIENTATION) > 0)
                End Get
            End Property
#End Region

#Region "PaperSize"
            Public ReadOnly Property PaperSize() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_PAPERSIZE) > 0)
                End Get
            End Property
#End Region

#Region "PaperLength"
            Public ReadOnly Property PaperLength() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_PAPERLENGTH) > 0)
                End Get
            End Property
#End Region

#Region "PaperWidth"
            Public ReadOnly Property PaperWidth() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_PAPERWIDTH) > 0)
                End Get
            End Property
#End Region

#Region "Scale"
            Public ReadOnly Property Scale() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_SCALE) > 0)
                End Get
            End Property
#End Region

#Region "Copies"
            Public ReadOnly Property Copies() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_COPIES) > 0)
                End Get
            End Property
#End Region

#Region "DefaultSource"
            Public ReadOnly Property DefaultSource() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_DEFAULTSOURCE) > 0)
                End Get
            End Property
#End Region

#Region "Quality"
            Public ReadOnly Property Quality() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_PRINTQUALITY) > 0)
                End Get
            End Property
#End Region

#Region "Position"
            Public ReadOnly Property Position() As Boolean
                Get
                    Return ((_dmFields And DeviceModeFieldFlags.DM_POSITION) > 0)
                End Get
            End Property
#End Region

#Region "Colour"
            Public ReadOnly Property Colour() As Boolean
                Get
                    Return ((_dmFields Or DeviceModeFieldFlags.DM_COLOR) > 0)
                End Get
            End Property
#End Region

#Region "Duplex"
            Public ReadOnly Property Duplex() As Boolean
                Get
                    Return ((_dmFields Or DeviceModeFieldFlags.DM_DUPLEX) > 0)
                End Get
            End Property
#End Region

#Region "Collation"
            Public ReadOnly Property Collation() As Boolean
                Get
                    Return ((_dmFields Or DeviceModeFieldFlags.DM_COLLATE) > 0)
                End Get
            End Property
#End Region

#Region "Formname"
            Public ReadOnly Property FormName() As Boolean
                Get
                    Return ((_dmFields Or DeviceModeFieldFlags.DM_FORMNAME) > 0)
                End Get
            End Property
#End Region

#Region "MediaType"
            Public ReadOnly Property MediaType() As Boolean
                Get
                    Return ((_dmFields Or DeviceModeFieldFlags.DM_MEDIATYPE) > 0)
                End Get
            End Property
#End Region

#End Region

        End Class
#End Region

    End Class
#End Region
#End Region
End Class
#End Region

#Region "PCL XL Spool File"
''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PCLXLSpoolFile
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Helper class for PCL XL format spool file
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	13/12/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Friend Class PCLXLSpoolFile

End Class
#End Region

#End Region