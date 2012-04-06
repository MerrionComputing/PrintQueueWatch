Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterEventArgs
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class representing the event arguments used in the change events 
''' relating to the printer ..
''' </summary>
''' <remarks>
''' This is passed as an argument in the 
''' PrinterInformationChanged event of the PrinterMonitorComponent
''' </remarks>
''' <example>Prints the user name of a printer if an error occurs
''' <code>
'''   Private WithEvents pmon As New PrinterMonitorComponent
'''
'''   pmon.AddPrinter("Microsoft Office Document Image Writer")
'''   pmon.AddPrinter("HP Laserjet 5")
''' 
'''     Private Sub pmon_PrinterInformationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pmon.PrinterInformationChanged
'''        With CType(e, PrinterEventArgs)
'''            If .PrinterInformation.IsInError Then
'''                Trace.WriteLine(.PrinterInformation.PrinterName)
'''            End If
'''        End With
'''    End Sub
''' </code>
''' </example>
Public Class PrinterEventArgs
    Inherits EventArgs

#Region "Private member variables"
    Private mPrinterInfo As PrinterInformation
    Private mEventTime As DateTime
    Private mPrinterChangeFlags As PrinterInfoChangeFlagDecoder
#End Region

#Region "Public Properties"
#Region "PrinterInformation"
    ''' <summary>
    ''' The PrinterInformation class 
    ''' that represents the settings of the printer for which the event was triggered
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This holds the printer information as it was when the event occured
    ''' </remarks>
    Public ReadOnly Property PrinterInformation() As PrinterInformation
        Get
            Return mPrinterInfo
        End Get
    End Property
#End Region

#Region "EventTime"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The date and time at which this printer information changed event was triggered
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property EventTime() As DateTime
        Get
            Return mEventTime
        End Get
    End Property
#End Region

#Region "PrinterChangeFlags"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The PrinterInfoChangeFlagDecoder class that holds details of what printer settings changed to trigger this 
    ''' printer change event
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property PrinterChangeFlags() As PrinterInfoChangeFlagDecoder
        Get
            Return mPrinterChangeFlags
        End Get
    End Property
#End Region

    Public Overloads Function Equals(ByVal PrinterEventArgs As PrinterEventArgs) As Boolean
        If Not PrinterEventArgs Is Nothing Then
            If PrinterEventArgs.PrinterChangeFlags.ChangeFlags = mPrinterChangeFlags.ChangeFlags Then
                Return True
            Else
                Return False
            End If
        End If
    End Function
#End Region

#Region "Public Constructors"
    Friend Sub New(ByVal dwFlags As Integer, ByVal PrinterInfo As PrinterInformation, ByVal time As DateTime)
        mPrinterInfo = PrinterInfo
        mPrinterChangeFlags = New PrinterInfoChangeFlagDecoder(dwFlags)
        mEventTime = time
    End Sub

    Friend Sub New(ByVal dwFlags As Integer, ByVal PrinterInfo As PrinterInformation)
        Me.New(dwFlags, PrinterInfo, System.DateTime.Now)
    End Sub

#End Region

End Class

'\\ --[PrinterInfoChangeFlagDecoder]-------------------------------
'\\ Splits the printer change flags up into components to allow 
'\\ developers to respond differently depending on the nature of
'\\ the printer change event
'\\ ---------------------------------------------------------------
''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterInfoChangeFlagDecoder
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' This class holds details of what printer settings changed to trigger this 
''' printer change event
''' </summary>
''' <remarks>
''' A single printer change event may have more than one cause 
''' </remarks>
Public Class PrinterInfoChangeFlagDecoder

#Region "Private member variables"
    Private _mdwFlags As Integer
#End Region

#Region "Public interface"

    ''' <summary>
    ''' True if the number of jobs on the print queue changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the JobCount member of the PrinterInformation passed with the event
    ''' </remarks>
    Public ReadOnly Property JobCountChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_CJOBS)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' True if the printer status has changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the status fields of the 
    ''' PrinterInformation passed with the event
    ''' </remarks>
    Public ReadOnly Property StatusChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_STATUS)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The printer attributes have changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the attributes fields of the 
    ''' PrinterInformation passed with the event
    ''' </remarks>
    <Description("The printer attributes have changed")> _
    Public ReadOnly Property AttributesChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_ATTRIBUTES)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The comment text associated with this printer has been changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the Comment
    ''' member of the PrinterInformation passed with the event
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("The comment text has changed")> _
    Public ReadOnly Property CommentChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_COMMENT)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default device mode settings of the printer have changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the default device mode related fields of the 
    ''' PrinterInformation passed with the event
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("The default DEVMODE has changed")> _
    Public ReadOnly Property DeviceModeChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_DEVMODE)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The printer location text has been changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the Location
    ''' member of the PrinterInformation passed with the event
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("The printer location text has changed")> _
    Public ReadOnly Property LocationChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_LOCATION)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The discretionary access control settings for the printer has changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("The discretionary access control for the printer has changed")> _
    Public ReadOnly Property SecurityChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_SECURITY_DESCRIPTOR)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The file used as a job seperator was changed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' The new value will be held in the SeperatorFilename
    ''' member of the PrinterInformation passed with the event.
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("The file used as a job seperator was changed")> _
    Public ReadOnly Property SeperatorFileChanged() As Boolean
        Get
            Return ((_mdwFlags And CLng(2 ^ Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_SEPFILE)) <> 0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Unknown change - insufficient change information was provided by the printer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This will be true if the printer driver does not issue details of why a printer change 
    ''' event occurs
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Description("Unknown change - insufficient change information was provided by the printer")> _
    Public ReadOnly Property UnknownChange() As Boolean
        Get
            Return (_mdwFlags = 0)
        End Get
    End Property

    Friend ReadOnly Property ChangeFlags() As Int32
        Get
            Return _mdwFlags
        End Get
    End Property

#End Region

#Region "Public constructors"
    Friend Sub New(ByVal dwFlags As Integer)
        _mdwFlags = dwFlags
    End Sub
#End Region

End Class