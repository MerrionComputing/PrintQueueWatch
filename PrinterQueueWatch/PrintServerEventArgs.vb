Imports System.ComponentModel
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

' -- SERVER_EVENTS ----------------------------------------------
' Not implemented in this release
#If SERVER_EVENTS = 1 Then
''' <summary>
''' Generic event argument class from which the print server event arguments all derive
''' </summary>
''' <remarks></remarks>
Public Class PrintServerEventArgs
    Inherits EventArgs

#Region "Private members"
    Private _EventTime As DateTime
#End Region

#Region "Public interface"
    Public ReadOnly Property EventTime() As DateTime
        Get
            Return _EventTime
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New()
        _EventTime = DateTime.Now
    End Sub
#End Region

End Class

''' <summary>
''' Provides data for the <see cref="PrintServer.PrinterDriverAdded">PrinterDriverAdded</see>, 
''' <see cref="PrintServer.PrinterDriverRemoved">PrinterDriverRemoved</see> and 
''' <see cref="PrintServer.PrinterDriverSet">PrinterDriverSet</see> events of the <see cref="PrintServer">Print Server</see> control
''' </summary>
''' <remarks>
''' </remarks>
Public Class PrintServerDriverEventArgs
    Inherits PrintServerEventArgs

#Region "Public enumerated types"
    Public Enum DriverEventTypes
        DriverAddedEvent = 0
        DriverRemovedEvent = 1
        DriverSettingsChangedEvent = 2
    End Enum
#End Region

#Region "Private members"
    Private _ChangeType As DriverEventTypes
#End Region

#Region "Public interface"
    Public ReadOnly Property ChangeType() As DriverEventTypes
        Get
            Return _ChangeType
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New(ByVal Flags As Int32)
        MyBase.New()
        If (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationDriverFlags.PRINTER_CHANGE_ADD_PRINTER_DRIVER) <> 0 Then
            _ChangeType = DriverEventTypes.DriverAddedEvent
        ElseIf (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationDriverFlags.PRINTER_CHANGE_DELETE_PRINTER_DRIVER) <> 0 Then
            _ChangeType = DriverEventTypes.DriverRemovedEvent
        Else
            _ChangeType = DriverEventTypes.DriverSettingsChangedEvent
        End If
    End Sub
#End Region

End Class

''' <summary>
''' Provides data for the <see cref="PrintServer.FormAdded">FormAdded</see>, 
''' <see cref="PrintServer.FormRemoved">FormRemoved</see> and 
''' <see cref="PrintServer.FormSet">FormSet</see> events of the <see cref="PrintServer">Print Server</see> control
''' </summary>
''' <remarks></remarks>
Public Class PrintServerFormEventArgs
    Inherits PrintServerEventArgs

#Region "Public enumerated types"
    Public Enum FormEventTypes
        FormAddedEvent = 0
        FormRemovedEvent = 1
        FormSettingsChangedEvent = 2
    End Enum
#End Region

#Region "Private members"
    Private _ChangeType As FormEventTypes
#End Region

#Region "Public interface"
    Public ReadOnly Property ChangeType() As FormEventTypes
        Get
            Return _ChangeType
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New(ByVal Flags As Int32)
        MyBase.New()
        If (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationFormFlags.PRINTER_CHANGE_ADD_FORM) <> 0 Then
            _ChangeType = FormEventTypes.FormAddedEvent
        ElseIf (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationFormFlags.PRINTER_CHANGE_DELETE_FORM) <> 0 Then
            _ChangeType = FormEventTypes.FormRemovedEvent
        Else
            _ChangeType = FormEventTypes.FormSettingsChangedEvent
        End If
    End Sub
#End Region
End Class

''' <summary>
''' Provides data for the <see cref="PrintServer.PortAdded">PortAdded</see> and 
''' <see cref="PrintServer.PortRemoved">PortRemoved</see> events of the <see cref="PrintServer">Print Server</see> control
''' </summary>
''' <remarks></remarks>
Public Class PrintServerPortEventArgs
    Inherits PrintServerEventArgs

#Region "Public enumerated types"
    Public Enum PortEventTypes
        PortAddedEvent = 0
        PortRemovedEvent = 1
        PortSettingsChangedEvent = 2
    End Enum
#End Region

#Region "Private members"
    Private _ChangeType As PortEventTypes
#End Region

#Region "Public interface"
    Public ReadOnly Property ChangeType() As PortEventTypes
        Get
            Return _ChangeType
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New(ByVal Flags As Int32)
        MyBase.New()
        If (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationPortFlags.PRINTER_CHANGE_ADD_PORT) <> 0 Then
            _ChangeType = PortEventTypes.PortAddedEvent
        ElseIf (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationPortFlags.PRINTER_CHANGE_DELETE_PORT) <> 0 Then
            _ChangeType = PortEventTypes.PortRemovedEvent
        Else
            _ChangeType = PortEventTypes.PortSettingsChangedEvent
        End If
    End Sub
#End Region

End Class

''' <summary>
''' Provides data for the <see cref="PrintServer.PrintProcessorAdded">PrintProcessorAdded</see> and 
''' <see cref="PrintServer.PrintProcessorRemoved">PrintProcessorRemoved</see>  events of the <see cref="PrintServer">Print Server</see> control
''' </summary>
''' <remarks></remarks>
Public Class PrintServerProcessorEventArgs
    Inherits PrintServerEventArgs

#Region "Public enumerated types"
    Public Enum ProcessorEventTypes
        ProcessorAddedEvent = 0
        ProcessorRemovedEvent = 1
    End Enum
#End Region

#Region "Private members"
    Private _ChangeType As ProcessorEventTypes
#End Region

#Region "Public interface"
    Public ReadOnly Property ChangeType() As ProcessorEventTypes
        Get
            Return _ChangeType
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New(ByVal Flags As Int32)
        MyBase.New()
        If (Flags And SpoolerApiConstantEnumerations.PrinterChangeNotificationProcessorFlags.PRINTER_CHANGE_ADD_PRINT_PROCESSOR) <> 0 Then
            _ChangeType = ProcessorEventTypes.ProcessorAddedEvent
        Else
            _ChangeType = ProcessorEventTypes.ProcessorRemovedEvent
        End If
    End Sub
#End Region

End Class


#End If
