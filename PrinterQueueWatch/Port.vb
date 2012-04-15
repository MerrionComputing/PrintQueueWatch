'\\ --[PrintServer]----------------------------------------
'\\ Class wrapper for the "port" related API 
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApi
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : Port
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents information about the port to which a printer is attached
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the ports on the current print server
''' <code>
'''        Dim server As New PrintServer
'''
'''        ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.Ports.Count - 1
'''            Me.ListBox1.Items.Add(server.Ports(ps).Name)
'''        Next
''' </code>
''' </example>
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class Port

#Region "Private properties"
    Private _Servername As String
    Private _pi2 As PORT_INFO_2
#End Region

#Region "Public interface"

#Region "Server Name"
    ''' <summary>
    ''' The name of the server on which this port resides
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Servername() As String
        Get
            Return _Servername
        End Get
    End Property
#End Region

#Region "Name"
    ''' <summary>
    ''' The name supported printer port (for example, "LPT1:").
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Name() As String
        Get
            Return _pi2.pPortName
        End Get
    End Property
#End Region

#Region "Monitor Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Identifies an installed monitor (for example, "PJL monitor"). 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This may be empty if no port monitor associated with this port
    ''' </remarks>
    Public ReadOnly Property MonitorName() As String
        Get
            Return _pi2.pMonitorName
        End Get
    End Property
#End Region

#Region "Description"
    ''' <summary>
    ''' More detailed description of this printer port
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Description() As String
        Get
            Return _pi2.pDescription
        End Get
    End Property
#End Region

#Region "PortTypes related"

#Region "Read"
    ''' <summary>
    ''' The port supports read functionality
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Read() As Boolean
        Get
            Return ((_pi2.PortType And SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_READ) = SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_READ)
        End Get
    End Property
#End Region

#Region "Write"
    ''' <summary>
    ''' The port supports write operation
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Write() As Boolean
        Get
            Return ((_pi2.PortType And SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_WRITE) = SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_WRITE)
        End Get
    End Property
#End Region

#Region "Redirected"
    ''' <summary>
    ''' True if the port is redirected
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property Redirected() As Boolean
        Get
            Return ((_pi2.PortType And SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_REDIRECTED) = SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_REDIRECTED)
        End Get
    End Property
#End Region

#Region "NetAttached"
    ''' <summary>
    ''' True if the port is network attached
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property NetAttached() As Boolean
        Get
            Return ((_pi2.PortType And SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_NET_ATTACHED) = SpoolerApiConstantEnumerations.PortTypes.PORT_TYPE_NET_ATTACHED)
        End Get
    End Property
#End Region

#End Region


#End Region

#Region "public constructor"
    Friend Sub New(ByVal Servername As String, ByVal pi2In As PORT_INFO_2)
        _Servername = Servername
        _pi2 = pi2In
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PortCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of Port objects
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the ports on the current printer
''' <code>
'''        Dim server As New PrintServer
'''
'''        ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.Ports.Count - 1
'''            Me.ListBox1.Items.Add(server.Ports(ps).Name)
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PortCollection
    Inherits Generic.List(Of Port)

#Region "Public interface"
    Default Public Shadows Property Item(ByVal index As Integer) As Port
        Get
            Return DirectCast(MyBase.Item(index), Port)
        End Get
        Set(ByVal Value As Port)
            MyBase.Item(index) = Value
        End Set
    End Property

    Public Overloads Sub Remove(ByVal obj As Port)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the ports on the current machine
    ''' </summary>
    ''' <remarks>
    ''' To get the ports of another machine use the overloaded constructor and
    ''' pass the server name as a reference
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pPorts As Int32
        Dim pcbProvided As Int32

        If Not EnumPorts(0, 2, 0, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pPorts = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPorts(0, 2, pPorts, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPorts
            While pcReturned > 0
                Dim mi2 As New PORT_INFO_2
                Marshal.PtrToStructure(New IntPtr(ptNext), mi2)
                Me.Add(New Port("", mi2))
                ptNext = ptNext + Marshal.SizeOf(mi2)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPorts > 0 Then
            Marshal.FreeHGlobal(CType(pPorts, IntPtr))
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enumerates the ports on the named server
    ''' </summary>
    ''' <param name="Servername">The server to list the ports on</param>
    ''' <remarks>
    ''' </remarks>
    ''' <exception cref="System.ComponentModel.Win32Exception">
    ''' The server does not exist of the user does not have access to it
    ''' </exception>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Servername As String)

        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pPorts As Int32
        Dim pcbProvided As Int32

        If Not EnumPorts(Servername, 2, 0, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pPorts = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPorts(Servername, 2, pPorts, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPorts
            While pcReturned > 0
                Dim mi2 As New PORT_INFO_2
                Marshal.PtrToStructure(New IntPtr(ptNext), mi2)
                Me.Add(New Port(Servername, mi2))
                ptNext = ptNext + Marshal.SizeOf(mi2)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPorts > 0 Then
            Marshal.FreeHGlobal(CType(pPorts, IntPtr))
        End If

    End Sub
#End Region

End Class
