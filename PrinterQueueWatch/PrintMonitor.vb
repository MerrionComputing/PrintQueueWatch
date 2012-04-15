'\\ --[PrintMonitor]----------------------------------------
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ --------------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApi
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintMonitor
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class wrapper for the windows API calls and constants relating to print monitors
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists teh details of all the print monitors on the current server
''' <code>
'''        Dim server As New PrintServer
'''
'''        Me.ListBox1.Items.Clear()
'''        For ms As Integer = 0 To server.PrintMonitors.Count - 1
'''            Me.ListBox1.Items.Add( server.PrintMonitors(ms).Name )
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintMonitor

#Region "Private members"
    Private _Name As String
    Private _Environment As String
    Private _DLLName As String
#End Region

#Region "Public properties"
#Region "Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print monitor
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property
#End Region

#Region "Environment"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The environment for which the print monitor was created
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' For example, "Windows NT x86", "Windows NT R4000", "Windows NT Alpha_AXP", or "Windows 4.0"
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Environment() As String
        Get
            Return _Environment
        End Get
    End Property
#End Region

#Region "DLLName"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The dynamic link library that implements this print monitor
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DLLName() As String
        Get
            Return _DLLName
        End Get
    End Property
#End Region
#End Region

#Region "Public constructors"
    Friend Sub New(ByVal Name As String, ByVal Environment As String, ByVal DLLName As String)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & Name & "," & Environment & "," & DLLName & ")", Me.GetType.ToString)
        End If
        _Name = Name
        _Environment = Environment
        _DLLName = DLLName
    End Sub
#End Region

#Region "ToString override"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Text description of this print monitor instance
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
        sOut.Append(_Name & ",")
        sOut.Append(_Environment & ",")
        sOut.Append(_DLLName & ",")
        sOut.Append(Me.GetType.ToString)
        Return sOut.ToString
    End Function
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintMonitors
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of print monitors on a server
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists the details of all the print monitors on the current server
''' <code>
'''        Dim server As New PrintServer
'''
'''        Me.ListBox1.Items.Clear()
'''        For ms As Integer = 0 To server.PrintMonitors.Count - 1
'''            Me.ListBox1.Items.Add( server.PrintMonitors(ms).Name )
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
<System.Security.SuppressUnmanagedCodeSecurity()> _
Public Class PrintMonitors
    Inherits Generic.List(Of PrintMonitor)

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a collection of print monitors on the current server
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()

        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pMonitors As Int32
        Dim pcbProvided As Int32

        If Not EnumMonitors(0, 2, 0, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pMonitors = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumMonitors(0, 2, pMonitors, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pMonitors
            While pcReturned > 0
                Dim mi2 As New MONITOR_INFO_2
                Marshal.PtrToStructure(New IntPtr(ptNext), mi2)
                Me.Add(New PrintMonitor(mi2.pName, mi2.pEnvironment, mi2.pDLLName))
                ptNext = (ptNext + Marshal.SizeOf(mi2))
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pMonitors > 0 Then
            Marshal.FreeHGlobal(CType(pMonitors, IntPtr))
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a collection of print monitors on the named server
    ''' </summary>
    ''' <param name="Servername">The name of the server to get the print monitors from</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Servername As String)

        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pMonitors As Int32
        Dim pcbProvided As Int32

        If Not EnumMonitors(Servername, 2, 0, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pMonitors = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumMonitors(Servername, 2, pMonitors, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pMonitors
            While pcReturned > 0
                Dim mi2 As New MONITOR_INFO_2
                Marshal.PtrToStructure(New IntPtr(ptNext), mi2)
                Me.Add(New PrintMonitor(mi2.pName, mi2.pEnvironment, mi2.pDLLName))
                ptNext = ptNext + Marshal.SizeOf(mi2)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pMonitors > 0 Then
            Marshal.FreeHGlobal(CType(pMonitors, IntPtr))
        End If

    End Sub
#End Region

End Class