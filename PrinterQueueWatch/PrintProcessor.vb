'\\ --[PrintProcessor]--------------------------------
'\\ Class wrapper for the windows API calls and constants
'\\ relating to the printer processors ..
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintProcessor
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class wrapper for the properties relating to the printer processors ..
''' </summary>
''' <remarks>
''' </remarks>
''' <example>
''' <code>
'''        Dim server As New PrintServer
''' 
'''        Dim Processor As PrintProcessor
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.PrintProcessors.Count - 1
'''            Me.ListBox1.Items.Add( server.PrintProcessors(ps).Name )
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintProcessor

#Region "Private properties"
    Private _ServerName As String
    Private _PPInfo1 As PRINTPROCESSOR_INFO_1
#End Region

#Region "Public interface"

#Region "Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print server
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
            Return _PPInfo1.pName
        End Get
    End Property
#End Region

#Region "Server"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The server name on which the print processor is installed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ServerName() As String
        Get
            Return _ServerName
        End Get
    End Property
#End Region

#Region "Data Types"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The collection of data types that this print processor can process
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DataTypes() As DataTypeCollection
        Get
            If _ServerName Is Nothing OrElse _ServerName = "" Then
                Return New DataTypeCollection(_PPInfo1.pName)
            Else
                Return New DataTypeCollection(_ServerName, _PPInfo1.pName)
            End If
        End Get
    End Property
#End Region

#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ServerName">The name of the server that this processor is installed on</param>
    ''' <param name="PrintProcessorname">The name of this print processor</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ServerName As String, ByVal PrintProcessorname As String)
        _ServerName = ServerName
        _PPInfo1 = New PRINTPROCESSOR_INFO_1
        _PPInfo1.pName = PrintProcessorname
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintProcessorCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The collection of print processors installed on a print server
''' </summary>
''' <remarks>
''' </remarks>
''' <example>
''' <code>
'''        Dim server As New PrintServer
''' 
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.PrintProcessors.Count - 1
'''            Me.ListBox1.Items.Add( server.PrintProcessors(ps).Name )
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintProcessorCollection
    Inherits Generic.List(Of PrintProcessor)

#Region "Public interface"
    Default Public Shadows Property Item(ByVal index As Integer) As PrintProcessor
        Get
            Return DirectCast(MyBase.Item(index), PrintProcessor)
        End Get
        Set(ByVal Value As PrintProcessor)
            MyBase.Item(index) = Value
        End Set
    End Property

    Friend Overloads Sub Remove(ByVal obj As PrintProcessor)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a collection of print processors installed on the current print server
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
        Dim pPrintProcessors As Int32
        Dim pcbProvided As Int32 = 0


        If Not EnumPrintProcessors(String.Empty, String.Empty, 1, pPrintProcessors, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pPrintProcessors = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrintProcessors(String.Empty, String.Empty, 1, pPrintProcessors, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPrintProcessors
            While pcReturned > 0
                Dim pi1 As New PRINTPROCESSOR_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), pi1)
                Me.Add(New PrintProcessor("", pi1.pName))
                ptNext = ptNext + Marshal.SizeOf(pi1)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPrintProcessors > 0 Then
            Marshal.FreeHGlobal(CType(pPrintProcessors, IntPtr))
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a collection of print processors installed on the named print server
    ''' </summary>
    ''' <param name="Servername">The name of the print server to return the print processors for</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Servername As String)
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pPrintProcessors As Int32
        Dim pcbProvided As Int32

        If Not EnumPrintProcessors(Servername, String.Empty, 1, 0, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pPrintProcessors = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrintProcessors(Servername, String.Empty, 2, pPrintProcessors, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pPrintProcessors
            While pcReturned > 0
                Dim pi1 As New PRINTPROCESSOR_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), pi1)
                Me.Add(New PrintProcessor("", pi1.pName))
                ptNext = ptNext + Marshal.SizeOf(pi1)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pPrintProcessors > 0 Then
            Marshal.FreeHGlobal(CType(pPrintProcessors, IntPtr))
        End If
    End Sub
#End Region
End Class