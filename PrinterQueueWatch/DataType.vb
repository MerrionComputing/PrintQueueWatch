'\\ --[DataType]--------------------------------
'\\ Class wrapper for the windows API calls and constants
'\\ relating to the printer data types ..
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : DataType
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The data type of a spool file
''' </summary>
''' <remarks>
''' This is the data type that the spool file contains.  It can be EMF or RAW.
''' </remarks>
''' <example>Lists all the data types supported by each of the print processors 
''' installed on the current server
''' <code>
'''        Dim server As New PrintServer
'''        Dim Processor As PrintProcessor
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.PrintProcessors.Count - 1
'''            ListBox1.Items.Add( server.PrintProcessors(ps).Name )
'''            For dt As Integer = 0 to server.PrintProcessors(ps).DataTypes.Count - 1
'''                Me.ListBox1.Items.Add(server.PrintProcessors(ps).DataTypes(dt).Name)
'''            Next
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class DataType

#Region "Private members"
    Private _dti1 As New DATATYPES_INFO_1
#End Region

#Region "Public interface"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the data type of the spool file
    ''' </summary>
    ''' <remarks>
    ''' If this value is RAW then the spool file contains a printer control language
    ''' (such as PostScript, PCL-5, PCL-XL etc.)
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _dti1.pName
        End Get
    End Property
#End Region

#Region "Public constructor"
    Public Sub New(ByVal Name As String)
        _dti1.pName = Name
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : DataTypeCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The collection of <see cref="DataType">data types</see> supported by a print processor 
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class DataTypeCollection
    Inherits Generic.List(Of DataType)

#Region "Public interface"
    ''' <summary>
    ''' Returns a single <see cref="DataType">data type</see> from the collection
    ''' </summary>
    ''' <param name="index">The zero-based position in the collection</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Shadows Property Item(ByVal index As Integer) As DataType
        Get
            Return DirectCast(MyBase.Item(index), DataType)
        End Get
        Friend Set(ByVal Value As DataType)
            MyBase.Item(index) = Value
        End Set
    End Property

    ''' <summary>
    ''' Removes the <see cref="DataType">data type</see> from this collection
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Friend Overloads Sub Remove(ByVal obj As DataType)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new collection containing all the data types
    ''' supported by the named print processor
    ''' </summary>
    ''' <param name="PrintProcessorName">The name of the print processor to retrieve the supported data types for</param>
    ''' <remarks>
    ''' The print processor must be installed on the local machine
    ''' <example >
    ''' Prints the name of all the data types supported by the <see cref="PrintProcessor">print processor</see> named WinPrint
    ''' <code>
    ''' Dim DataTypes As New DataTypeCollection("Winprint")
    ''' For Each dt As DataType In DataTypes
    '''    Trace.WriteLine(dt.Name)
    ''' Next dt
    ''' </code>
    ''' </example>
    ''' </remarks>
    ''' 
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal PrintProcessorName As String)
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pDataTypes As Int32
        Dim pcbProvided As Int32

        If Not EnumPrinterProcessorDataTypes(String.Empty, PrintProcessorName, 1, pDataTypes, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pDataTypes = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrinterProcessorDataTypes(String.Empty, PrintProcessorName, 1, pDataTypes, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pDataTypes
            While pcReturned > 0
                Dim dti1 As New DATATYPES_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), dti1)
                Me.Add(New DataType(dti1.pName))
                ptNext = (ptNext + Marshal.SizeOf(dti1))
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pDataTypes > 0 Then
            Marshal.FreeHGlobal(CType(pDataTypes, IntPtr))
        End If

    End Sub

    ''' <summary>
    ''' Creates a new collection containing all the data types
    ''' supported by the named print processor on the named server
    ''' </summary>
    ''' <param name="Servername">The server on which the print processor resides</param>
    ''' <param name="PrintProcessorName">The print processor name</param>
    ''' <remarks>
    ''' <example >
    ''' Prints the name of all the data types supported by the <see cref="PrintProcessor">print processor</see> named WinPrint
    ''' <code>
    ''' Dim DataTypes As New DataTypeCollection("DUBPDOM1","Winprint")
    ''' For Each dt As DataType In DataTypes
    '''    Trace.WriteLine(dt.Name)
    ''' Next dt
    ''' </code>
    ''' </example>
    ''' </remarks>
    Public Sub New(ByVal Servername As String, ByVal PrintProcessorName As String)
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pDataTypes As Int32
        Dim pcbProvided As Int32

        If EnumPrinterProcessorDataTypes(Servername, PrintProcessorName, 1, pDataTypes, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pDataTypes = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumPrinterProcessorDataTypes(Servername, PrintProcessorName, 1, pDataTypes, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pDataTypes
            While pcReturned > 0
                Dim dti1 As New DATATYPES_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), dti1)
                Me.Add(New DataType(dti1.pName))
                ptNext = (ptNext + Marshal.SizeOf(dti1))
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pDataTypes > 0 Then
            Marshal.FreeHGlobal(CType(pDataTypes, IntPtr))
        End If
    End Sub
#End Region
End Class