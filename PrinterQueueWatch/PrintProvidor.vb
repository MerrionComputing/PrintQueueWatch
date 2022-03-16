'\\ --[PrintProvidor]--------------------------------------
'\\ Class wrapper for the top level "Print Providor"
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' Project	 : PrinterQueueWatch
''' Class	 : PrintProvidor
''' 
''' <summary>
''' Class representing the properties of a print provider on this domain
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all  the printer providors visible from this process
''' <code>
'''
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.PrintProvidors.Count - 1
'''            Me.ListBox1.Items.Add( server.PrintProvidors(ps).Name )
'''        Next
''' 
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintProvidor

#Region "Private properties"
    Private _pi1 As New PRINTER_INFO_1
#End Region

#Region "Public interface"

#Region "Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the print provider
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' e.g. "Windows NT Local provider" (for local printers),
    ''' "Windows NT Remote provider" (for network printers) etc.
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _pi1.pName
        End Get
    End Property
#End Region

#Region "Description"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The description of the print provider
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' e.g. "Windows NT Local Printers" etc.
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Description() As String
        Get
            Return _pi1.pDescription
        End Get
    End Property
#End Region

#Region "Comment"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The comment text (if any) associated with this provider
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Comment() As String
        Get
            Return _pi1.pComment
        End Get
    End Property
#End Region

#Region "PrintDomains"
    ''' <summary>
    ''' The print domains serviced by this print provider
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    ''' </history>
    Public ReadOnly Property PrintDomains() As PrintDomainCollection
        Get
            Return New PrintDomainCollection(Me.Name)
        End Get
    End Property
#End Region

#End Region

#Region "Public constructor"
    Friend Sub New(ByVal Name As String, ByVal Description As String, ByVal Comment As String, ByVal Flags As Integer)
        With _pi1
            .pName = Name
            .pDescription = Description
            .pComment = Comment
            .Flags = Flags
        End With
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintProvidorCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The collection of print providors accessible from this machine
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the printer providors visible from this process
''' <code>
'''       Dim server As New PrintServer
'''
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To server.Providors.Count - 1 
'''            Me.ListBox1.Items.Add( server.Providors(ps).Name )
'''        Next
''' 
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	21/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Security.SuppressUnmanagedCodeSecurity()> _
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintProvidorCollection
    Inherits Generic.List(Of PrintProvidor)

#Region "Public interface"
    Default Public Shadows Property Item(ByVal index As Integer) As PrintProvidor
        Get
            Return DirectCast(MyBase.Item(index), PrintProvidor)
        End Get
        Set(ByVal Value As PrintProvidor)
            MyBase.Item(index) = Value
        End Set
    End Property

    Public Overloads Sub Remove(ByVal obj As PrintProvidor)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructor"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a collection of all the print providors accessible from this machine
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	21/11/2005	Created
    '''     [Duncan]    01/05/2014  Use IntPtr for 32/64 bit compatibility
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        '\\ Return all the print providors visible from this machine
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pProvidors As IntPtr
        Dim pcbProvided As Int32

        ' If Not EnumPrinters(0,string.Empty  1, 0, 0, pcbNeeded, pcReturned) Then
        If Not EnumPrinters(PrinterQueueWatch.SpoolerApiConstantEnumerations.EnumPrinterFlags.PRINTER_ENUM_NAME, 0, 1, pProvidors, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pProvidors = Marshal.AllocHGlobal(pcbNeeded)
                pcbProvided = pcbNeeded
                If Not EnumPrinters(PrinterQueueWatch.SpoolerApiConstantEnumerations.EnumPrinterFlags.PRINTER_ENUM_NAME, 0, 1, pProvidors, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As IntPtr = pProvidors
            While pcReturned > 0
                Dim pi1 As New PRINTER_INFO_1
                Marshal.PtrToStructure(ptNext, pi1)
                Me.Add(New PrintProvidor(pi1.pName, pi1.pDescription, pi1.pComment, pi1.Flags))
                ptNext = ptNext + Marshal.SizeOf(pi1)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pProvidors.ToInt64 > 0 Then
            Marshal.FreeHGlobal(pProvidors)
        End If

    End Sub
#End Region

End Class
