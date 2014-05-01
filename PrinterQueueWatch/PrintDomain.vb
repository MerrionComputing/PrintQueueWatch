'\\ --[PrintProvidor]--------------------------------------
'\\ Class wrapper for the top level "Print Domain"
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApi
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrintDomain
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the print domains on all the printer providors visible from this process
''' <code>
'''       Dim Providors As New PrintProvidorCollection
'''
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To Providors.Count - 1
'''            Me.ListBox1.Items.Add( Providors(ps).Name )
'''            For ds As Integer = 0 To  Providors(ps).PrintDomains.Count - 1
'''                Me.ListBox1.Items.Add( Providors(ps).PrintDomains(ds).Name )
'''            Next
'''        Next
''' 
''' </code>
''' </example>
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrintDomain

#Region "Private properties"
    Private _pi1 As New PRINTER_INFO_1
    Private _ProvidorName As String
#End Region

#Region "Public interface"

#Region "Name"
    ''' <summary>
    ''' The name of the print domain
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	17/02/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _pi1.pName
        End Get
    End Property
#End Region

#Region "Description"
    ''' <summary>
    ''' The description of the print domain
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Description() As String
        Get
            Return _pi1.pDescription
        End Get
    End Property
#End Region

#Region "Comment"
    ''' <summary>
    ''' The comment associated with this print domain
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Comment() As String
        Get
            Return _pi1.pComment
        End Get
    End Property
#End Region

#End Region

#Region "Public constructor"
    Friend Sub New(ByVal ProvidorName As String, ByVal Name As String, ByVal Description As String, ByVal Comment As String, ByVal Flags As Integer)

        _ProvidorName = ProvidorName
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
''' Class	 : PrintDomainCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of PrintDomain objects for a given print providor
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the print domains on all the printer providors visible from this process
''' <code>
'''       Dim Providors As New PrintProvidorCollection
'''
'''        Me.ListBox1.Items.Clear()
'''        For ps As Integer = 0 To Providors.Count - 1
'''            Me.ListBox1.Items.Add( Providors(ps).Name )
'''            For ds As Integer = 0 To  Providors(ps).PrintDomains.Count - 1
'''                Me.ListBox1.Items.Add( Providors(ps).PrintDomains(ds).Name )
'''            Next
'''        Next
''' 
''' </code>
''' </example>
<System.Security.SuppressUnmanagedCodeSecurity()> _
Public Class PrintDomainCollection
    Inherits Generic.List(Of PrintDomain)

#Region "Public interface"
    ''' <summary>
    ''' The Item property returns a single <see cref="PrintDomain">print domain</see> from a print domain collection.
    ''' </summary>
    ''' <param name="index">The zero-based item position</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Shadows Property Item(ByVal index As Integer) As PrintDomain
        Get
            Return DirectCast(MyBase.Item(index), PrintDomain)
        End Get
        Friend Set(ByVal Value As PrintDomain)
            MyBase.Item(index) = Value
        End Set
    End Property

    ''' <summary>
    ''' Removes a print domain from this collection
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Friend Overloads Sub Remove(ByVal obj As PrintDomain)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructor"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a list of all the print domains visible from this machine for the named providor
    ''' </summary>
    ''' <param name="ProvidorName">The name of the PrintProvidor to get the list of print domains for</param>
    ''' <remarks>
    ''' </remarks>
    Public Sub New(ByVal ProvidorName As String)
        '\\ Return all the print providors visible from this machine
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer 
        Dim pDomains As IntPtr
        Dim pcbProvided As Int32

        If Not EnumPrinters(PrinterQueueWatch.SpoolerApiConstantEnumerations.EnumPrinterFlags.PRINTER_ENUM_NAME, ProvidorName, 1, pDomains, 0, pcbNeeded, pcReturned) Then
            '\\ Allocate the required buffer to get all the monitors into...
            If pcbNeeded > 0 Then
                pDomains = Marshal.AllocHGlobal(pcbNeeded)
                pcbProvided = pcbNeeded
                If Not EnumPrinters(PrinterQueueWatch.SpoolerApiConstantEnumerations.EnumPrinterFlags.PRINTER_ENUM_NAME, ProvidorName, 1, pDomains, pcbProvided, pcbNeeded, pcReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcReturned > 0 Then
            '\\ Get all the domains for the given print providor
            Dim ptNext As IntPtr = pDomains
            While pcReturned > 0
                Dim pi1 As New PRINTER_INFO_1
                Marshal.PtrToStructure(ptNext, pi1)
                Me.Add(New PrintDomain(ProvidorName, pi1.pName, pi1.pDescription, pi1.pComment, pi1.Flags))
                ptNext = ptNext + Marshal.SizeOf(pi1)
                pcReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pDomains.ToInt64 > 0 Then
            Marshal.FreeHGlobal(pDomains)
        End If

    End Sub
#End Region

End Class