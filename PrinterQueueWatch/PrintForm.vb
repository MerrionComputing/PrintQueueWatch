'\\ --[PrinterForm]----------------------------------------
'\\ Class wrapper for the "form" related API 
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApi
Imports System.Runtime.InteropServices
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterForm
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents a form that has been installed on a printer
''' </summary>
''' <remarks>
''' </remarks>
''' <example>List the print forms on the named printer
''' <code>
'''        Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
'''
'''        For pf As Integer = 0 to pi.PrinterForms.Count -1
'''            pi.PrinterForms(pf).Name
'''        Next
''' </code>
''' </example>
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrinterForm

#Region "Private properties"
    Private _hPrinter As Int32
    Private _fi1 As New FORM_INFO_1
#End Region

#Region "Public interface"

#Region "Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the printer form 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _fi1.Name
        End Get
    End Property
#End Region

#Region "Width"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The width of the form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is measured in millimeters
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Width() As Integer
        Get
            Return _fi1.Width
        End Get
        Set(ByVal value As Integer)
            If value <> _fi1.Width Then
                _fi1.Width = value
                Call SaveForm()
            End If
        End Set
    End Property
#End Region

#Region "Height"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The height of the printer form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This value is measured in millimeters
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Height() As Integer
        Get
            Return _fi1.Height
        End Get
        Set(ByVal value As Integer)
            If value <> _fi1.Height Then
                _fi1.Height = value
                Call SaveForm()
            End If
        End Set
    End Property
#End Region

#Region "ImageableArea"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the width and height, in thousandths of millimeters, of the form. 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This may be smaller than the height and width if there is a non-printable margin
    ''' on the form
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	19/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ImageableArea() As Rectangle
        Get
            Return New Rectangle(CInt(_fi1.Left), CInt(_fi1.Top), CInt(_fi1.Right - _fi1.Left), CInt(_fi1.Bottom - _fi1.Top))
        End Get
        Set(ByVal Value As Rectangle)
            With Value
                _fi1.Left = .Left
                _fi1.Top = .Top
                _fi1.Right = .Width + .Left
                _fi1.Bottom = .Top + .Height
            End With
            Call SaveForm()
        End Set
    End Property
#End Region

#Region "Flags related"
    Public ReadOnly Property IsBuiltInForm() As Boolean
        Get
            Return (_fi1.Flags = SpoolerApiConstantEnumerations.FormTypeFlags.FORM_BUILTIN)
        End Get
    End Property

    Public ReadOnly Property IsUserForm() As Boolean
        Get
            Return (_fi1.Flags = SpoolerApiConstantEnumerations.FormTypeFlags.FORM_USER)
        End Get
    End Property

    Public ReadOnly Property IsPrinterForm() As Boolean
        Get
            Return (_fi1.Flags = SpoolerApiConstantEnumerations.FormTypeFlags.FORM_PRINTER)
        End Get
    End Property
#End Region

#End Region

#Region "Public constructor"
    Friend Sub New(ByVal hPrinter As Int32, _
                   ByVal Flags As Int32, _
                   ByVal Name As String, _
                   ByVal Width As Int32, _
                   ByVal Height As Int32, _
                   ByVal Left As Int32, _
                   ByVal Top As Int32, _
                   ByVal Right As Int32, _
                   ByVal Bottom As Int32)

        _hPrinter = hPrinter

        With _fi1
            .Name = Name
            .Flags = Flags
            .Bottom = Bottom
            .Top = Top
            .Left = Left
            .Right = Right
            .Height = Height
            .Width = Width
        End With

    End Sub

    Public Sub New(ByVal Name As String)

    End Sub
#End Region

#Region "Private methods"
    Private Sub SaveForm()
        If Not SetForm(_hPrinter, _fi1.Name, 1, _fi1) Then
            If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                Trace.WriteLine("SetForm call failed", Me.GetType.ToString)
            End If
            Throw New Win32Exception
        End If
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterFormCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of PrintForm objects supported by a printer
''' </summary>
''' <remarks>
''' </remarks>
''' <example>List the print forms on the named printer
''' <code>
'''        Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)
'''
'''        For pf As Integer = 0 To In pi.PrinterForms.Count - 1
'''            Me.ListBox1.Items.Add( pi.PrinterForms(pf).Name )
'''        Next
''' </code>
''' </example>
''' <seealso cref="PrinterQueueWatch.PrinterInformation.PrinterForms" />
''' <history>
''' 	[Duncan]	19/11/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<System.Runtime.InteropServices.ComVisible(False)> _
<System.Security.SuppressUnmanagedCodeSecurity()> _
Public Class PrinterFormCollection
    Inherits Generic.List(Of PrinterForm)

#Region "Public interface"
    Default Public Shadows Property Item(ByVal index As Integer) As PrinterForm
        Get
            Return DirectCast(MyBase.Item(index), PrinterForm)
        End Get
        Set(ByVal Value As PrinterForm)
            MyBase.Item(index) = Value
        End Set
    End Property

    Public Overloads Sub Remove(ByVal obj As PrinterForm)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"

    Friend Sub New(ByVal hPrinter As Int32)

        Dim pForm As Int32
        Dim pcbNeeded As Int32
        Dim pcFormsReturned As Int32
        Dim pcbProvided As Int32

        If Not EnumForms(hPrinter, 1, pForm, 0, pcbNeeded, pcFormsReturned) Then
            If pcbNeeded > 0 Then
                pForm = CInt(Marshal.AllocHGlobal(pcbNeeded))
                pcbProvided = pcbNeeded
                If Not EnumForms(hPrinter, 1, pForm, pcbProvided, pcbNeeded, pcFormsReturned) Then
                    Throw New Win32Exception
                End If
            End If
        End If

        If pcFormsReturned > 0 Then
            '\\ Get all the monitors for the given server
            Dim ptNext As Int32 = pForm
            While pcFormsReturned > 0
                Dim fi1 As New FORM_INFO_1
                Marshal.PtrToStructure(New IntPtr(ptNext), fi1)
                Me.Add(New PrinterForm(hPrinter, fi1.Flags, fi1.Name, fi1.Width, fi1.Height, fi1.Left, fi1.Top, fi1.Right, fi1.Bottom))
                ptNext = ptNext + Marshal.SizeOf(fi1)
                pcFormsReturned -= 1
            End While
        End If

        '\\ Free the allocated buffer memory
        If pForm > 0 Then
            Marshal.FreeHGlobal(CType(pForm, IntPtr))
        End If

    End Sub
#End Region

End Class