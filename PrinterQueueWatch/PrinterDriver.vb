'\\ --[PrinterDriver]--------------------------------
'\\ Class wrapper for the windows API calls and constants
'\\ relating to the printer drivers ..
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ ----------------------------------------------------
Imports System.ComponentModel
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations
Imports PrinterQueueWatch.PrinterMonitoringExceptions
Imports System.Runtime.InteropServices
Imports System.IO

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterDriver
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Specifies the settings for a printer driver
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the printer drivers on the current server
''' <code>
'''        Dim server As New PrintServer
'''
'''        Dim Driver As PrinterDriver
'''        Me.ListBox1.Items.Clear()
'''        For Each Driver In server.PrinterDrivers
'''            Me.ListBox1.Items.Add(Driver.Name)
'''        Next
''' </code>
''' </example>
''' <seealso cref="PrinterDriver" />
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrinterDriver

#Region "Private members"
    Private _Driver_Info_3 As DRIVER_INFO_3
#End Region

#Region "Shared methods"

#Region "AddPrinterDriver"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Installs a printer driver on the named server
    ''' </summary>
    ''' <param name="Servername">The server to install the driver on</param>
    ''' <param name="OperatingSystemVersion">The operating system version that the driver targets</param>
    ''' <param name="DriverName">The name of the driver</param>
    ''' <param name="Environment">The environment for which the driver was written (for example, "Windows NT x86", "Windows NT R4000", "Windows NT Alpha_AXP", or "Windows 4.0")</param>
    ''' <param name="DriverFile">The driver program file</param>
    ''' <param name="DriverDataFile">The file which contains the driver data</param>
    ''' <param name="DriverConfigFile">file name or a full path and file name for the device-driver's configuration .dll</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Before an application calls the AddPrinterDriver function, all files 
    ''' required by the driver must be copied to the system's printer-driver directory.
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function AddPrinterDriver(ByVal Servername As String, ByVal OperatingSystemVersion As PrinterDriverOperatingSystemVersion, ByVal DriverName As String, ByVal Environment As String, ByVal DriverFile As FileInfo, ByVal DriverDataFile As FileInfo, ByVal DriverConfigFile As FileInfo) As PrinterDriver

        '\\ Validate input data
        If Not DriverFile.Exists Then
            Throw New FileNotFoundException("File not found: " & DriverFile.FullName)
            Exit Function
        End If

        '\\ Make a DRIVER_INFO_2 with the parameters passed in
        Dim di2 As New DRIVER_INFO_2
        di2.cVersion = OperatingSystemVersion
        di2.pName = DriverName
        di2.pDriverPath = DriverFile.FullName
        If DriverConfigFile.Exists Then
            di2.pConfigFile = DriverConfigFile.FullName
        End If
        If DriverDataFile.Exists Then
            di2.pDatafile = DriverDataFile.FullName
        End If
        di2.pEnvironment = Environment

        '\\ Call the AddPrinterDriver API call
        If SpoolerApi.AddPrinterDriver(Servername, 2, di2) Then
            Return New PrinterDriver(di2)
        Else
            Throw New Win32Exception
        End If

    End Function
#End Region

#Region "GetPrinterDriverDirectory"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the directory on the current machine in which the printer drivers and
    ''' their support files are kept
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function GetPrinterDriverDirectory() As DirectoryInfo
        Return GetPrinterDriverDirectory(String.Empty, String.Empty)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the directory on the named machine in which the printer drivers and
    ''' their support files are kept
    ''' </summary>
    ''' <param name="Servername">The name of the machine to get the directory for</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function GetPrinterDriverDirectory(ByVal Servername As String) As DirectoryInfo
        Return GetPrinterDriverDirectory(Servername, String.Empty)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the directory on the named machine in which the printer drivers and
    ''' their support files are kept for the named environment
    ''' </summary>
    ''' <param name="Servername">The name of the machine to get the directory for</param>
    ''' <param name="Environment">The environment for which the driver was written (for example, "Windows NT x86", "Windows NT R4000", "Windows NT Alpha_AXP", or "Windows 4.0")</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function GetPrinterDriverDirectory(ByVal Servername As String, ByVal Environment As String) As DirectoryInfo

        Dim DriverDirectory As String = ""
        Dim BytesNeeded As Integer

        If SpoolerApi.GetPrinterDriverDirectory(Servername, Environment, 1, DriverDirectory, DriverDirectory.Length, BytesNeeded) Then
            '\\ Empty string should not return any values...
            Return New DirectoryInfo(DriverDirectory)
        Else
            DriverDirectory = New String(Char.Parse(" "), BytesNeeded)
            If SpoolerApi.GetPrinterDriverDirectory(Servername, Environment, 1, DriverDirectory, DriverDirectory.Length, BytesNeeded) Then
                '\\ Empty string should not return any values...
                Return New DirectoryInfo(DriverDirectory)
            Else
                Throw New Win32Exception
            End If
        End If

    End Function
#End Region

#End Region

#Region "Public interface"

#Region "OperatingSystemVersion"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The operating system for which this driver was written
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    Public ReadOnly Property OperatingSystemVersion() As SpoolerApiConstantEnumerations.PrinterDriverOperatingSystemVersion
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.cVersion
            End If
        End Get
    End Property
#End Region

#Region "Name"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The unique name by which this printer driver is known
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pName
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "Environment"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The environment for which the driver was written 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' For example, "Windows NT x86", "Windows NT R4000", "Windows NT Alpha_AXP", or "Windows 4.0"
    ''' </remarks>
    Public ReadOnly Property Environment() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pEnvironment
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "DriverPath"
    ''' <summary>
    ''' The file name or full path and file name for the file that contains the device driver 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>"
    ''' For example, "c:\drivers\pscript.dll"
    ''' This value will be relative to the server
    ''' </remarks>
    Public ReadOnly Property DriverPath() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pDriverPath
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "DataFile"
    ''' <summary>
    ''' The file name or a full path and file name for the file that contains driver data
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' For example, "c:\drivers\Qms810.ppd" 
    ''' </remarks>
    Public ReadOnly Property DataFile() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pDatafile
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "ConfigurationFile"
    ''' <summary>
    ''' The file name or a full path and file name for the device-driver's configuration .dll
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' These configuration files provide the user interface for the extra 
    ''' </remarks>
    Public ReadOnly Property ConfigurationFile() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pConfigFile
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "HelpFile"
    ''' <summary>
    ''' The file name or a full path and file name for the device driver's help file.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This member may be blank if the driver has no help file
    ''' </remarks>
    Public ReadOnly Property HelpFile() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pHelpFile
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "MonitorName"
    ''' <summary>
    ''' The name of a language monitor attached to this driver (for example, "PJL monitor")
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This member can be empty and is specified only for printers capable of bidirectional communication
    ''' </remarks>
    Public ReadOnly Property MonitorName() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pMonitorName
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#Region "DefaultDataType"
    ''' <summary>
    ''' The default data type used by this printer driver in writing spool files for print jobs
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' This can be EMF or RAW.  The latter indicates that a printer control language 
    ''' (such as PCL or PostScript) is used.
    ''' </remarks>
    Public ReadOnly Property DefaultDataType() As String
        Get
            If Not _Driver_Info_3 Is Nothing Then
                Return _Driver_Info_3.pDefaultDataType
            Else
                Return ""
            End If
        End Get
    End Property
#End Region

#End Region

#Region "Public constructors"
    Friend Sub New(ByVal hPrinter As Int32)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & hPrinter.ToString & ")", Me.GetType.ToString)
        End If
        _Driver_Info_3 = New DRIVER_INFO_3(hPrinter)
    End Sub

    Friend Sub New(ByVal dInfo3 As DRIVER_INFO_3)
        _Driver_Info_3 = dInfo3
    End Sub

    Friend Sub New(ByVal dInfo2 As DRIVER_INFO_2)
        With _Driver_Info_3
            .cVersion = dInfo2.cVersion
            .pConfigFile = dInfo2.pConfigFile
            .pDatafile = dInfo2.pDatafile
            .pDriverPath = dInfo2.pDriverPath
            .pEnvironment = dInfo2.pEnvironment
            .pName = dInfo2.pName
        End With
    End Sub
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : PrinterQueueWatch
''' Class	 : PrinterDriverCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The collection of printer drivers on a server
''' </summary>
''' <remarks>
''' </remarks>
''' <example>Lists all the printer drivers on the current server
''' <code>
'''        Dim server As New PrintServer
'''
'''        Dim Driver As PrinterDriver
'''        Me.ListBox1.Items.Clear()
'''        For Each Driver In server.PrinterDrivers
'''            Me.ListBox1.Items.Add(Driver.Name)
'''        Next
''' </code>
''' </example>
<System.Runtime.InteropServices.ComVisible(False)> _
Public Class PrinterDriverCollection
    Inherits Generic.List(Of PrinterDriver)

#Region "Public interface"
    ''' <summary>
    ''' The Item property returns a single <see cref="PrinterDriver">printer driver</see> from this collection.
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Shadows Property Item(ByVal index As Integer) As PrinterDriver
        Get
            Return DirectCast(MyBase.Item(index), PrinterDriver)
        End Get
        Friend Set(ByVal Value As PrinterDriver)
            MyBase.Item(index) = Value
        End Set
    End Property

    Friend Overloads Sub Remove(ByVal obj As PrinterDriver)
        MyBase.Remove(obj)
    End Sub
#End Region

#Region "Public constructors"

    Public Sub New()
        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer (in bytes)
        Dim pDriverInfo As Int32
        Dim nItem As Integer
        Dim pNextDriverInfo As Int32

        If Not EnumPrinterDrivers(String.Empty, String.Empty, 3, pDriverInfo, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pDriverInfo = CInt(Marshal.AllocHGlobal(pcbNeeded))
                If EnumPrinterDrivers(String.Empty, String.Empty, 3, pDriverInfo, pcbNeeded, pcbNeeded, pcReturned) Then
                    If pcReturned > 0 Then
                        pNextDriverInfo = pDriverInfo
                        For nItem = 1 To pcReturned
                            Dim pdInfo3 As New DRIVER_INFO_3
                            '\\ Read the DRIVER_INFO_3 from the buffer
                            Marshal.PtrToStructure(New IntPtr(pNextDriverInfo), pdInfo3)
                            '\\ Add this to the return list
                            Me.Add(New PrinterDriver(pdInfo3))
                            '\\ Move the buffer pointer on to the next DRIVER_INFO_3 structure
                            pNextDriverInfo = pNextDriverInfo + Marshal.SizeOf(pdInfo3)
                        Next
                    End If
                End If
                Marshal.FreeHGlobal(CType(pDriverInfo, IntPtr))
            End If
        End If

    End Sub

    Public Sub New(ByVal Servername As String)

        Dim pcbNeeded As Int32 '\\ Holds the requires size of the output buffer (in bytes)
        Dim pcReturned As Int32 '\\ Holds the returned size of the output buffer (in bytes)
        Dim pDriverInfo As Int32
        Dim nItem As Integer
        Dim pNextDriverInfo As Int32

        If Not EnumPrinterDrivers(Servername, String.Empty, 3, pDriverInfo, 0, pcbNeeded, pcReturned) Then
            If pcbNeeded > 0 Then
                pDriverInfo = CInt(Marshal.AllocHGlobal(pcbNeeded))
                If EnumPrinterDrivers(Servername, String.Empty, 3, pDriverInfo, pcbNeeded, pcbNeeded, pcReturned) Then
                    If pcReturned > 0 Then
                        pNextDriverInfo = pDriverInfo
                        For nItem = 1 To pcReturned
                            Dim pdInfo3 As New DRIVER_INFO_3
                            '\\ Read the DRIVER_INFO_3 from the buffer
                            Marshal.PtrToStructure(New IntPtr(pNextDriverInfo), pdInfo3)
                            '\\ Add this to the return list
                            Me.Add(New PrinterDriver(pdInfo3))
                            '\\ Move the buffer pointer on to the next DRIVER_INFO_3 structure
                            pNextDriverInfo = pNextDriverInfo + Marshal.SizeOf(pdInfo3)
                        Next
                    End If
                End If
                Marshal.FreeHGlobal(CType(pDriverInfo, IntPtr))
            End If
        End If

    End Sub
#End Region

End Class
