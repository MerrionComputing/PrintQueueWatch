'\\ --[SpoolerStructs]--------------------------------------------------------------
'\\ The structures used by the Win32 API calls concerned with the spooler as 
'\\ defined in winspool.drv
'\\ (c) 2003 Merrion Computing Ltd
'\\ http://www.merrioncomputing.com
'\\ ---------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports System.Drawing.Printing
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations


Namespace SpoolerStructs

#Region "SYSTEMTIME STRUCTURE"
    <StructLayout(LayoutKind.Sequential)> Friend Class SYSTEMTIME
        Public wYear As Int16
        Public wMonth As Int16
        Public wDayOfWeek As Int16
        Public wDay As Int16
        Public wHour As Int16
        Public wMinute As Int16
        Public wSecond As Int16
        Public wMilliseconds As Int16

        Public Function ToDateTime() As DateTime

            Return New DateTime(wYear, wMonth, wDay, wHour, wMinute, wSecond, wMilliseconds)

        End Function

    End Class
#End Region

#Region "JOB_INFO_1 STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class JOB_INFO_1
        Public JobId As Int32
        <MarshalAs(UnmanagedType.LPWStr)> Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pMachineName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pUserName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDocument As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDatatype As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pStatus As String
        <MarshalAs(UnmanagedType.U4)> Public Status As Int32
        Public Priority As Int32
        Public Position As Int32
        Public TotalPage As Int32
        Public PagesPrinted As Int32
        <MarshalAs(UnmanagedType.Struct)> Public Submitted As SYSTEMTIME

        Public Sub New()
            '\\ If this structure is not "Live" then a printer handle and job id are not used
        End Sub

        Public Sub New(ByVal hPrinter As Int32, _
                     ByVal dwJobId As Int32)

            Dim BytesWritten As Int32
            Dim ptBuf As Int32

            '\\ Get the required buffer size
            If Not GetJob(hPrinter, dwJobId, 1, ptBuf, 0, BytesWritten) Then
                If BytesWritten = 0 Then
                    Throw New Win32Exception
                    Exit Sub
                End If
            End If

            '\\ Allocate a buffer the right size
            If BytesWritten > 0 Then
                ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
            End If

            '\\ Populate the JOB_INFO_1 structure
            If Not GetJob(hPrinter, dwJobId, 1, ptBuf, BytesWritten, BytesWritten) Then
                Throw New Win32Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("GetJob for JOB_INFO_1 failed on handle: " & hPrinter.ToString & " for job: " & dwJobId, Me.GetType.ToString)
                End If
                Exit Sub
            Else
                Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
            End If

            '\\ Free the allocated memory
            Marshal.FreeHGlobal(CType(ptBuf, IntPtr))

        End Sub

        Public Sub New(ByVal lpJob As Int32)
            Marshal.PtrToStructure(New IntPtr(lpJob), Me)
        End Sub

    End Class
#End Region

#Region "DEVMODE STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Friend Class DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, Sizeconst:=32)> Public pDeviceName As String
        Public dmSpecVersion As Short
        Public dmDriverVersion As Short
        Public dmSize As Short
        Public dmDriverExtra As Short
        Public dmFields As Integer
        Public dmOrientation As Short
        Public dmPaperSize As Short
        Public dmPaperLength As Short
        Public dmPaperWidth As Short
        Public dmScale As Short
        Public dmCopies As Short
        Public dmDefaultSource As Short
        Public dmPrintQuality As Short
        Public dmColor As Short
        Public dmDuplex As Short
        Public dmYResolution As Short
        Public dmTTOption As Short
        Public dmCollate As Short
        <MarshalAs(UnmanagedType.ByValTStr, Sizeconst:=32)> Public dmFormName As String
        Public dmUnusedPadding As Short
        Public dmBitsPerPel As Integer
        Public dmPelsWidth As Integer
        Public dmPelsHeight As Integer
        Public dmNup As Integer
        Public dmDisplayFrequency As Integer
        Public dmICMMethod As Integer
        Public dmICMIntent As Integer
        Public dmMediaType As Integer
        Public dmDitherType As Integer
        Public dmReserved1 As Integer
        Public dmReserved2 As Integer
        Public dmPanningWidth As Integer
        Public dmPanningHeight As Integer


    End Class
#End Region

#Region "JOB_INFO_2 STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class JOB_INFO_2

        Public JobId As Int32
        <MarshalAs(UnmanagedType.LPTStr)> Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pMachineName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pUserName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pDocument As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pNotifyName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pDatatype As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pParameters As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pDriverName As String
        <MarshalAs(UnmanagedType.U4)> Public LPDeviceMode As Int32
        <MarshalAs(UnmanagedType.LPTStr)> Public pStatus As String
        Public lpSecurity As Int32
        <MarshalAs(UnmanagedType.U4)> Public Status As PrintJobStatuses
        Public Priority As Int32
        Public Position As Int32
        Public StartTime As Int32
        Public UntilTime As Int32
        Public TotalPage As Int32
        Public JobSize As Int32
        <MarshalAs(UnmanagedType.Struct)> Public Submitted As SYSTEMTIME
        Public Time As Int32
        Public PagesPrinted As Int32

#Region "Private member variables"
        Dim dmOut As New DEVMODE
#End Region

        Public Sub New()
            '\\ If this structure is not "Live" then a printer handle and job id are not used
        End Sub

        Public Sub New(ByVal hPrinter As Int32, _
                     ByVal dwJobId As Int32)

            Dim BytesWritten As Int32
            Dim ptBuf As Int32

            '\\ Get the required buffer size
            If Not GetJob(hPrinter, dwJobId, 2, ptBuf, 0, BytesWritten) Then
                If BytesWritten = 0 Then
                    Throw New Win32Exception
                    If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                        Trace.WriteLine("GetJob for JOB_INFO_2 failed on handle: " & hPrinter.ToString & " for job: " & dwJobId, Me.GetType.ToString)
                    End If
                    Exit Sub
                End If
            End If

            '\\ Allocate a buffer the right size
            If BytesWritten > 0 Then
                ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
            End If

            '\\ Populate the JOB_INFO_2 structure
            If Not GetJob(hPrinter, dwJobId, 2, ptBuf, BytesWritten, BytesWritten) Then
                Throw New Win32Exception
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("GetJob for JOB_INFO_2 failed on handle: " & hPrinter.ToString & " for job: " & dwJobId, Me.GetType.ToString)
                End If
            Else
                Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                '\\ And get the DEVMODE before the memory is freed...
                Marshal.PtrToStructure(New IntPtr(LPDeviceMode), dmOut)
            End If
            '\\ Free the allocated memory
            Marshal.FreeHGlobal(CType(ptBuf, IntPtr))

            '\\ Prevent Null Reference exceptions
            If pPrinterName Is Nothing Then pPrinterName = ""
            If pUserName Is Nothing Then pUserName = ""
            If pDocument Is Nothing Then pDocument = ""
            If pNotifyName Is Nothing Then pNotifyName = ""
            If pDatatype Is Nothing Then pDatatype = ""
            If pPrintProcessor Is Nothing Then pPrintProcessor = ""
            If pParameters Is Nothing Then pParameters = ""
            If pDriverName Is Nothing Then pDriverName = ""
            If pStatus Is Nothing Then pStatus = ""

        End Sub

        Public ReadOnly Property DeviceMode() As DEVMODE
            Get
                Return dmOut
            End Get
        End Property

    End Class
#End Region

#Region "PRINTER_DEFAULTS STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Friend Class PRINTER_DEFAULTS
        Public lpDataType As Int32
        Public lpDevMode As Int32
        <MarshalAs(UnmanagedType.U4)> Public DesiredAccess As PrinterAccessRights

#Region "Public constructor"
        '\\ If the device name is known we need to know what access rights we have to this device..
        Public Sub New(ByVal PrinterInfo As PrinterInformation)

            If PrinterInfo.CanLoggedInUserAdministerPrinter Then
                DesiredAccess = PrinterAccessRights.PRINTER_ALL_ACCESS
            Else
                DesiredAccess = PrinterAccessRights.PRINTER_ACCESS_USE
            End If

        End Sub

        Public Sub New()

            If LoggedInAsAdministrator() Then
                DesiredAccess = PrinterAccessRights.PRINTER_ALL_ACCESS
            Else
                DesiredAccess = PrinterAccessRights.PRINTER_ACCESS_USE
            End If

        End Sub

        Public Sub New(ByVal DefaultDesiredAccess As PrinterAccessRights)

            DesiredAccess = DefaultDesiredAccess

        End Sub

#End Region

    End Class
#End Region

#Region "PRINTER_INFO_1"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class PRINTER_INFO_1
        <MarshalAs(UnmanagedType.U4)> Public Flags As Int32
        <MarshalAs(UnmanagedType.LPWStr)> Public pDescription As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pComment As String

#Region "Public constructors"
        Public Sub New(ByVal hPrinter As Int32)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As Integer

            ptBuf = CInt(Marshal.AllocHGlobal(1))

            If Not GetPrinter(hPrinter, 1, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(New IntPtr(ptBuf))
                    ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
                    If GetPrinter(hPrinter, 1, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                    Else
#If ERROR_BUBBLING Then
                        Throw New Win32Exception()
#End If
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                Else
#If ERROR_BUBBLING Then
                    Throw New Win32Exception()
#End If
                End If
            Else
#If ERROR_BUBBLING Then
                Throw New Win32Exception()
#End If
            End If

        End Sub

        Public Sub New()

        End Sub
#End Region
    End Class
#End Region

#Region "PRINTER_INFO_2 STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class PRINTER_INFO_2
        <MarshalAs(UnmanagedType.LPWStr)> Public pServerName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pShareName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pPortName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDriverName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pComment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pLocation As String
        <MarshalAs(UnmanagedType.U4)> Public lpDeviceMode As Int32
        <MarshalAs(UnmanagedType.LPWStr)> Public pSeperatorFilename As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDataType As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pParameters As String
        <MarshalAs(UnmanagedType.U4)> Public lpSecurityDescriptor As Int32
        Public Attributes As Int32
        Public Priority As Int32
        Public DefaultPriority As Int32
        Public StartTime As Int32
        Public UntilTime As Int32
        Public Status As Int32
        Public JobCount As Int32
        Public AveragePPM As Int32

#Region "Private member variables"
        Private dmOut As New DEVMODE
#End Region

#Region "Public constructors"
        Public Sub New(ByVal hPrinter As Int32)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As Integer

            ptBuf = CInt(Marshal.AllocHGlobal(1))

            If Not GetPrinter(hPrinter, 2, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                    ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
                    If GetPrinter(hPrinter, 2, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                        '\\ Fill any missing members
                        If pServerName Is Nothing Then
                            pServerName = ""
                        End If
                        '\\ If the devicemode is available, get it
                        If lpDeviceMode > 0 Then
                            Marshal.PtrToStructure(New IntPtr(lpDeviceMode), dmOut)
                        End If
                    Else
#If ERROR_BUBBLING Then
                        Throw New Win32Exception()
#End If
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                Else
#If ERROR_BUBBLING Then
                    Throw New Win32Exception()
#End If
                End If
            Else
#If ERROR_BUBBLING Then
                Throw New Win32Exception()
#End If
            End If

        End Sub

        Public ReadOnly Property DeviceMode() As DEVMODE
            Get
                Return dmOut
            End Get
        End Property

        Public Sub New()

        End Sub


#End Region

    End Class

#End Region

#Region "PRINTER_INFO_3 STRUCTURE"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class PRINTER_INFO_3
        <MarshalAs(UnmanagedType.U4)> Public pSecurityDescriptor As Integer

#Region "Public constructors"
        Public Sub New(ByVal hPrinter As Int32)

            Dim BytesWritten As Int32
            Dim ptBuf As Int32

            ptBuf = CInt(Marshal.AllocHGlobal(1))

            If Not GetPrinter(hPrinter, 3, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                    ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
                    If GetPrinter(hPrinter, 3, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                    Else
                        Throw New Win32Exception
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                Else
                    Throw New Win32Exception
                End If
            Else
                Throw New Win32Exception
            End If

        End Sub

        Public Sub New()

        End Sub

#End Region

    End Class

#End Region

#Region "PRINTER_NOTIFY_OPTIONS_TYPE STRUCTURE"
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PRINTER_NOTIFY_OPTIONS_TYPE
        Public wType As Int16
        Public wReserved0 As Int16
        Public dwReserved1 As Int32
        Public dwReserved2 As Int32
        Public Count As Int32
        Public pFields As Int32  '\\ Pointer to an array of (Count * PRINTER_NOTIFY_INFO_DATA) 
    End Structure
#End Region

#Region "PRINTER_NOTIFY_INFO STRUCTURE"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PRINTER_NOTIFY_INFO
        Public Version As Int32
        Public Flags As Int32
        Public Count As Int32

    End Class
#End Region

#Region "DRIVER_INFO_2"
    <StructLayout(LayoutKind.Sequential), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class DRIVER_INFO_2
        <MarshalAs(UnmanagedType.U4)> Public cVersion As SpoolerApiConstantEnumerations.PrinterDriverOperatingSystemVersion
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pEnvironment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDriverPath As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDatafile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pConfigFile As String


#Region "Public constructors"
        Public Sub New(ByVal hPrinter As Int32)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As Int32

            ptBuf = CInt(Marshal.AllocHGlobal(1))

            If Not GetPrinterDriver(hPrinter, Nothing, 2, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                    ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
                    If GetPrinterDriver(hPrinter, Nothing, 2, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                    Else
                        Throw New Win32Exception
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                Else
                    Throw New Win32Exception
                End If
            Else
                Throw New Win32Exception
            End If

        End Sub

        Public Sub New()

        End Sub
#End Region

    End Class
#End Region

#Region "DRIVER_INFO_3"
    <StructLayout(LayoutKind.Sequential), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class DRIVER_INFO_3
        <MarshalAs(UnmanagedType.U4)> Public cVersion As SpoolerApiConstantEnumerations.PrinterDriverOperatingSystemVersion
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pEnvironment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDriverPath As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDatafile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pConfigFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pHelpFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDependentFiles As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pMonitorName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDefaultDataType As String

#Region "Public constructors"
        Public Sub New(ByVal hPrinter As Int32)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As Integer

            ptBuf = CInt(Marshal.AllocHGlobal(1))

            If Not GetPrinterDriver(hPrinter, Nothing, 3, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                    ptBuf = CInt(Marshal.AllocHGlobal(BytesWritten))
                    If GetPrinterDriver(hPrinter, Nothing, 3, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(New IntPtr(ptBuf), Me)
                    Else
                        Throw New Win32Exception
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(CType(ptBuf, IntPtr))
                Else
                    Throw New Win32Exception
                End If
            Else
                Throw New Win32Exception
            End If

        End Sub

        Public Sub New()

        End Sub
#End Region

    End Class
#End Region

#Region "MONITOR_INFO_2"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class MONITOR_INFO_2
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pEnvironment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDLLName As String

#Region "Public constructors"
        Public Sub New(ByVal lpMonitorInfo2 As Integer)
            Marshal.PtrToStructure(New IntPtr(lpMonitorInfo2), Me)
        End Sub

        Public Sub New()

        End Sub
#End Region

    End Class
#End Region

#Region "PRINTPROCESSOR_INFO_1"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PRINTPROCESSOR_INFO_1
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
    End Class
#End Region

#Region "DATATYPES_INFO_1"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class DATATYPES_INFO_1
        <MarshalAs(UnmanagedType.LPWStr)> Public pName As String
    End Class
#End Region

#Region "PORT_INFO_1"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_1
        <MarshalAs(UnmanagedType.LPWStr)> Public pPortName As String
    End Class
#End Region

#Region "PORT_INFO_2"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_2
        <MarshalAs(UnmanagedType.LPWStr)> Public pPortName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pMonitorName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDescription As String
        <MarshalAs(UnmanagedType.U4)> Public PortType As PortTypes
        <MarshalAs(UnmanagedType.U4)> Public Reserved As Int32

    End Class
#End Region

#Region "PORT_INFO_3"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_3
        <MarshalAs(UnmanagedType.U4)> Public dwStatus As Long
        <MarshalAs(UnmanagedType.LPWStr)> Public pszStatus As String
        <MarshalAs(UnmanagedType.U4)> Public dwSeverity As Long

    End Class
#End Region

#Region "FORM_INFO_1"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class FORM_INFO_1
        <MarshalAs(UnmanagedType.U4)> Public Flags As Int32 'FormTypeFlags
        <MarshalAs(UnmanagedType.LPWStr)> Public Name As String
        <MarshalAs(UnmanagedType.U4)> Public Width As Int32
        <MarshalAs(UnmanagedType.U4)> Public Height As Int32
        <MarshalAs(UnmanagedType.U4)> Public Left As Int32
        <MarshalAs(UnmanagedType.U4)> Public Top As Int32
        <MarshalAs(UnmanagedType.U4)> Public Right As Int32
        <MarshalAs(UnmanagedType.U4)> Public Bottom As Int32

    End Class

#End Region

#Region "DOC_INFO_1"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class DOC_INFO_1
        <MarshalAs(UnmanagedType.LPWStr)> Public DocumentName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public OutputFilename As String
        <MarshalAs(UnmanagedType.LPWStr)> Public DataType As String
    End Class
#End Region

End Namespace

