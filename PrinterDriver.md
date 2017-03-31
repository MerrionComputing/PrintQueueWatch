# PrinterDriver class
Specifies the settings for a printer driver (the DLL which provides the interface between the hardware and the spooler system)

## Example code (VB.NET)
Lists all the printer drivers on the current server
{{
        Dim server As New PrintServer

        Dim Driver As PrinterDriver
        Me.ListBox1.Items.Clear()
        For Each Driver In server.PrinterDrivers
            Me.ListBox1.Items.Add(Driver.Name)
        Next
}}

## Properties
* **OperatingSystemVersion** - The operating system version for which this driver was written
* **Name** (String) - The unique name by which this printer driver is known
* **Environment** (String) - The environment for which the driver was written. For example, "Windows NT x86", "Windows NT R4000", "Windows NT Alpha_AXP", or "Windows 4.0"
* **DriverPath** (String) - The file name or full path and file name for the file that contains the device driver, for example, "c:\drivers\pscript.dll"
* **DataFile** (String) - The file name or a full path and file name for the file that contains driver data, for example, "c:\drivers\Qms810.ppd" 
* **ConfigurationFile** (String) - The file name or a full path and file name for the device-driver's configuration .dll
* **HelpFile** (String) - The file name or a full path and file name for the device driver's help file.
* **MonitorName** (String) - The name of a printer language monitor attached to this driver (for example, "PJL monitor")
* **DefaultDataType** (String) - The default data type used by this printer driver in writing spool files for print jobs (see [DataType](DataType) class)

## Underlying API declarations
{{
   <StructLayout(LayoutKind.Sequential), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class DRIVER_INFO_3
        <MarshalAs(UnmanagedType.U4)> Public cVersion As SpoolerApiConstantEnumerations.PrinterDriverOperatingSystemVersion
        <MarshalAs(UnmanagedType.LPStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pEnvironment As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDriverPath As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDatafile As String
        <MarshalAs(UnmanagedType.LPStr)> Public pConfigFile As String
        <MarshalAs(UnmanagedType.LPStr)> Public pHelpFile As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDependentFiles As String
        <MarshalAs(UnmanagedType.LPStr)> Public pMonitorName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDefaultDataType As String

#Region "Public constructors"
        Public Sub New(ByVal hPrinter As IntPtr)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As IntPtr

            ptBuf = Marshal.AllocHGlobal(1)

            If Not GetPrinterDriver(hPrinter, Nothing, 3, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(ptBuf)
                    ptBuf = Marshal.AllocHGlobal(BytesWritten)
                    If GetPrinterDriver(hPrinter, Nothing, 3, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(ptBuf, Me)
                    Else
                        Throw New Win32Exception
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(ptBuf)
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
}}