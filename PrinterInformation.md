# PrinterInformation class

Class which holds the settings for a single printer 

## Code example (VB.Net)
Lists the name and location of all the printers on this server
{{
        Dim foo As New PrintServer
        For Each fp As PrinterInformation In foo.Printers
            Trace.WriteLine(String.Format("Printer {0} is located at {1) ", fp.PrinterName, fp.Location))
        Next
}}

## Properties
* **ServerName** (String) - The name of the server on which this printer is installed.  (This value may be blank if the printer is attached to the current machine)
* **PrinterName** (String) - The unique name by which the printer is known (also referred to as DeviceName)
* **ShareName** (String) - The name that the printer is shared as (if it is shared)
* **PortName** (String) - The name of the port that the printer is attached to (see the [Port](Port) class for more info)
* **DriverName** (String) - The name of the printer driver software used by this printer (see the [PrinterDriver](PrinterDriver) class for more info)
* **Comment** (String) - The administrator defined comment for this printer
* **Location** (String) - The administrator defined location for this printer
* **SeperatorFilename** (String) - The name of the file (if any) that is printed to seperate print jobs on this printer
* **PrintProcessor** (String) - The name of the print processor associated to this printer (see the [PrintProcessor](PrintProcessor) class for more info)
* **DefaultDataType** (String) - The default spool data type (e.g. RAW, EMF etc.) used by this printer (see the [DataType](DataType) class for more info)
* **Parameters** (String) - Additional configuration parameters used when printing on this printer
* **IsDefault** (Boolean) - If this printer is the default printer on this machine
* **IsShared** (Boolean) - If the printer is being shared
* **IsNetworkPrinter** (Boolean) - If the printer is a network attached device
* **IsLocalPrinter** (Boolean) - If the printer is attached to this local machine
* **Priority** (Int32) - The default priority of print jobs sent to this printer.  Priority can range from 1 (lowest) to 99 (highest). 
* **IsReady** (Boolean) - True if the printer is status ready to print
* **IsDoorOpen** (Boolean) - True if the printer is paused because the printer door is open
* **IsInError** (Boolean) - True if the printer has an error
* **IsInitialising** (Boolean) - True if the printer is initialising
* **IsAwaitingManualFeed** (Boolean) - True if the printer is paused waiting for the user to insert a manual paper feed
* **IsOutOfToner** (Boolean) - True if the printer is stalled because it is out of toner or ink
* **IsTonerLow** (Boolean) - True if the printer is stalled because it is low on toner or ink
* **IsUnavailable** (Boolean) - True if the printer is currently unnavailable
* **IsOffline** (Boolean) - True if the printer is offline
* **IsOutOfMemory** (Boolean) - True if the printer is stalled because it has run out of memory
* **IsOutputBinFull** (Boolean) - True if the printer is stalled because it's output tray is full
* **IsPaperJammed** (Boolean) - True if the printer is stalled because it has a paper jam
* **IsOutOfPaper** (Boolean) - True if the printer is stalled because it is out of paper
* **Paused** (Boolean) - True if the printer is paused.  This property can be set to pause or unpause the printer.
* **IsDeletingJob** (Boolean) - True if the printer is deleting a job
* **IsInPowerSave** (Boolean) - True if the printer is in power saving mode
* **IsPrinting** (Boolean) - True if the printer is currently printing a job
* **IsWaitingOnUserIntervention** (Boolean) - True if the printer is stalled awaiting manual intervention
* **IsWarmingUp** (Boolean) - True if the printer is warming up to be ready to print
* **JobCount** (Int32) - The number of print jobs queued to be printed by this printer
* **AveragePagesPerMonth** (Int32) - The average throughput of this printer in pages per month
* **TimeWindow** - The time window within which jobs can be scheduled against this printer
* **DefaultPaperSource** - The default tray used by this printer
* **Copies** (Int32) - The number of copies of each print job to produce by default
* **Landscape** (Boolean) - True if the printer orientation is set to Landscape
* **Colour** (Boolean) - True if a colour printer is set to print in colour, false for monochrome
* **Collate** (Boolean) - Specifies whether collation should be used when printing multiple copies
* **PrintQuality** - The default print quality for print jobs on this printer
* **Scale** - The scale (percentage) to print the page at
* **PrintJobs** - The collection of print jobs queued to print on the printer (see the [PrintJob](PrintJob) class)
* **Monitored** - Sets whether or not events occuring on this printer are raised by the component
* **PauseAllNewJobs** - If true and the printer is being monitored, all print jobs are paused as they are added to the spool queue
* **PrinterForms** - Returns the collection of print forms installed on the printer
* **Availability** - The time range that the printer is available for

## Methods
* **PausePrinting** - Pauses the printer

## Underlying API code
{{
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class PRINTER_INFO_2
        <MarshalAs(UnmanagedType.LPStr)> Public pServerName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pShareName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pPortName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDriverName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pComment As String
        <MarshalAs(UnmanagedType.LPStr)> Public pLocation As String
        <MarshalAs(UnmanagedType.U4)> Public lpDeviceMode As Int32
        <MarshalAs(UnmanagedType.LPStr)> Public pSeperatorFilename As String
        <MarshalAs(UnmanagedType.LPStr)> Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDataType As String
        <MarshalAs(UnmanagedType.LPStr)> Public pParameters As String
        <MarshalAs(UnmanagedType.U4)> Public lpSecurityDescriptor As Int32
        Public Attributes As Int32
        Public Priority As Int32
        Public DefaultPriority As Int32
        Public StartTime As Int32
        Public UntilTime As Int32
        Public Status As Int32
        Public JobCount As Int32
        Public AveragePPM As Int32

        Private dmOut As New DEVMODE

        Public Sub New(ByVal hPrinter As IntPtr)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As IntPtr

            ptBuf = Marshal.AllocHGlobal(1)

            If Not GetPrinter(hPrinter, 2, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(ptBuf)
                    ptBuf = Marshal.AllocHGlobal(BytesWritten)
                    If GetPrinter(hPrinter, 2, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(ptBuf, Me)
                        '\\ Fill any missing members
                        If pServerName Is Nothing Then
                            pServerName = ""
                        End If
                        '\\ If the devicemode is available, get it
                        If lpDeviceMode > 0 Then
                            Dim ptrDevMode As New IntPtr(lpDeviceMode)
                            Marshal.PtrToStructure(ptrDevMode, dmOut)
                        End If
                    Else
                       Throw New Win32Exception()
                   End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(ptBuf)
                Else
                   Throw New Win32Exception()
                End If
            Else
                Throw New Win32Exception()
            End If

        End Sub

        Public ReadOnly Property DeviceMode() As DEVMODE
            Get
                Return dmOut
            End Get
        End Property

        Public Sub New()

        End Sub

    End Class

}}


 