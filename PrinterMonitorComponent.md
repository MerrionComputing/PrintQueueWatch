# PrinterMonitorComponent

The PrinterMonitorComponent allows you to set one or more printers to monitor and raises events in response to events which occur on these printers.  Events include a  new job being added, pages being printed, job settings changes and a job being deleted from the print queue.



## Code example (VB.NET)
{{
    Private WithEvents pmon As New PrinterMonitorComponent

    Public Sub New()
        MyBase.New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Monitor all the installed printers 
        For Each p As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            pmon.AddPrinter(p)
            pmon.PrinterInformation(p).PauseAllNewJobs = False
        Next p

    End Sub

    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As PrintJobEventArgs) Handles pmon.JobAdded

        With CType(e, PrintJobEventArgs)
            Trace.WriteLine("New job added " & .PrintJob.UserName)
        End With

    End Sub
}}

## Events
[JobAdded](JobAdded) - Raised when a new print job is added to a printer queue
[JobWritten](JobWritten) - Raised when a job is written to on one of the print spool queues being monitored
[JobDeleted](JobDeleted) - Raised when a print job is removed from the queue - either deleted by the user or printed by the printer
[JobSet](JobSet) - Raised when a job's properties are changed in one of the print spool queues being monitored
[PrinterInformationChanged](PrinterInformationChanged) - Raised when the settings on the printer being monitored change

## Properties
[Monitoring](Monitoring) (Boolean) - True if the component is monitoring any printers
[DeviceName](DeviceName) (String) - The name of the printer to monitor.  _If multiple printers are to be monitored then the [AddPrinter](AddPrinter) method is used instead_
[PrintJobs](PrintJobs) ([PrintJobCollection](PrintJobCollection)) - The collection of print jobs in the queue of the named device
[PrinterInformation](PrinterInformation) - The printer settings of the named device
[JobCount](JobCount) (Int32) - The number of print jobs in the queue of the named device


## Methods
[AddPrinter](AddPrinter) - Adds a named printer to the collection of printers to monitor
[RemovePrinter](RemovePrinter) - Removes a named printer from the collection of printers being monitored
[Disconnect](Disconnect) - Stops all printer monitoring

## Other settings
[ComponentTraceSwitch](ComponentTraceSwitch) - Controls the level of trace information emmitted from the component 
