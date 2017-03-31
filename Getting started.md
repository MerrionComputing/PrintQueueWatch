# Getting started with the PrintQueueWatch component 

The getting started page is a quick and dirty subset of the documentation based on the format:  _"I want to ....." _

## I want to ... get a list of all the documents that are printed on one printer

In order to get a list of the documents printed on a printer you need to monitor that printer and when you get a [JobWritten](JobWritten) event check if the .Printed property is true and if so, make a note of that print job's details.

**Step 1 : Add the PrintQueueWatch component to your project**
To do this slect Project -> References and add a reference to the latest PrintQueueWatch.dll to your project

**Step 2: Add an instance of the component on your form**
Drag the component from the toolbox onto your form surface.  Visual Studio will create an instance and wire it up.

{"
Imports PrinterQueueWatch

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private WithEvents pmon As New PrinterMonitorComponent
"}

**Step 3: Add an event handler to handle the JobAdded event**
{"
    Private Sub pmon_JobAdded(ByVal sender As Object, ByVal e As PrintJobEventArgs) Handles pmon.JobAdded

        With CType(e, PrintJobEventArgs)
            Trace.WriteLine(.PrintJob.UserName)
        End With

    End Sub
"}

**Step 4: Add the printers to monitor**
In the form load event handler add all the local device names to the list of printers to monitor

{"
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        For Each p As PrinterInformation In New PrinterInformationCollection
            pmon.AddPrinter(p.PrinterName)
            Trace.WriteLine(String.Format("Printer {0} added to monitor", p.PrinterName))
        Next p

    End Sub
"}

**Step 5: Add code to close the monitoring nicely**
Add the following in the form unload handler
{"
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            pmon.Dispose()
        Catch EX As Exception
            Trace.WriteLine(EX.ToString)
        End Try
    End Sub
"}