Public Class Form_PrintQueueWatchTest

    Private Sub Form_PrintQueueWatchTest_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        ' Don't leave any printers being monitored when the form is gone...
        Me.PrinterMonitorComponent1.Disconnect()
    End Sub

    Private Sub Form_PrintQueueWatchTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each p In New PrinterQueueWatch.PrinterInformationCollection()
            Me.CheckedListBox_Printers.Items.Add(p.PrinterName)
        Next

    End Sub


    Private Sub CheckedListBox_Printers_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox_Printers.ItemCheck

        If (e.NewValue = CheckState.Checked) Then
            Me.PrinterMonitorComponent1.AddPrinter(CheckedListBox_Printers.Items(e.Index))
        Else
            Me.PrinterMonitorComponent1.RemovePrinter(CheckedListBox_Printers.Items(e.Index))
        End If

    End Sub



    Private Sub PrinterMonitorComponent1_JobAdded(sender As Object, e As PrintJobEventArgs) Handles PrinterMonitorComponent1.JobAdded
        System.Diagnostics.Trace.TraceInformation("Job added " & e.PrintJob.JobId & " called " & e.PrintJob.Document & " on " & e.PrintJob.PrinterName)

        'Do any other fundctionality here...

    End Sub


    Private Sub PrinterMonitorComponent1_JobDeleted(sender As Object, e As PrintJobEventArgs) Handles PrinterMonitorComponent1.JobDeleted
        System.Diagnostics.Trace.TraceInformation("Job deleted " & e.PrintJob.JobId & " called " & e.PrintJob.Document & " on " & e.PrintJob.PrinterName)

        'Do any other fundctionality here...
    End Sub


    Private Sub PrinterMonitorComponent1_PrinterInformationChanged(sender As Object, e As PrinterEventArgs) Handles PrinterMonitorComponent1.PrinterInformationChanged
        System.Diagnostics.Trace.TraceInformation("Printer information changed " & e.PrinterInformation.PrinterName & " - " & e.PrinterChangeFlags.ToString())
    End Sub
End Class
