<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_PrintQueueWatchTest
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CheckedListBox_Printers = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PrinterMonitorComponent1 = New PrinterQueueWatch.PrinterMonitorComponent()
        Me.SuspendLayout()
        '
        'CheckedListBox_Printers
        '
        Me.CheckedListBox_Printers.FormattingEnabled = True
        Me.CheckedListBox_Printers.Location = New System.Drawing.Point(3, 27)
        Me.CheckedListBox_Printers.Name = "CheckedListBox_Printers"
        Me.CheckedListBox_Printers.Size = New System.Drawing.Size(218, 304)
        Me.CheckedListBox_Printers.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Printer"
        '
        'PrinterMonitorComponent1
        '
        Me.PrinterMonitorComponent1.MonitorJobAddedEvent = True
        Me.PrinterMonitorComponent1.MonitorJobDeletedEvent = True
        Me.PrinterMonitorComponent1.MonitorJobSetEvent = True
        Me.PrinterMonitorComponent1.MonitorJobWrittenEvent = True
        Me.PrinterMonitorComponent1.MonitorPrinterChangeEvent = True
        '
        'Form_PrintQueueWatchTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(671, 344)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CheckedListBox_Printers)
        Me.Name = "Form_PrintQueueWatchTest"
        Me.Text = "PQW Test"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckedListBox_Printers As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PrinterMonitorComponent1 As PrinterQueueWatch.PrinterMonitorComponent

End Class
