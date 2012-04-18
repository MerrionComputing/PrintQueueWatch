Imports System.Runtime.InteropServices
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

'\\--PrinterNotifyOptionsType---------------------------------------------------
'\\ Implements the API structure PRINTER_NOTIFY_OPTIONS_TYPE used in monitoring
'\\ a print queue using FindFirstPrinterChangeNotification and 
'\\ FindNextPrinterChangeNotification loop
'\\ (c) 2002-2003 Merrion Computing Ltd
'\\     http://www.merrioncomputing.com
'\\------------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)> _
Public Class PrinterNotifyOptionsType
    Implements IDisposable

    Public wJobType As Int16
    Public wJobReserved0 As Int16
    Public dwJobReserved1 As Int32
    Public dwJobReserved2 As Int32
    Public JobFieldCount As Int32
    Public pJobFields As IntPtr
    Public wPrinterType As Int16
    Public wPrinterReserved0 As Int16
    Public dwPrinterReserved1 As Int32
    Public dwPrinterReserved2 As Int32
    Public PrinterFieldCount As Int32
    Public pPrinterFields As IntPtr

#Region "Public Enumerated Types"

    Private Const JOB_FIELDS_COUNT As Int32 = 24
    Private Const PRINTER_FIELDS_COUNT As Int32 = 8

#End Region

    Private Sub SetupFields(ByVal MinimumJobInfoRequired As Boolean)

        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("SetupFields()", Me.GetType.ToString)
        End If
        wJobType = CShort(Printer_Notification_Types.JOB_NOTIFY_TYPE)
        wPrinterType = CShort(Printer_Notification_Types.PRINTER_NOTIFY_TYPE)

        '\\ Free up the global memory
        If pJobFields.ToInt64 <> 0 Then
            Marshal.FreeHGlobal(pJobFields)
        End If
        If pPrinterFields.ToInt64 <> 0 Then
            Marshal.FreeHGlobal(pPrinterFields)
        End If

        If MinimumJobInfoRequired Then
            JobFieldCount = 2
        Else
            JobFieldCount = JOB_FIELDS_COUNT
        End If

        pJobFields = Marshal.AllocHGlobal((JOB_FIELDS_COUNT * 2) - 1)
        Dim PrintJobFields(JOB_FIELDS_COUNT - 1) As Short

        PrintJobFields(0) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DOCUMENT
        PrintJobFields(1) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_STATUS
        If Not MinimumJobInfoRequired Then
            PrintJobFields(2) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_MACHINE_NAME
            PrintJobFields(3) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PORT_NAME
            PrintJobFields(4) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_USER_NAME
            PrintJobFields(5) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_NOTIFY_NAME
            PrintJobFields(6) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DATATYPE
            PrintJobFields(7) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRINT_PROCESSOR
            PrintJobFields(8) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PARAMETERS
            PrintJobFields(9) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DRIVER_NAME
            PrintJobFields(10) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DEVMODE
            PrintJobFields(11) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_STATUS_STRING
            PrintJobFields(12) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR
            PrintJobFields(13) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRINTER_NAME
            PrintJobFields(14) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRIORITY
            PrintJobFields(15) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_POSITION
            PrintJobFields(16) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_SUBMITTED
            PrintJobFields(17) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_START_TIME
            PrintJobFields(18) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_UNTIL_TIME
            PrintJobFields(19) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_TIME
            PrintJobFields(20) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_TOTAL_PAGES
            PrintJobFields(21) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PAGES_PRINTED
            PrintJobFields(22) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_TOTAL_BYTES
            PrintJobFields(23) = Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_BYTES_PRINTED
        End If
        Marshal.Copy(PrintJobFields, 0, pJobFields, PrintJobFields.GetLength(0))

        '\\ Request less printer notification details for economy sake...
        If MinimumJobInfoRequired Then
            PrinterFieldCount = 1
        Else
            PrinterFieldCount = PRINTER_FIELDS_COUNT
        End If

        pPrinterFields = Marshal.AllocHGlobal((PRINTER_FIELDS_COUNT - 1) * 2)
        Dim PrinterFields(PRINTER_FIELDS_COUNT - 1) As Short
        PrinterFields(0) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_STATUS

        If Not MinimumJobInfoRequired Then
            PrinterFields(1) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_CJOBS
            PrinterFields(2) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_ATTRIBUTES
            PrinterFields(3) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_COMMENT
            PrinterFields(4) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_DEVMODE
            PrinterFields(5) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_LOCATION
            PrinterFields(6) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_SECURITY_DESCRIPTOR
            PrinterFields(7) = Printer_Notify_Field_Indexes.PRINTER_NOTIFY_FIELD_SEPFILE
        End If
        Marshal.Copy(PrinterFields, 0, pPrinterFields, PrinterFields.GetLength(0))

    End Sub

    Public Sub New(ByVal MinimumJobInfoRequired As Boolean)
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & MinimumJobInfoRequired.ToString & ")", Me.GetType.ToString)
        End If
        Call SetupFields(MinimumJobInfoRequired)

    End Sub


    Public Sub ReleaseResources()
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("ReleaseResources()", Me.GetType.ToString)
        End If
        If pJobFields.ToInt64 <> 0 Then
            Marshal.FreeHGlobal(CType(pJobFields, IntPtr))
            pJobFields = Nothing
        End If
        If pPrinterFields.ToInt64 <> 0 Then
            Marshal.FreeHGlobal(CType(pPrinterFields, IntPtr))
            pPrinterFields = Nothing
        End If
    End Sub

#Region "IDisposable interface implementation"
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")>
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("Dispose()", Me.GetType.ToString)
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            '\\ Free up the global memory
            Call ReleaseResources()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region
End Class

'\\--PrinterNotifyOptions-------------------------------------------------------
'\\ Implements the API structure PRINTER_NOTIFY_OPTIONS used in monitoring
'\\ a print queue using FindFirstPrinterChangeNotification and 
'\\ FindNextPrinterChangeNotification loop
'\\ (c) 2002-2003 Merrion Computing Ltd
'\\     http://www.merrioncomputing.com
'\\------------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)> _
Public Class PrinterNotifyOptions
    Implements IDisposable

    Public dwVersion As Int32
    Public dwFlags As Int32
    Public Count As Int32
    Public lpTypes As IntPtr

    Public Sub New(ByVal MinimumJobInfoRequired As Boolean)

        If PrinterMonitorComponent.ComponentTraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & MinimumJobInfoRequired.ToString & ")", Me.GetType.ToString)
        End If

        '\\ As it stands, version is always 2
        dwVersion = 2
        Count = 2
        Dim pJobTypes As PrinterNotifyOptionsType
        Dim BytesNeeded As Integer

        pJobTypes = New PrinterNotifyOptionsType(MinimumJobInfoRequired)
        BytesNeeded = (2 + pJobTypes.JobFieldCount + pJobTypes.PrinterFieldCount) * 2

        lpTypes = Marshal.AllocHGlobal(BytesNeeded)

        Marshal.StructureToPtr(pJobTypes, lpTypes, True)

    End Sub

#Region "IDisposable interface"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            '\\ Free up the global memory
            If lpTypes.ToInt64 <> 0 Then
                Marshal.FreeHGlobal(lpTypes)
            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

#End Region

End Class
