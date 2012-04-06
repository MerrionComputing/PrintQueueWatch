'\\ --[PrinterNotifyInfoData]----------------------------------------
'\\ Reads the information buffer provided by the print spooler in
'\\ response to FindFirstPrinterChangeNotification and 
'\\ FindNextPrinterChangeNotification
'\\ (c) Merrion Computing Ltd 
'\\     http://www.merrioncomputing.com
'\\ -----------------------------------------------------------------
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

<StructLayout(LayoutKind.Sequential)> _
Public Class PrinterNotifyInfoData
    Public wType As Int16
    Public wField As Int16
    Public dwReserved As Int32
    Public dwId As Int32
    Public cbBuff As Int32
    Public pBuff As IntPtr

    Public Sub New(ByVal lpAddress As IntPtr)

        If PrinterNotifyInfoData.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & lpAddress.ToString & ")", Me.GetType.ToString)
        End If
        Marshal.PtrToStructure(lpAddress, Me)

    End Sub

    Public ReadOnly Property Type() As Printer_Notification_Types
        Get
            Return CType(wType, Printer_Notification_Types)
        End Get
    End Property

    Public ReadOnly Property Field() As Job_Notify_Field_Indexes
        Get
            Return CType(wField, Job_Notify_Field_Indexes)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Marshal.PtrToStringAnsi(pBuff)
    End Function

    Public Function ToInt32() As Int32
        Return cbBuff
    End Function

#Region "Tracing"
    Public Shared TraceSwitch As New TraceSwitch("PrinterNotifyInfoData", "Printer Monitor Component Tracing")
#End Region

End Class

<StructLayout(LayoutKind.Sequential)> _
Public Class PrinterNotifyInfo

#Region "Private member variables"

    Private colPrintJobs As ArrayList
    Private msInfo As New PRINTER_NOTIFY_INFO()
    Private mlPrinterInfoChanged As Integer = 0

#End Region

    Public Sub New(ByVal mhPrinter As IntPtr, ByVal lpAddress As IntPtr, ByRef PrintJobs As PrintJobCollection)

        If PrinterNotifyInfoData.TraceSwitch.TraceVerbose Then
            Trace.WriteLine("New(" & mhPrinter.ToString & "," & lpAddress.ToString & ")", Me.GetType.ToString)
        End If

        If Not lpAddress.ToInt32 = 0 Then
            '\\ Create the array list of jobs involved in this event
            colPrintJobs = New ArrayList

            '\\ Read the data of this printer notification event
            Marshal.PtrToStructure(lpAddress, msInfo)

            Dim nInfoDataItem As Integer
            '\\ Offset the pointer by the size of this class
            Dim lOffset As Integer = lpAddress.ToInt32 + Marshal.SizeOf(msInfo)

            '\\ Process the .adata array
            For nInfoDataItem = 0 To msInfo.Count - 1
                Dim itemdata As New PrinterNotifyInfoData(New IntPtr(lOffset))
                If itemdata.Type = Printer_Notification_Types.JOB_NOTIFY_TYPE Then

                    Dim pjThis As PrintJob
                    '\\ If this job is not on the printer job list, add it...
                    pjThis = PrintJobs.AddOrGetById(itemdata.dwId, mhPrinter)

                    If Not colPrintJobs.Contains(itemdata.dwId) Then
                        colPrintJobs.Add(itemdata.dwId)
                    End If


                    If Not pjThis Is Nothing Then
                        With pjThis
                            Select Case itemdata.Field
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRINTER_NAME
                                    .InitPrinterName = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_USER_NAME
                                    .InitUsername = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_MACHINE_NAME
                                    .InitMachineName = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DATATYPE
                                    .InitDataType = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DOCUMENT
                                    .InitDocument = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_DRIVER_NAME
                                    .InitDrivername = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_NOTIFY_NAME
                                    .InitNotifyUsername = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PAGES_PRINTED
                                    .InitPagesPrinted = itemdata.ToInt32
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PARAMETERS
                                    .InitParameters = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_POSITION
                                    .InitPosition = itemdata.ToInt32
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRINT_PROCESSOR
                                    .InitPrintProcessorName = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_PRIORITY
                                    .InitPriority = itemdata.ToInt32
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_STATUS
                                    .InitStatus = itemdata.ToInt32
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_STATUS_STRING
                                    .InitStatusDescription = Marshal.PtrToStringUni(itemdata.pBuff)
                                Case Job_Notify_Field_Indexes.JOB_NOTIFY_FIELD_TOTAL_PAGES
                                    .InitTotalPages = itemdata.ToInt32
                                Case Else
                                    '\\ These are not available except where the print job is "live" 
                            End Select
                        End With
                    End If
                Else
                    '\\ Printer Info changed event
                    mlPrinterInfoChanged = CInt(mlPrinterInfoChanged + (2 ^ itemdata.Field))
                End If
                lOffset = lOffset + Marshal.SizeOf(itemdata)
            Next

            '\\ And free the associated memory
            If Not FreePrinterNotifyInfo(lpAddress) Then
                If PrinterMonitorComponent.ComponentTraceSwitch.TraceError Then
                    Trace.WriteLine("FreePrinterNotifyInfo(" & lpAddress.ToString & ") failed", Me.GetType.ToString)
                End If
                Throw New Win32Exception
            End If
        End If

    End Sub

    Public ReadOnly Property Flags() As Int32
        Get
            Return (msInfo.Flags)
        End Get
    End Property

    Public ReadOnly Property PrintJobs() As ArrayList
        Get
            Return colPrintJobs
        End Get
    End Property

    Public ReadOnly Property PrinterInfoChanged() As Boolean
        Get
            Return (mlPrinterInfoChanged > 0)
        End Get
    End Property

    Friend ReadOnly Property PrinterInfoChangeFlags() As Integer
        Get
            Return mlPrinterInfoChanged
        End Get
    End Property

End Class
