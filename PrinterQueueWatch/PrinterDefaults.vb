Imports System.Runtime.InteropServices
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

<StructLayout(LayoutKind.Sequential)> _
Friend Class PrinterDefaults

#Region "Public interface"

    Public DataType As String
    Public lpDevMode As Int32
    Public DesiredAccess As PrinterAccessRights

#End Region

End Class
