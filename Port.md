# Port class

Represents information about the port to which a printer is attached

## Code example (VB.NET)
Lists all the ports on the current print server
{{
        Dim server As New PrintServer

        ListBox1.Items.Clear()
        For ps As Integer = 0 To server.Ports.Count - 1
            Me.ListBox1.Items.Add(server.Ports(ps).Name)
        Next
}}

## Properties
* **Servername** (String)- The name of the server on which this port resides
* **Name** (String)- The name of the port (e.g. "LPT1:")
* **MonitorName** (String)- Identifies an installed monitor (for example, "PJL monitor")
* **Description** (String)- More detailed description of this printer port
* **Read** (Boolean) - Does the port supports read functionality
* **Write** (Boolean) - Does the port supports write operations
* **Redirected** (Boolean) - Is the port redirected to another port
* **NetAttached** (boolean) - Is the port network attached

## Underlying API code
{{
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_1
        <MarshalAs(UnmanagedType.LPStr)> Public pPortName As String
    End Class
#End Region

#Region "PORT_INFO_2"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_2
        <MarshalAs(UnmanagedType.LPStr)> Public pPortName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pMonitorName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pDescription As String
        <MarshalAs(UnmanagedType.U4)> Public PortType As PortTypes
        <MarshalAs(UnmanagedType.U4)> Public Reserved As Int32

    End Class
#End Region

#Region "PORT_INFO_3"
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class PORT_INFO_3
        <MarshalAs(UnmanagedType.U4)> Public dwStatus As Long
        <MarshalAs(UnmanagedType.LPStr)> Public pszStatus As String
        <MarshalAs(UnmanagedType.U4)> Public dwSeverity As Long

    End Class
#End Region

    <DllImport("winspool.drv", EntryPoint:="EnumPorts", _
SetLastError:=True, CharSet:=CharSet.Ansi, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
Public Function EnumPorts( _
                       <InAttribute()> ByVal ServerName As IntPtr, _
                       <InAttribute()> ByVal Level As Int32, _
                       <OutAttribute()> ByVal pbOut As IntPtr, _
                       <InAttribute()> ByVal cbIn As Int32, _
                       <OutAttribute()> ByRef pcbNeeded As Int32, _
                       <OutAttribute()> ByRef pcReturned As Int32 _
                ) As Boolean

    End Function
}}