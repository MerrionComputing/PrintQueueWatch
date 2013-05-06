'\\ --[SpoolerApi]---------------------------------------------------------------------------
'\\ Module for all the spooler related API calls, from winspool.drv
'\\ (c) 2003 Merrion Computing Ltd
'\\ http://www.merrioncomputing.com
'\\ ------------------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports PrinterQueueWatch.SpoolerStructs
Imports PrinterQueueWatch.SpoolerApiConstantEnumerations

Module SpoolerApi
    ' -- Notes: -------------------------------------------------------------------------------
    ' Always use <InAttribute()> and <OutAttribute()> to cut down on unnecessary marshalling
    ' -----------------------------------------------------------------------------------------
#Region "Api Declarations"

#Region "OpenPrinter"
    <DllImport("winspool.drv", EntryPoint:="OpenPrinter", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenPrinter(<InAttribute()> ByVal pPrinterName As String, _
                            <OutAttribute()> ByRef phPrinter As IntPtr, _
                               <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pDefault As PRINTER_DEFAULTS _
                               ) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="OpenPrinter", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenPrinter(<InAttribute()> ByVal pPrinterName As String, _
                              <OutAttribute()> ByRef phPrinter As IntPtr, _
                              <InAttribute()> ByVal pDefault As Integer _
                               ) As Boolean

    End Function



    <DllImport("winspool.drv", EntryPoint:="OpenPrinter", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenPrinter(<InAttribute()> ByVal pPrinterName As String, _
                                <OutAttribute()> ByRef phPrinter As IntPtr, _
                                <InAttribute()> ByVal pDefault As PrinterDefaults _
                                       ) As Boolean

    End Function


#End Region

#Region "ClosePrinter"
    <DllImport("winspool.drv", EntryPoint:="ClosePrinter", _
SetLastError:=True, _
ExactSpelling:=True, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClosePrinter(<InAttribute()> ByVal hPrinter As IntPtr) As Boolean

    End Function
#End Region

#Region "GetPrinter"
    <DllImport("winspool.drv", EntryPoint:="GetPrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetPrinter _
            (<InAttribute()> ByVal hPrinter As IntPtr, _
             <InAttribute()> ByVal Level As Int32, _
             <OutAttribute()> ByVal lpPrinter As IntPtr, _
             <InAttribute()> ByVal cbBuf As Int32, _
             <OutAttribute()> ByRef lpbSizeNeeded As IntPtr) As Boolean

    End Function
#End Region

#Region "EnumPrinters"
    <DllImport("winspool.drv", EntryPoint:="EnumPrinters", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPrinters(<InAttribute()> ByVal Flags As EnumPrinterFlags, _
                                 <InAttribute()> ByVal Name As String, _
                                 <InAttribute()> ByVal Level As Int32, _
                                 <OutAttribute()> ByVal lpBuf As Int32, _
                                 <InAttribute()> ByVal cbBuf As Int32, _
                                 <OutAttribute()> ByRef pcbNeeded As Int32, _
                                 <OutAttribute()> ByRef pcbReturned As Int32) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="EnumPrinters", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPrinters(<InAttribute()> ByVal Flags As EnumPrinterFlags, _
                             <InAttribute()> ByVal Name As Int32, _
                             <InAttribute()> ByVal Level As Int32, _
                             <OutAttribute()> ByVal lpBuf As Int32, _
                             <InAttribute()> ByVal cbBuf As Int32, _
                             <OutAttribute()> ByRef pcbNeeded As Int32, _
                             <OutAttribute()> ByRef pcbReturned As Int32) As Boolean

    End Function

#End Region

#Region "GetPrinterDriver"
    <DllImport("winspool.drv", EntryPoint:="GetPrinterDriver", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetPrinterDriver _
            (<InAttribute()> ByVal hPrinter As IntPtr, _
            <InAttribute()> ByVal pEnvironment As String, _
            <InAttribute()> ByVal Level As Int32, _
            <OutAttribute()> ByVal lpDriverInfo As IntPtr, _
            <InAttribute()> ByVal cbBuf As IntPtr, _
            <OutAttribute()> ByRef lpbSizeNeeded As IntPtr) As Boolean

    End Function
#End Region

#Region "EnumPrinterDrivers"
    <DllImport("winspool.drv", EntryPoint:="EnumPrinterDrivers", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPrinterDrivers(<InAttribute()> ByVal ServerName As String, _
                                       <InAttribute()> ByVal Environment As String, _
                                       <InAttribute()> ByVal Level As Int32, _
                                       <OutAttribute()> ByVal lpBuf As IntPtr, _
                                       <InAttribute()> ByVal cbBuf As Int32, _
                                       <OutAttribute()> ByRef pcbNeeded As Int32, _
                                       <OutAttribute()> ByRef pcbReturned As Int32) As Boolean

    End Function
#End Region

#Region "SetPrinter"
    <DllImport("winspool.drv", EntryPoint:="SetPrinter", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetPrinter _
             (<InAttribute()> ByVal hPrinter As IntPtr, _
             <InAttribute()> ByVal Level As Int32, _
             <InAttribute()> ByVal pPrinter As IntPtr, _
             <InAttribute()> ByVal Command As PrinterControlCommands) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="SetPrinter", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetPrinter _
         (<InAttribute()> ByVal hPrinter As IntPtr, _
         <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Level As PrinterInfoLevels, _
         <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinter As PRINTER_INFO_1, _
         <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Command As PrinterControlCommands) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="SetPrinter", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetPrinter _
             (<InAttribute()> ByVal hPrinter As IntPtr, _
             <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Level As PrinterInfoLevels, _
             <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinter As PRINTER_INFO_2, _
             <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Command As PrinterControlCommands) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="SetPrinter", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetPrinter _
         (<InAttribute()> ByVal hPrinter As IntPtr, _
         <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Level As PrinterInfoLevels, _
         <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinter As PRINTER_INFO_3, _
         <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Command As PrinterControlCommands) As Boolean

    End Function
#End Region

#Region "GetJob"
    <DllImport("winspool.drv", EntryPoint:="GetJob", _
 SetLastError:=True, CharSet:=CharSet.Auto, _
 ExactSpelling:=False, _
 CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetJob _
                (<InAttribute()> ByVal hPrinter As IntPtr, _
                 <InAttribute()> ByVal dwJobId As Int32, _
                 <InAttribute()> ByVal Level As Int32, _
                 <OutAttribute()> ByVal lpJob As IntPtr, _
                 <InAttribute()> ByVal cbBuf As Int32, _
                 <OutAttribute()> ByRef lpbSizeNeeded As IntPtr) As Boolean

    End Function



#End Region

#Region "FindFirstPrinterChangeNotification"


    <DllImport("winspool.drv", EntryPoint:="FindFirstPrinterChangeNotification", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindFirstPrinterChangeNotification _
                    (<InAttribute()> ByVal hPrinter As IntPtr, _
                     <InAttribute()> ByVal fwFlags As Int32, _
                     <InAttribute()> ByVal fwOptions As Int32, _
                     <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinterNotifyOptions As PrinterNotifyOptions _
                        ) As Microsoft.Win32.SafeHandles.SafeWaitHandle

    End Function

    <DllImport("winspool.drv", EntryPoint:="FindFirstPrinterChangeNotification", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function UnsafeFindFirstPrinterChangeNotification _
                    (<InAttribute()> ByVal hPrinter As IntPtr, _
                     <InAttribute()> ByVal fwFlags As Int32, _
                     <InAttribute()> ByVal fwOptions As Int32, _
                     <InAttribute()> ByVal pPrinterNotifyOptions As IntPtr _
                        ) As IntPtr

    End Function

    <DllImport("winspool.drv", EntryPoint:="FindFirstPrinterChangeNotification", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindFirstPrinterChangeNotification _
                (<InAttribute()> ByVal hPrinter As IntPtr, _
                 <InAttribute()> ByVal fwFlags As Int32, _
                 <InAttribute()> ByVal fwOptions As Int32, _
                 <InAttribute()> ByVal pPrinterNotifyOptions As IntPtr _
                    ) As Microsoft.Win32.SafeHandles.SafeWaitHandle

    End Function
#End Region

#Region "FindNextPrinterChangeNotification"
    <DllImport("winspool.drv", EntryPoint:="FindNextPrinterChangeNotification", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindNextPrinterChangeNotification _
                            (<InAttribute()> ByVal hChangeObject As IntPtr, _
                             <OutAttribute()> ByRef pdwChange As Int32, _
                             <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinterNotifyOptions As PrinterNotifyOptions, _
                             <OutAttribute()> ByRef lppPrinterNotifyInfo As IntPtr _
                                 ) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="FindNextPrinterChangeNotification", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindNextPrinterChangeNotification _
                        (<InAttribute()> ByVal hChangeObject As Microsoft.Win32.SafeHandles.SafeWaitHandle, _
                         <OutAttribute()> ByRef pdwChange As Int32, _
                         <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pPrinterNotifyOptions As PrinterNotifyOptions, _
                         <OutAttribute()> ByRef lppPrinterNotifyInfo As IntPtr _
                             ) As Boolean
    End Function
#End Region

#Region "FreePrinterNotifyInfo"
    <DllImport("winspool.drv", EntryPoint:="FreePrinterNotifyInfo", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function FreePrinterNotifyInfo _
             (<InAttribute()> ByVal lppPrinterNotifyInfo As IntPtr) As Boolean

    End Function
#End Region

#Region "FindClosePrinterChangeNotification"
    <DllImport("winspool.drv", EntryPoint:="FindClosePrinterChangeNotification", _
    SetLastError:=True, CharSet:=CharSet.Unicode, _
    ExactSpelling:=False, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindClosePrinterChangeNotification _
        (<InAttribute()> ByVal hChangeObject As Int32) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="FindClosePrinterChangeNotification", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function FindClosePrinterChangeNotification _
    (<InAttribute()> ByVal hChangeObject As Microsoft.Win32.SafeHandles.SafeWaitHandle) As Boolean

    End Function
#End Region

#Region "EnumJobs"
    <DllImport("winspool.drv", EntryPoint:="EnumJobs", _
 SetLastError:=True, CharSet:=CharSet.Unicode, _
 ExactSpelling:=False, _
 CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumJobs _
                (<InAttribute()> ByVal hPrinter As IntPtr, _
                 <InAttribute()> ByVal FirstJob As Int32, _
                 <InAttribute()> ByVal NumberOfJobs As Int32, _
 <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal Level As JobInfoLevels, _
 <OutAttribute()> ByVal pbOut As IntPtr, _
 <InAttribute()> ByVal cbIn As Int32, _
 <OutAttribute()> ByRef pcbNeeded As Int32, _
 <OutAttribute()> ByRef pcReturned As Int32 _
  ) As Boolean


    End Function
#End Region

#Region "SetJob"
    <DllImport("winspool.drv", EntryPoint:="SetJob", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetJob _
                    (<InAttribute()> ByVal hPrinter As IntPtr, _
                     <InAttribute()> ByVal dwJobId As Int32, _
                     <InAttribute()> ByVal Level As Int32, _
                     <InAttribute()> ByVal lpJob As IntPtr, _
                     <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal dwCommand As PrintJobControlCommands _
                    ) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="SetJob", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetJob _
                    (ByVal hPrinter As IntPtr, _
                     ByVal dwJobId As Int32, _
                     ByVal Level As Int32, _
                     <MarshalAs(UnmanagedType.LPStruct)> ByVal lpJob As JOB_INFO_1, _
                     ByVal dwCommand As PrintJobControlCommands _
                    ) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="SetJob", _
 SetLastError:=True, CharSet:=CharSet.Unicode, _
 ExactSpelling:=False, _
 CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetJob _
                (ByVal hPrinter As IntPtr, _
                 ByVal dwJobId As Int32, _
                 ByVal Level As Int32, _
                 <MarshalAs(UnmanagedType.LPStruct)> ByVal lpJob As JOB_INFO_2, _
                 ByVal dwCommand As PrintJobControlCommands _
                ) As Boolean

    End Function
#End Region

#Region "EnumMonitors"
    <DllImport("winspool.drv", EntryPoint:="EnumMonitors", _
 SetLastError:=True, CharSet:=CharSet.Unicode, _
 ExactSpelling:=False, _
 CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumMonitors(<InAttribute()> ByVal ServerName As String, _
                              <InAttribute()> ByVal Level As Int32, _
                              <OutAttribute()> ByVal lpBuf As Int32, _
                              <InAttribute()> ByVal cbBuf As Int32, _
                              <OutAttribute()> ByRef pcbNeeded As Int32, _
                              <OutAttribute()> ByRef pcReturned As Int32) As Boolean


    End Function

    <DllImport("winspool.drv", EntryPoint:="EnumMonitors", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumMonitors(<InAttribute()> ByVal pServerName As Int32, _
                          <InAttribute()> ByVal Level As Int32, _
                          <OutAttribute()> ByVal lpBuf As Int32, _
                          <InAttribute()> ByVal cbBuf As Int32, _
                          <OutAttribute()> ByRef pcbNeeded As Int32, _
                          <OutAttribute()> ByRef pcReturned As Int32) As Boolean


    End Function
#End Region

#Region "DocumentProperties"
    <DllImport("winspool.drv", EntryPoint:="DocumentProperties", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function DocumentProperties _
        (<InAttribute()> ByVal hwnd As Int32, _
         <InAttribute()> ByVal hPrinter As IntPtr, _
         <InAttribute()> ByVal pPrinterName As String, _
         <OutAttribute()> ByRef pDevModeOut As Int32, _
         <InAttribute()> ByVal pDevModeIn As Int32, _
         <InAttribute()> ByVal Mode As DocumentPropertiesModes) As Int32

    End Function
#End Region

#Region "DeviceCapabilities"
    <DllImport("winspool.drv", EntryPoint:="DeviceCapabilities", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function DeviceCapabilities(<InAttribute()> ByVal pPrinterName As String, _
                                        <InAttribute()> ByVal pPortName As String, _
                                        <InAttribute(), MarshalAs(UnmanagedType.U4)> ByVal CapbilityIndex As PrintDeviceCapabilitiesIndexes, _
                                        <OutAttribute()> ByRef lpOut As Int32, _
                                        <InAttribute()> ByVal pDevMode As Int32) As Int32

    End Function

#End Region

#Region "GetPrinterDriverDirectory"
    <DllImport("winspool.drv", EntryPoint:="GetPrinterDriverDirectory", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetPrinterDriverDirectory( _
                     <InAttribute()> ByVal ServerName As String, _
                     <InAttribute()> ByVal Environment As String, _
                     <InAttribute()> ByVal Level As Int32, _
                     <OutAttribute()> ByRef DriverDirectory As String, _
                     <InAttribute()> ByVal BufferSize As Int32, _
                     <OutAttribute()> ByRef BytesNeeded As Int32) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="GetPrinterDriverDirectory", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetPrinterDriverDirectory( _
                 <InAttribute()> ByVal ServerName As Int32, _
                 <InAttribute()> ByVal Environment As Int32, _
                 <InAttribute()> ByVal Level As Int32, _
                 <OutAttribute()> ByRef DriverDirectory As String, _
                 <InAttribute()> ByVal BufferSize As Int32, _
                 <OutAttribute()> ByRef BytesNeeded As Int32) As Boolean

    End Function
#End Region

#Region "AddPrinterDriver"
    <DllImport("winspool.drv", EntryPoint:="AddPrinterDriver", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function AddPrinterDriver( _
                       <InAttribute()> ByVal ServerName As String, _
                       <InAttribute()> ByVal Level As Int32, _
                       <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal pDriverInfo As DRIVER_INFO_2) As Boolean

    End Function

#End Region

#Region "EnumPorts"
    <DllImport("winspool.drv", EntryPoint:="EnumPorts", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPorts( _
                       <InAttribute()> ByVal ServerName As String, _
                       <InAttribute()> ByVal Level As Int32, _
                       <OutAttribute()> ByVal pbOut As Int32, _
                       <InAttribute()> ByVal cbIn As Int32, _
                       <OutAttribute()> ByRef pcbNeeded As Int32, _
                       <OutAttribute()> ByRef pcReturned As Int32 _
                ) As Boolean

    End Function

    <DllImport("winspool.drv", EntryPoint:="EnumPorts", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPorts( _
                       <InAttribute()> ByVal ServerName As Int32, _
                       <InAttribute()> ByVal Level As Int32, _
                       <OutAttribute()> ByVal pbOut As Int32, _
                       <InAttribute()> ByVal cbIn As Int32, _
                       <OutAttribute()> ByRef pcbNeeded As Int32, _
                       <OutAttribute()> ByRef pcReturned As Int32 _
                ) As Boolean

    End Function
#End Region

#Region "SetPort"
    <DllImport("winspool.drv", EntryPoint:="SetPort", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetPort( _
      <InAttribute()> ByVal ServerName As String, _
      <InAttribute()> ByVal PortName As String, _
      <InAttribute()> ByVal Level As Long, _
      <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal PortInfo As PORT_INFO_3) As Boolean


    End Function

#End Region

#Region "EnumPrintProcessors"
    <DllImport("winspool.drv", EntryPoint:="EnumPrintProcessors", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPrintProcessors(<InAttribute()> ByVal ServerName As String, _
                                       <InAttribute()> ByVal Environment As String, _
                                       <InAttribute()> ByVal Level As Int32, _
                                       <OutAttribute()> ByVal lpBuf As Int32, _
                                       <InAttribute()> ByVal cbBuf As Int32, _
                                       <OutAttribute()> ByRef pcbNeeded As Int32, _
                                       <OutAttribute()> ByRef pcbReturned As Int32) As Boolean


    End Function

#End Region

#Region "EnumPrintProcessorDataTypes"
    <DllImport("winspool.drv", EntryPoint:="EnumPrintProcessorDatatypes", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumPrinterProcessorDataTypes(<InAttribute()> ByVal ServerName As String, _
                                            <InAttribute()> ByVal PrintProcessorName As String, _
                                            <InAttribute()> ByVal Level As Int32, _
                                            <OutAttribute()> ByVal pDataTypes As Int32, _
                                            <InAttribute()> ByVal cbBuf As Int32, _
                                            <OutAttribute()> ByRef pcbNeeded As Int32, _
                                            <OutAttribute()> ByRef pcReturned As Int32) As Boolean

    End Function

#End Region

#Region "GetForm"
    <DllImport("winspool.drv", EntryPoint:="GetForm", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetForm(<InAttribute()> ByVal PrinterHandle As Int32, _
                             <InAttribute()> ByVal FormName As String, _
                             <InAttribute()> ByVal Level As Integer, _
                             <OutAttribute()> ByRef pForm As FORM_INFO_1, _
                             <InAttribute()> ByVal cbBuf As Int32, _
                             <OutAttribute()> ByRef pcbNeeded As Int32 _
                          ) As Boolean


    End Function
#End Region

#Region "SetForm"
    <DllImport("winspool.drv", EntryPoint:="SetForm", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetForm(<InAttribute()> ByVal PrinterHandle As IntPtr, _
                             <InAttribute()> ByVal FormName As String, _
                             <InAttribute()> ByVal Level As Integer, _
                             <InAttribute()> ByRef pForm As FORM_INFO_1 _
                          ) As Boolean

    End Function
#End Region

#Region "EnumForms"
    <DllImport("winspool.drv", EntryPoint:="EnumForms", _
SetLastError:=True, CharSet:=CharSet.Unicode, _
ExactSpelling:=False, _
CallingConvention:=CallingConvention.StdCall)> _
    Public Function EnumForms( _
                      <InAttribute()> ByVal hPrinter As IntPtr, _
                      <InAttribute()> ByVal Level As Int32, _
                      <OutAttribute()> ByVal pForm As Int32, _
                      <InAttribute()> ByVal cbBuf As Int32, _
                      <OutAttribute()> ByRef pcbNeeded As Int32, _
                      <OutAttribute()> ByRef pcFormsReturned As Int32) As Boolean

    End Function
#End Region

#Region "ReadPrinter"
    <DllImport("winspool.drv", EntryPoint:="ReadPrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function ReadPrinter(<InAttribute()> ByVal hPrinter As IntPtr, _
                                 <OutAttribute()> ByVal pBuffer As IntPtr, _
                                 <InAttribute()> ByVal cbBuf As Int32, _
                                 <OutAttribute()> ByRef pcbNeeded As Int32) As Boolean

    End Function

#End Region

#Region "WritePrinter"
    <DllImport("winspool.drv", EntryPoint:="WritePrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function WritePrinter(<InAttribute()> ByVal hPrinter As IntPtr, _
                                 <OutAttribute()> ByVal pBuffer As IntPtr, _
                                 <InAttribute()> ByVal cbBuf As Int32, _
                                 <OutAttribute()> ByRef pcbNeeded As Int32) As Boolean

    End Function
#End Region

#Region "StartDocPrinter"
    <DllImport("winspool.drv", EntryPoint:="StartDocPrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function StartDocPrinter(<InAttribute()> ByVal hPrinter As IntPtr, _
                                    <InAttribute()> ByVal Level As Int32, _
                                    <InAttribute(), MarshalAs(UnmanagedType.LPStruct)> ByVal DocInfo As DOC_INFO_1) As Boolean

    End Function
#End Region

#Region "EndDocPrinter"
    <DllImport("winspool.drv", EntryPoint:="EndDocPrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function EndDocPrinter(<InAttribute()> ByVal hPrinter As IntPtr) As Boolean

    End Function
#End Region

#Region "StartPagePrinter"
    <DllImport("winspool.drv", EntryPoint:="StartPagePrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function StartPagePrinter(<InAttribute()> ByVal hPrinter As IntPtr) As Boolean

    End Function
#End Region

#Region "EndPagePrinter"
    <DllImport("winspool.drv", EntryPoint:="EndPagePrinter", _
     SetLastError:=True, CharSet:=CharSet.Unicode, _
     ExactSpelling:=False, _
     CallingConvention:=CallingConvention.StdCall)> _
    Public Function EndPagePrinter(<InAttribute()> ByVal hPrinter As IntPtr) As Boolean

    End Function
#End Region

#End Region

End Module