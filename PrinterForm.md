# PrinterForm class

Represents a form (paper size) that has been installed on a printer

## Code example (VB.NET)
List the print forms on the named printer
{{
        Dim pi As New PrinterInformation("Microsoft Office Document Image Writer", SpoolerApiConstantEnumerations.PrinterAccessRights.PRINTER_ALL_ACCESS, True)

        For pf As Integer = 0 to pi.PrinterForms.Count -1
            pi.PrinterForms(pf).Name
       Next
}}

## Properties
* **Name** (String)- The name of the printer form 
* **Width** (Integer) - The width of the form measured in millimeters
* **Height** (Integer)- The height of the printer form measured in millimeters

## Underlying API Code
{{
    <StructLayout(LayoutKind.Sequential)> _
    Friend Class FORM_INFO_1
        <MarshalAs(UnmanagedType.U4)> Public Flags As Int32 'FormTypeFlags
        <MarshalAs(UnmanagedType.LPStr)> Public Name As String
        <MarshalAs(UnmanagedType.U4)> Public Width As Int32
        <MarshalAs(UnmanagedType.U4)> Public Height As Int32
        <MarshalAs(UnmanagedType.U4)> Public Left As Int32
        <MarshalAs(UnmanagedType.U4)> Public Top As Int32
        <MarshalAs(UnmanagedType.U4)> Public Right As Int32
        <MarshalAs(UnmanagedType.U4)> Public Bottom As Int32

    End Class
}}