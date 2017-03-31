# PrintDomain class

## Example code (VB.NET)
Lists all the print domains on all the printer providors visible from this process
{{
       Dim Providors As New PrintProvidorCollection

        Me.ListBox1.Items.Clear()
        For ps As Integer = 0 To Providors.Count - 1
            Me.ListBox1.Items.Add( Providors(ps).Name )
            For ds As Integer = 0 To  Providors(ps).PrintDomains.Count - 1
                Me.ListBox1.Items.Add( Providors(ps).PrintDomains(ds).Name )
            Next
        Next
}}

## Properties
* **Name** (String) - The name of the print domain
* **Description** (String) - The description of the print domain
* **Comment** (String) - The comment associated with this print domain

## Underlying API Declarations
{{
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi), System.Security.SuppressUnmanagedCodeSecurity()> _
    Friend Class PRINTER_INFO_1
        <MarshalAs(UnmanagedType.U4)> Public Flags As Int32
        <MarshalAs(UnmanagedType.LPStr)> Public pDescription As String
        <MarshalAs(UnmanagedType.LPStr)> Public pName As String
        <MarshalAs(UnmanagedType.LPStr)> Public pComment As String

        Public Sub New(ByVal hPrinter As IntPtr)

            Dim BytesWritten As Int32 = 0
            Dim ptBuf As IntPtr

            ptBuf = Marshal.AllocHGlobal(1)

            If Not GetPrinter(hPrinter, 1, ptBuf, 1, BytesWritten) Then
                If BytesWritten > 0 Then
                    '\\ Free the buffer allocated
                    Marshal.FreeHGlobal(ptBuf)
                    ptBuf = Marshal.AllocHGlobal(BytesWritten)
                    If GetPrinter(hPrinter, 1, ptBuf, BytesWritten, BytesWritten) Then
                        Marshal.PtrToStructure(ptBuf, Me)
                    Else
                        Throw New Win32Exception()
                    End If
                    '\\ Free this buffer again
                    Marshal.FreeHGlobal(ptBuf)
                Else
                    Throw New Win32Exception()
               End If
            Else
                Throw New Win32Exception()
            End If

        End Sub

        Public Sub New()

        End Sub

    End Class
}}
