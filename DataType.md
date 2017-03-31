# DataType class

The data type of a spool file or print job.
This is the data type that the spool file contains.  It can be EMF or RAW.  
(EMF means a special printer based superset of the windows metafile format, as documented in [http://www.codeproject.com/KB/printing/EMFSpoolViewer.aspx](http://www.codeproject.com/KB/printing/EMFSpoolViewer.aspx), 
RAW means one of the printer control languages such as PostScript, PCL5, PCL6 etc.)

## Code example (VB.NET)
List all the data types for each [PrintProcessor](PrintProcessor) installed on the current server
{{
        Dim server As New PrintServer
        Dim Processor As PrintProcessor
        Me.ListBox1.Items.Clear()
        For ps As Integer = 0 To server.PrintProcessors.Count - 1
            ListBox1.Items.Add( server.PrintProcessors(ps).Name )
            For dt As Integer = 0 to server.PrintProcessors(ps).DataTypes.Count - 1
                Me.ListBox1.Items.Add(server.PrintProcessors(ps).DataTypes(dt).Name)
            Next
        Next
}}

