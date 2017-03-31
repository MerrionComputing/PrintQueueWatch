# JobWritten event

This event is raised when data is written to the spool file (by an application) _or_ when data is written to the printer from the spool file.

The JobWritten event will be raised 1 or more times for every print job - usually many times.  When the PrintJob spooled property is true, all of the print job has been written to the spool file.  When the PrintJob printed property is true, all of the print job has been written to the printer.