'\\ --[SpoolerApiConstantEnumerations]-------------------------------------------------------
'\\ Namespace for all the spooler related enumerated types, from winspool.drv
'\\ (c) 2003 Merrion Computing Ltd
'\\ http://www.merrioncomputing.com
'\\ ------------------------------------------------------------------------------------------
Namespace SpoolerApiConstantEnumerations

#Region "PrinterChangeNotificationGeneralFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationGeneralFlags
        PRINTER_CHANGE_JOB = &HFF00
        PRINTER_CHANGE_PRINTER = &HFF
        PRINTER_CHANGE_FORM = &H70000
        PRINTER_CHANGE_PORT = &H700000
        PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
        PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
    End Enum
#End Region

#Region "PrinterChangeNotificationFormFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationFormFlags
        PRINTER_CHANGE_NO_FORM_CHANGE = &H0
        PRINTER_CHANGE_ADD_FORM = &H10000
        PRINTER_CHANGE_SET_FORM = &H20000
        PRINTER_CHANGE_DELETE_FORM = &H40000
    End Enum
#End Region

#Region "PrinterChangeNotificationPortFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationPortFlags
        PRINTER_CHANGE_NO_PORT_CHANGE = &H0
        PRINTER_CHANGE_ADD_PORT = &H100000
        PRINTER_CHANGE_CONFIGURE_PORT = &H200000
        PRINTER_CHANGE_DELETE_PORT = &H400000
    End Enum
#End Region

#Region "PrinterChangeNotificationJobFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationJobFlags
        PRINTER_CHANGE_NO_JOB_CHANGE = &H0
        PRINTER_CHANGE_ADD_JOB = &H100
        PRINTER_CHANGE_SET_JOB = &H200
        PRINTER_CHANGE_DELETE_JOB = &H400
        PRINTER_CHANGE_WRITE_JOB = &H800
    End Enum
#End Region

#Region "PrinterChangeNotificationPrinterFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationPrinterFlags
        PRINTER_CHANGE_NO_PRINTER_CHANGE = &H0
        PRINTER_CHANGE_ADD_PRINTER = &H1
        PRINTER_CHANGE_SET_PRINTER = &H2
        PRINTER_CHANGE_DELETE_PRINTER = &H4
        PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
    End Enum
#End Region

#Region "PrinterChangeNotificationProcessorFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationProcessorFlags
        PRINTER_CHANGE_NO_PRINT_PROCESSOR_CHANGE = &H0
        PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
        PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
    End Enum
#End Region

#Region "PrinterChangeNotificationDriverFlags"
    <FlagsAttribute()> _
    Public Enum PrinterChangeNotificationDriverFlags
        PRINTER_CHANGE_NO_DRIVER_CHANGE = &H0
        PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
        PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
        PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
    End Enum
#End Region

#Region "PrintJobStatuses"
    <FlagsAttribute()> _
    Public Enum PrintJobStatuses
        ' JOB_STATUS_PAUSED The print job has been paused during printing
        JOB_STATUS_PAUSED = &H1
        ' JOB_STATUS_ERROR The print job is stalled by an error
        JOB_STATUS_ERROR = &H2
        ' JOB_STATUS_DELETING The print job is being deleted from the queue
        JOB_STATUS_DELETING = &H4
        ' JOB_STATUS_SPOOLING The print job is being added to the spooler queue
        JOB_STATUS_SPOOLING = &H8
        ' JOB_STATUS_PRINTING The print job is in the midst of being printed
        JOB_STATUS_PRINTING = &H10
        ' JOB_STATUS_OFFLINE The print job is stalled because the printer is off line
        JOB_STATUS_OFFLINE = &H20
        ' JOB_STATUS_PAPEROUT The print job is stalled because the printer has run out of paper
        JOB_STATUS_PAPEROUT = &H40
        ' JOB_STATUS_PRINTER The print job is waiting on the printer
        JOB_STATUS_PRINTED = &H80
        ' JOB_STATUS_DELETED The print job has been deleted
        JOB_STATUS_DELETED = &H100
        ' JOB_STATUS_BLOCKED_DEVICEQUEUE The device queue is blocked so the print job is also blocked
        JOB_STATUS_BLOCKED_DEVICEQUEUE = &H200
        ' JOB_STATUS_USER_INTERVENTION The queued job is paused by user intervention
        JOB_STATUS_USER_INTERVENTION = &H400
        ' JOB_STATUS_RESTART The queued print job is restarting after being paused
        JOB_STATUS_RESTART = &H800
    End Enum
#End Region

#Region "DeviceOrientations"
    ' DeviceOrientations describe the orientation of the device
    Public Enum DeviceOrientations
        ' DMORIENT_PORTRAIT The device is in portait (short edge on top) mode
        DMORIENT_PORTRAIT = 1
        ' DMORIENT_LANDSCAPE The device is in landscape (long edge on top) mode
        DMORIENT_LANDSCAPE = 2
    End Enum
#End Region

#Region "PrintJobControlCommands"
    Public Enum PrintJobControlCommands
        ' JOB_CONTROL_SETJOB Set the job info
        JOB_CONTROL_SETJOB = 0
        ' JOB_CONTROL_PAUSE Pause printing of the current job
        JOB_CONTROL_PAUSE = 1
        ' JOB_CONTROL_RESUME Resume printing a paused job
        JOB_CONTROL_RESUME = 2
        ' JOB_CONTROL_CANCEL Cancel a queued or printing job
        JOB_CONTROL_CANCEL = 3
        ' JOB_CONTROL_RESTART Restart a paused job that was printing
        JOB_CONTROL_RESTART = 4
        ' JOB_CONTROL_DELETE Remove a job from the queue
        JOB_CONTROL_DELETE = 5
        ' JOB_CONTROL_SENT_TO_PRINTER End a print job
        JOB_CONTROL_SENT_TO_PRINTER = 6
        ' JOB_CONTROL_LAST_PAGE_EJECTED End a print job
        JOB_CONTROL_LAST_PAGE_EJECTED = 7
    End Enum
#End Region

#Region "DeviceModeResolutions"
    ' DeviceModeResolutions are the different resolution settings for devices that _
    'support multiple resolutions
    Public Enum DeviceModeResolutions
        ' DMRES_DRAFT Device is set to draft quality
        DMRES_DRAFT = -1
        ' DMRES_LOW Device is set to low resolution quality
        DMRES_LOW = -2
        ' DMRES_MEDIUM Device is set to an intermediate resolution quality
        DMRES_MEDIUM = -3
        ' DMRES_HIGH Devide is set to high resolution quality
        DMRES_HIGH = -4
    End Enum
#End Region

#Region "DeviceColourModes"
    ' DeviceColourModes are the possible colour use settings for colour devices
    Public Enum DeviceColourModes
        ' DMCOLOR_MONOCHROME Do not use colour
        DMCOLOR_MONOCHROME = 1
        ' DMCOLOR_COLOR Do use colour
        DMCOLOR_COLOR = 2
    End Enum
#End Region

#Region "DeviceDuplexSettings"
    ' DeviceDuplexSettings is how a duplex device is set for duplex printing
    Public Enum DeviceDuplexSettings
        ' DMDUP_SIMPLEX Device is not using double sided printing
        DMDUP_SIMPLEX = 1
        ' DMDUP_VERTICAL Device is performing double sided printing on the short edge
        DMDUP_VERTICAL = 2
        ' DMDUP_HORIZONTAL Device is performing double sided printing on the long edge
        DMDUP_HORIZONTAL = 3
    End Enum
#End Region

#Region "DeviceTrueTypeFontSettings"
    ' DeviceTrueTypeFontSettings specifies how the device handles TrueType fonts
    Public Enum DeviceTrueTypeFontSettings
        ' DMTT_BITMAP Device treats TrueType fonts as bitmaps
        DMTT_BITMAP = 1
        ' DMTT_DOWNLOAD Device downloads TrueType fonts as soft fonts
        DMTT_DOWNLOAD = 2
        ' DMTT_SUBDEV Device substitutes TrueType fonts for the nearest printer font
        DMTT_SUBDEV = 3
        ' DMTT_DOWNLOAD_OUTLINE Device downloads TrueType fonts as outline soft fonts
        DMTT_DOWNLOAD_OUTLINE = 4
    End Enum
#End Region

#Region "DeviceCollateSettings"
    ' DeviceCollateSettings set whether or not multi-copy print jobs should be collated
    Public Enum DeviceCollateSettings
        ' DMCOLLATE_FALSE Do not collate the jobs
        DMCOLLATE_FALSE = 0
        ' DMCOLLATE_TRUE Do collate the jobs
        DMCOLLATE_TRUE = 1
    End Enum
#End Region

#Region "EnumPrinterFlags"
    <Flags()> _
    Public Enum EnumPrinterFlags
        PRINTER_ENUM_DEFAULT = &H1
        PRINTER_ENUM_LOCAL = &H2
        PRINTER_ENUM_CONNECTIONS = &H4
        PRINTER_ENUM_FAVORITE = &H4
        PRINTER_ENUM_NAME = &H8
        PRINTER_ENUM_REMOTE = &H10
        PRINTER_ENUM_SHARED = &H20
        PRINTER_ENUM_NETWORK = &H40
    End Enum
#End Region

#Region "JobInfoLevels"
    ' JobInfoLevels - The level (JOB_LEVEL_n) structure to read from the spooler
    Public Enum JobInfoLevels
        ' JobInfoLevel1 - Read a JOB_INFO_1 structure
        JobInfoLevel1 = &H1
        ' JobInfoLevel2 - Read a JOB_INFO_2 structure
        JobInfoLevel2 = &H2
        ' JobInfoLevel3 - Read a JOB_INFO_3 structure
        JobInfoLevel3 = &H3
    End Enum
#End Region

#Region "PrinterAccessRights"
    <FlagsAttribute()> _
    Public Enum PrinterAccessRights
        ' READ_CONTROL - Allowed to read printer information
        READ_CONTROL = &H20000
        ' WRITE_DAC - Allowed to write device access control info
        WRITE_DAC = &H40000
        ' WRITE_OWNER - Allowed to change the object owner
        WRITE_OWNER = &H80000
        ' SERVER_ACCESS_ADMINISTER 
        SERVER_ACCESS_ADMINISTER = &H1
        '  SERVER_ACCESS_ENUMERATE
        SERVER_ACCESS_ENUMERATE = &H2
        ' PRINTER_ACCESS_ADMINISTER Allows administration of a printer
        PRINTER_ACCESS_ADMINISTER = &H4
        ' PRINTER_ACCESS_USE Allows printer general use (printing, querying)
        PRINTER_ACCESS_USE = &H8
        ' PRINTER_ALL_ACCESS Allows use and administration.
        PRINTER_ALL_ACCESS = &HF000C
        ' SERVER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | SERVER_ACCESS_ADMINISTER | SERVER_ACCESS_ENUMERATE)
        SERVER_ALL_ACCESS = &HF0003
    End Enum
#End Region

#Region "PrinterControlCommands"
    Public Enum PrinterControlCommands
        ' PRINTER_CONTROL_SETPRINTERINFO - Update the printer info data
        PRINTER_CONTROL_SETPRINTERINFO = 0
        ' PRINTER_CONTROL_PAUSE Pause the printing of the currently active job
        PRINTER_CONTROL_PAUSE = 1
        ' PRINTER_CONTROL_RESUME Resume printing if paused
        PRINTER_CONTROL_RESUME = 2
        ' PRINTER_CONTROL_PURGE Terminate and delete the currently printing job
        PRINTER_CONTROL_PURGE = 3
        ' PRINTER_CONTROL_SET_STATUS Set the printer status for the current job
        PRINTER_CONTROL_SET_STATUS = 4
    End Enum
#End Region

#Region "PrinterStatuses"
    ' PrinterStatuses are the various possible states a printer can be in. Note that this information is provided by the printer driver and in many cases the status will always be "Ready" unless there is one or more jobs in error on that printer.
    <FlagsAttribute()> _
    Public Enum PrinterStatuses
        ' PRINTER_STATUS_READY The printer is free and ready to print
        PRINTER_STATUS_READY = &H0
        ' PRINTER_STATUS_PAUSED Printing is paused
        PRINTER_STATUS_PAUSED = &H1
        ' PRINTER_STATUS_ERROR Printing is suspended due to a general error
        PRINTER_STATUS_ERROR = &H2
        ' PRINTER_STATUS_PENDING_DELETION Printing is suspended while a job is being deleted
        PRINTER_STATUS_PENDING_DELETION = &H4
        ' PRINTER_STATUS_PAPER_JAM Printing is suspended because there has been a paper jam
        PRINTER_STATUS_PAPER_JAM = &H8
        ' PRINTER_STATUS_PAPER_OUT Printing is suspended because the printer has run out of paper
        PRINTER_STATUS_PAPER_OUT = &H10
        ' PRINTER_STATUS_MANUAL_FEED Printing is suspended pending a manual paper feed
        PRINTER_STATUS_MANUAL_FEED = &H20
        ' PRINTER_STATUS_PAPER_PROBLEM Printing is supended due to a paper problem
        PRINTER_STATUS_PAPER_PROBLEM = &H40
        ' PRINTER_STATUS_OFFLINE The printer is off line
        PRINTER_STATUS_OFFLINE = &H80
        ' PRINTER_STATUS_IO_ACTIVE The printer is reading data from the IO port
        PRINTER_STATUS_IO_ACTIVE = &H100
        ' PRINTER_STATUS_BUSY The printer is busy
        PRINTER_STATUS_BUSY = &H200
        ' PRINTER_STATUS_PRINTING The printer is active and printing a job
        PRINTER_STATUS_PRINTING = &H400
        ' PRINTER_STATUS_OUTPUT_BIN_FULL The output tray is full and the printer has paused until it is emptied
        PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
        ' PRINTER_STATUS_NOT_AVAILABLE The printer status is not known
        PRINTER_STATUS_NOT_AVAILABLE = &H1000
        ' PRINTER_STATUS_WAITING The printer is waiting for a job
        PRINTER_STATUS_WAITING = &H2000
        ' PRINTER_STATUS_PROCESSING The printer is active and processing a print job
        PRINTER_STATUS_PROCESSING = &H4000
        ' PRINTER_STATUS_INITIALIZING The printer is initialising and not yet ready to print
        PRINTER_STATUS_INITIALIZING = &H8000
        ' PRINTER_STATUS_WARMING_UP The printer is warming up and not yet ready to print
        PRINTER_STATUS_WARMING_UP = &H10000
        ' PRINTER_STATUS_TONER_LOW Printing is suspended because the level of toner or ink is too low for a reasonable quality print
        PRINTER_STATUS_TONER_LOW = &H20000
        ' PRINTER_STATUS_NO_TONER Printing has been suspended because the printer is out of toner or ink
        PRINTER_STATUS_NO_TONER = &H40000
        ' PRINTER_STATUS_PAGE_PUNT (Win95 only) The printer is suspended while a page is deleted
        PRINTER_STATUS_PAGE_PUNT = &H80000 'win95
        ' PRINTER_STATUS_USER_INTERVENTION The printer has been paused by the user intervention
        PRINTER_STATUS_USER_INTERVENTION = &H100000
        ' PRINTER_STATUS_OUT_OF_MEMORY Printing is suspended because the printer has run out of memory
        PRINTER_STATUS_OUT_OF_MEMORY = &H200000
        ' PRINTER_STATUS_DOOR_OPEN Printing is supended because a door on the printer is open
        PRINTER_STATUS_DOOR_OPEN = &H400000
        ' PRINTER_STATUS_SERVER_UNKNOWN Printing is suspended due to an unknown server error
        PRINTER_STATUS_SERVER_UNKNOWN = &H800000
        ' PRINTER_STATUS_POWER_SAVE The printer is suspended in power saving mode
        PRINTER_STATUS_POWER_SAVE = &H1000000
    End Enum
#End Region

#Region "PrinterAttributes"
    <FlagsAttribute()> _
    Public Enum PrinterAttributes
        ' PRINTER_ATTRIBUTE_QUEUED
        PRINTER_ATTRIBUTE_QUEUED = &H1
        ' PRINTER_ATTRIBUTE_DIRECT
        PRINTER_ATTRIBUTE_DIRECT = &H2
        ' PRINTER_ATTRIBUTE_DEFAULT
        PRINTER_ATTRIBUTE_DEFAULT = &H4
        ' PRINTER_ATTRIBUTE_SHARED
        PRINTER_ATTRIBUTE_SHARED = &H8
        ' PRINTER_ATTRIBUTE_NETWORK
        PRINTER_ATTRIBUTE_NETWORK = &H10
        ' PRINTER_ATTRIBUTE_HIDDEN
        PRINTER_ATTRIBUTE_HIDDEN = &H20
        ' PRINTER_ATTRIBUTE_LOCAL
        PRINTER_ATTRIBUTE_LOCAL = &H40
        ' PRINTER_ATTRIBUTE_ENABLE_DEVQ
        PRINTER_ATTRIBUTE_ENABLE_DEVQ = &H80
        ' PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS
        PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS = &H100
        ' PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST
        PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST = &H200
        ' PRINTER_ATTRIBUTE_WORK_OFFLINE
        PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
        ' PRINTER_ATTRIBUTE_ENABLE_BIDI The printer can operate in bidirectional mode
        PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800
        ' PRINTER_ATTRIBUTE_RAW_ONLY The printer can only accept raw data
        PRINTER_ATTRIBUTE_RAW_ONLY = &H1000
        ' PRINTER_ATTRIBUTE_PUBLISHED The printer is shared as a network resource
        PRINTER_ATTRIBUTE_PUBLISHED = &H2000
    End Enum
#End Region

#Region "PrinterInfoLevels"
    ' PrinterInfoLevels - The level (PRINTER_INFO_n) structure to read from the spooler
    Public Enum PrinterInfoLevels
        ' PrinterInfoLevel1 - Read a PRINT_INFO_1 structure
        PrinterInfoLevel1 = 1
        ' PrinterInfoLevel2 - Read a PRINT_INFO_2 structure
        PrinterInfoLevel2 = 2
        ' PrinterInfoLevel3 - Read a PRINT_INFO_3 structure (Windows NT/2000/XP only)
        PrinterInfoLevel3 = 3
        ' PrinterInfoLevel4 - Read a PRINT_INFO_4 structure (Windows NT/2000/XP only)
        PrinterInfoLevel4 = 4
        ' PrinterInfoLevel5 - Read a PRINT_INFO_5 structure
        PrinterInfoLevel5 = 5
        ' PrinterInfoLevel6 - Read a PRINT_INFO_6 structure
        PrinterInfoLevel6 = 6
        ' PrinterInfoLevel7 - Read a PRINT_INFO_7 structure (Windows 2000/XP only)
        PrinterInfoLevel7 = 7
        ' PrinterInfoLevel8 - Read a PRINT_INFO_8 structure (Windows 2000/XP only)
        PrinterInfoLevel8 = 8
        ' PrinterInfoLevel9 - Read a PRINT_INFO_9 structure (Windows 2000/XP only)
        PrinterInfoLevel9 = 9
    End Enum
#End Region

#Region "Printer_Notification_Types"
    Public Enum Printer_Notification_Types
        PRINTER_NOTIFY_TYPE = &H0
        JOB_NOTIFY_TYPE = &H1
    End Enum
#End Region

#Region "Printer_Notify_Field_Indexes"
    Public Enum Printer_Notify_Field_Indexes As Short
        PRINTER_NOTIFY_FIELD_SERVER_NAME = &H0
        PRINTER_NOTIFY_FIELD_PRINTER_NAME = &H1
        PRINTER_NOTIFY_FIELD_SHARE_NAME = &H2
        PRINTER_NOTIFY_FIELD_PORT_NAME = &H3
        PRINTER_NOTIFY_FIELD_DRIVER_NAME = &H4
        PRINTER_NOTIFY_FIELD_COMMENT = &H5
        PRINTER_NOTIFY_FIELD_LOCATION = &H6
        PRINTER_NOTIFY_FIELD_DEVMODE = &H7
        PRINTER_NOTIFY_FIELD_SEPFILE = &H8
        PRINTER_NOTIFY_FIELD_PRINT_PROCESSOR = &H9
        PRINTER_NOTIFY_FIELD_PARAMETERS = &HA
        PRINTER_NOTIFY_FIELD_DATATYPE = &HB
        PRINTER_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
        PRINTER_NOTIFY_FIELD_ATTRIBUTES = &HD
        PRINTER_NOTIFY_FIELD_PRIORITY = &HE
        PRINTER_NOTIFY_FIELD_DEFAULT_PRIORITY = &HF
        PRINTER_NOTIFY_FIELD_START_TIME = &H10
        PRINTER_NOTIFY_FIELD_UNTIL_TIME = &H11
        PRINTER_NOTIFY_FIELD_STATUS = &H12
        PRINTER_NOTIFY_FIELD_STATUS_STRING = &H13
        PRINTER_NOTIFY_FIELD_CJOBS = &H14
        PRINTER_NOTIFY_FIELD_AVERAGE_PPM = &H15
        PRINTER_NOTIFY_FIELD_TOTAL_PAGES = &H16
        PRINTER_NOTIFY_FIELD_PAGES_PRINTED = &H17
        PRINTER_NOTIFY_FIELD_TOTAL_BYTES = &H18
        PRINTER_NOTIFY_FIELD_BYTES_PRINTED = &H19
        PRINTER_NOTIFY_FIELD_OBJECT_GUID = &H1A
    End Enum
#End Region

#Region "Job_Notify_Field_Indexes"
    Public Enum Job_Notify_Field_Indexes As Short
        JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
        JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
        JOB_NOTIFY_FIELD_PORT_NAME = &H2
        JOB_NOTIFY_FIELD_USER_NAME = &H3
        JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
        JOB_NOTIFY_FIELD_DATATYPE = &H5
        JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
        JOB_NOTIFY_FIELD_PARAMETERS = &H7
        JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
        JOB_NOTIFY_FIELD_DEVMODE = &H9
        JOB_NOTIFY_FIELD_STATUS = &HA
        JOB_NOTIFY_FIELD_STATUS_STRING = &HB
        JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
        JOB_NOTIFY_FIELD_DOCUMENT = &HD
        JOB_NOTIFY_FIELD_PRIORITY = &HE
        JOB_NOTIFY_FIELD_POSITION = &HF
        JOB_NOTIFY_FIELD_SUBMITTED = &H10
        JOB_NOTIFY_FIELD_START_TIME = &H11
        JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
        JOB_NOTIFY_FIELD_TIME = &H13
        JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
        JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
        JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
        JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
    End Enum
#End Region

#Region "PrinterPaperBins"
    Public Enum PrinterPaperBins
        ' DMBIN_UPPER The upper (or only) tray
        DMBIN_UPPER = 1
        ' DMBIN_LOWER The lower paper tray
        DMBIN_LOWER = 2
        ' DMBIN_MIDDLE The middle paper tray
        DMBIN_MIDDLE = 3
        ' DMBIN_MANUAL The manual form feed
        DMBIN_MANUAL = 4
        ' DMBIN_ENVELOPE The envelope tray
        DMBIN_ENVELOPE = 5
        ' DMBIN_ENVMANUAL Manual envelope feed
        DMBIN_ENVMANUAL = 6
        ' DMBIN_AUTO The printer automatically selects the source
        DMBIN_AUTO = 7
        ' DMBIN_TRACTOR The tractor feed paper source
        DMBIN_TRACTOR = 8
        ' DMBIN_SMALLFMT The smaller format paper tray
        DMBIN_SMALLFMT = 9
        ' DMBIN_LARGEFMT The larger format paper tray
        DMBIN_LARGEFMT = 10
        ' DMBIN_LARGECAPACITY The high capacity tray
        DMBIN_LARGECAPACITY = 11
        ' DMBIN_CASSETTE The paper cassete source
        DMBIN_CASSETTE = 14
        ' DMBIN_FORMSOURCE The preprinted from feed
        DMBIN_FORMSOURCE = 15
        ' DMBIN_USER Paper source is device specific
        DMBIN_USER = 256
    End Enum
#End Region


#Region "SpoolerWin32ErrorCodes"
    Public Enum SpoolerWin32ErrorCodes
        ERROR_ACCESS_DENIED = 5
    End Enum
#End Region

#Region "PrinterDriverOperatingSystemVersion"
    Public Enum PrinterDriverOperatingSystemVersion
        Driver_Win9x = 0
        Driver_NT351 = 1
        Driver_NT4 = 2
        Driver_2000_XP = 3
    End Enum
#End Region

#Region "DocumentPropertiesModes"
    <FlagsAttribute()> _
    Public Enum DocumentPropertiesModes
        DM_GETSIZE = 0
        DM_COPY = 2
        DM_PROMPT = 4
        DM_MODIFY = 8
    End Enum
#End Region

#Region "PrintDeviceCapabilitiesIndexes"
    Public Enum PrintDeviceCapabilitiesIndexes
        '  DC_FIELDS Which fields of the device mode are used
        DC_FIELDS = 1
        '  DC_PAPERS Which Printer Paper Sizes the device supports
        DC_PAPERS = 2
        '  DC_PAPERSIZE The dimensions of the paper in 10ths of a millimeter
        DC_PAPERSIZE = 3
        '  DC_MINEXTENT The minimum paper width and height the printer can support
        DC_MINEXTENT = 4
        '  DC_MAXEXTENT The maximum paper width and height the printer can support
        DC_MAXEXTENT = 5
        '  DC_BINS The standard paper bins supported by this printer
        DC_BINS = 6
        '  DC_DUPLEX Whether the printer supports duplex printing
        DC_DUPLEX = 7
        '  DC_SIZE
        DC_SIZE = 8
        '  DC_EXTRA
        DC_EXTRA = 9
        '  DC_VERSION
        DC_VERSION = 10
        '  DC_DRIVER
        DC_DRIVER = 11
        '  DC_BINNAMES
        DC_BINNAMES = 12
        '  DC_ENUMRESOLUTIONS
        DC_ENUMRESOLUTIONS = 13
        '  DC_FILEDEPENDENCIES
        DC_FILEDEPENDENCIES = 14
        '  DC_TRUETYPE
        DC_TRUETYPE = 15
        '  DC_PAPERNAMES
        DC_PAPERNAMES = 16
        '  DC_ORIENTATION
        DC_ORIENTATION = 17
        '  DC_COPIES
        DC_COPIES = 18
        '  DC_BINADJUST
        DC_BINADJUST = 19
        '  DC_EMF_COMPLIANT
        DC_EMF_COMPLIANT = 20
        '  DC_DATATYPE_PRODUCED
        DC_DATATYPE_PRODUCED = 21
        '  DC_COLLATE - Returns non zero if the device supports collation
        DC_COLLATE = 22
        '  DC_MANUFACTURER
        DC_MANUFACTURER = 23
        '  DC_MODEL
        DC_MODEL = 24
        '  DC_PERSONALITY
        DC_PERSONALITY = 25
        '  DC_PRINTRATE
        DC_PRINTRATE = 26
        '  DC_PRINTRATEUNIT
        DC_PRINTRATEUNIT = 27
        '  DC_PRINTERMEM
        DC_PRINTERMEM = 28
        '  DC_MEDIAREADY
        DC_MEDIAREADY = 29
        '  DC_STAPLE - Returns non zero if the device supports stapling
        DC_STAPLE = 30
        '  DC_PRINTRATEPPM
        DC_PRINTRATEPPM = 31
        '  DC_COLORDEVICE
        DC_COLORDEVICE = 32
        '  DC_NUP
        DC_NUP = 33
    End Enum
#End Region

#Region "PortTypes"
    <FlagsAttribute()> _
    Public Enum PortTypes
        PORT_TYPE_WRITE = &H1
        PORT_TYPE_READ = &H2
        PORT_TYPE_REDIRECTED = &H4
        PORT_TYPE_NET_ATTACHED = &H8
    End Enum
#End Region

#Region "PortStatuses"
    Public Enum PortStatuses
        PORT_STATUS_OK = 0
        PORT_STATUS_OFFLINE = 1
        PORT_STATUS_PAPER_JAM = 2
        PORT_STATUS_PAPER_OUT = 3
        PORT_STATUS_OUTPUT_BIN_FULL = 4
        PORT_STATUS_PAPER_PROBLEM = 5
        PORT_STATUS_NO_TONER = 6
        PORT_STATUS_DOOR_OPEN = 7
        PORT_STATUS_USER_INTERVENTION = 8
        PORT_STATUS_OUT_OF_MEMORY = 9
        PORT_STATUS_TONER_LOW = 10
        PORT_STATUS_WARMING_UP = 11
        PORT_STATUS_POWER_SAVE = 12
    End Enum
#End Region

#Region "PortStatusSeverity"
    Public Enum PortStatusSeverity
        PORT_STATUS_TYPE_ERROR = 1
        PORT_STATUS_TYPE_WARNING = 2
        PORT_STATUS_TYPE_INFO = 3
    End Enum
#End Region

#Region "FormTypeFlags"
    Public Enum FormTypeFlags
        FORM_USER = &H0
        FORM_BUILTIN = &H1
        FORM_PRINTER = &H2
    End Enum
#End Region

#Region "SpoolerRegisterKeys"
    '\\ --[SpoolerRegisterKeys]-------------------------------
    '\\ Predefined registry key strings used by GetPrinterData 
    '\\ and SetPrinterData API functions
    '\\ ------------------------------------------------------
    Public Class SpoolerRegisterKeys

        Public Enum SpoolerRegisterKeyIndexes
            SPLREG_DEFAULT_SPOOL_DIRECTORY = 1 '"DefaultSpoolDirectory"
            SPLREG_PORT_THREAD_PRIORITY_DEFAULT = 2      '"PortThreadPriorityDefault"
            SPLREG_PORT_THREAD_PRIORITY = 3              '"PortThreadPriority"
            SPLREG_SCHEDULER_THREAD_PRIORITY_DEFAULT = 4 '"SchedulerThreadPriorityDefault"
            SPLREG_SCHEDULER_THREAD_PRIORITY = 5         '"SchedulerThreadPriority"
            SPLREG_BEEP_ENABLED = 6 '"BeepEnabled"
            SPLREG_NET_POPUP = 7        '"NetPopup"
            SPLREG_RETRY_POPUP = 8 '"RetryPopup"
            SPLREG_NET_POPUP_TO_COMPUTER = 9 '"NetPopupToComputer"
            SPLREG_EVENT_LOG = 10 '"EventLog"
            SPLREG_MAJOR_VERSION = 11 '"MajorVersion"
            SPLREG_MINOR_VERSION = 12 '"MinorVersion"
            SPLREG_ARCHITECTURE = 13 '"Architecture"
            SPLREG_OS_VERSION = 14 '"OSVersion"
            SPLREG_OS_VERSIONEX = 15 '"OSVersionEx"
            SPLREG_DS_PRESENT = 16 '"DsPresent"
            SPLREG_DS_PRESENT_FOR_USER = 17 '"DsPresentForUser"
            SPLREG_REMOTE_FAX = 18          '"RemoteFax"
            SPLREG_RESTART_JOB_ON_POOL_ERROR = 19 '"RestartJobOnPoolError"
            SPLREG_RESTART_JOB_ON_POOL_ENABLED = 20 '"RestartJobOnPoolEnabled"
            SPLREG_DNS_MACHINE_NAME = 21  '"DNSMachineName"
        End Enum

        Public Shared Function PredefinedKeyName(ByVal KeyIndex As SpoolerRegisterKeyIndexes) As String
            Select Case KeyIndex
                Case SpoolerRegisterKeyIndexes.SPLREG_DEFAULT_SPOOL_DIRECTORY
                    Return "DefaultSpoolDirectory"
                Case SpoolerRegisterKeyIndexes.SPLREG_PORT_THREAD_PRIORITY_DEFAULT
                    Return "PortThreadPriorityDefault"
                Case SpoolerRegisterKeyIndexes.SPLREG_PORT_THREAD_PRIORITY
                    Return "PortThreadPriority"
                Case SpoolerRegisterKeyIndexes.SPLREG_SCHEDULER_THREAD_PRIORITY_DEFAULT
                    Return "SchedulerThreadPriorityDefault"
                Case SpoolerRegisterKeyIndexes.SPLREG_SCHEDULER_THREAD_PRIORITY
                    Return "SchedulerThreadPriority"
                Case SpoolerRegisterKeyIndexes.SPLREG_BEEP_ENABLED
                    Return "BeepEnabled"
                Case SpoolerRegisterKeyIndexes.SPLREG_NET_POPUP
                    Return "NetPopup"
                Case SpoolerRegisterKeyIndexes.SPLREG_RETRY_POPUP
                    Return "RetryPopup"
                Case SpoolerRegisterKeyIndexes.SPLREG_NET_POPUP_TO_COMPUTER
                    Return "NetPopupToComputer"
                Case SpoolerRegisterKeyIndexes.SPLREG_EVENT_LOG
                    Return "EventLog"
                Case SpoolerRegisterKeyIndexes.SPLREG_MAJOR_VERSION
                    Return "MajorVersion"
                Case SpoolerRegisterKeyIndexes.SPLREG_MINOR_VERSION
                    Return "MinorVersion"
                Case SpoolerRegisterKeyIndexes.SPLREG_ARCHITECTURE
                    Return "Architecture"
                Case SpoolerRegisterKeyIndexes.SPLREG_OS_VERSION
                    Return "OSVersion"
                Case SpoolerRegisterKeyIndexes.SPLREG_OS_VERSIONEX
                    Return "OSVersionEx"
                Case SpoolerRegisterKeyIndexes.SPLREG_DS_PRESENT
                    Return "DsPresent"
                Case SpoolerRegisterKeyIndexes.SPLREG_DS_PRESENT_FOR_USER
                    Return "DsPresentForUser"
                Case SpoolerRegisterKeyIndexes.SPLREG_REMOTE_FAX
                    Return "RemoteFax"
                Case SpoolerRegisterKeyIndexes.SPLREG_RESTART_JOB_ON_POOL_ERROR
                    Return "RestartJobOnPoolError"
                Case SpoolerRegisterKeyIndexes.SPLREG_RESTART_JOB_ON_POOL_ENABLED
                    Return "RestartJobOnPoolEnabled"
                Case SpoolerRegisterKeyIndexes.SPLREG_DNS_MACHINE_NAME
                    Return "DNSMachineName"
                Case Else
                    Return ""
            End Select
        End Function

    End Class
#End Region

End Namespace
