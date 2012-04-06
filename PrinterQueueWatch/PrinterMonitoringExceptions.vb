'--[PrinterMonitoringExceptions]-----------------------------------------
'\\ Namespace for all the spooler related exceptions 
'\\ (c) 2003 Merrion Computing Ltd
'\\ http://www.merrioncomputing.com
'\\ ------------------------------------------------------------------
Namespace PrinterMonitoringExceptions

#Region "InsufficentPrinterAccessRightsException"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : PrinterQueueWatch
    ''' Class	 : PrinterMonitoringExceptions.InsufficentPrinterAccessRightsException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Thrown when an attempt is made to access the printer by a process that does not
    ''' have sufficient access rights  
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Serializable()> _
    Public Class InsufficentPrinterAccessRightsException
        Inherits System.Exception

        Public Sub New()
            MyBase.New(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pem_NoPrinterAccess"))
        End Sub

        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub

        Public Sub New(ByVal innerException As System.Exception)
            MyBase.New("", innerException)
        End Sub
    End Class

#End Region

#Region "InsufficientPrintJobAccessRightsException"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : PrinterQueueWatch
    ''' Class	 : PrinterMonitoringExceptions.InsufficentPrintJobAccessRightsException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Thrown when an attempt is made to access the print job by a process that does not
    ''' have sufficient access rights
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Duncan]	20/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Serializable()>
    Public Class InsufficentPrintJobAccessRightsException
        Inherits System.Exception

        Public Sub New()
            MyBase.New(PrinterMonitorComponent.ComponentLocalisationResourceManager.GetString("pem_NoJobAccess"))
        End Sub

        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class
#End Region

#Region "PrintJobTransferException"
    ''' <summary>
    ''' An exception that is thrown when an error occurs transfering a print job from one print queue to another
    ''' </summary>
    <Serializable()>
    Public Class PrintJobTransferException
        Inherits System.Exception

        Public Sub New()
            MyBase.New(My.Resources.pem_JobTransferFailed)
        End Sub

        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class
#End Region

End Namespace
