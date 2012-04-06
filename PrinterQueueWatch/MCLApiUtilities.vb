'\\ --[MCLApiUtilities]--------------------------------------------------------------
'\\ Functions to help with the windows API.
'\\ ---------------------------------------------------------------------------------
Imports System.Runtime.InteropServices
Imports System.Security.Principal

Module MCLApiUtilities

    Public Function LoggedInAsAdministrator() As Boolean
        Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim principal As New WindowsPrincipal(identity)
        Try
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch e As Exception
            Return False
        End Try

    End Function


End Module
