
Public Class SapCoPlExcelAddin

    Private Sub SapCoPlExcelAddin_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapCoPlExcelAddin_Shutdown() Handles Me.Shutdown

    End Sub

End Class
