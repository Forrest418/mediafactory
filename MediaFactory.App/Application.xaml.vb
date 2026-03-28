Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Threading
Imports MediaFactory.Utils

Class Application
    Private Shared ReadOnly StartupLogDirectory As String = Path.Combine(AppPaths.LogsRoot, "startup")
    Private Shared ReadOnly StartupLogPath As String = Path.Combine(StartupLogDirectory, $"{DateTime.Now:yyyyMMdd}.log")

    Protected Overrides Sub OnStartup(e As StartupEventArgs)
        MyBase.OnStartup(e)

        Directory.CreateDirectory(StartupLogDirectory)
        ShutdownMode = ShutdownMode.OnMainWindowClose

        AddHandler DispatcherUnhandledException, AddressOf Application_DispatcherUnhandledException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException
        AddHandler TaskScheduler.UnobservedTaskException, AddressOf TaskScheduler_UnobservedTaskException

        WriteStartupLog("Startup begin.")

        Try
            Dim window As New MainWindow()
            MainWindow = window
            WriteStartupLog("MainWindow created.")
            window.Show()
            WriteStartupLog("MainWindow shown.")
        Catch ex As Exception
            WriteStartupLog("Startup failed.", ex)
            MessageBox.Show(ex.ToString(),
                            "MediaFactory Startup Failed",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Shutdown(-1)
        End Try
    End Sub

    Protected Overrides Sub OnExit(e As ExitEventArgs)
        WriteStartupLog($"Application exit with code {e.ApplicationExitCode}.")
        MyBase.OnExit(e)
    End Sub

    Private Sub Application_DispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        WriteStartupLog("Dispatcher unhandled exception.", e.Exception)
    End Sub

    Private Sub CurrentDomain_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim [error] = TryCast(e.ExceptionObject, Exception)
        WriteStartupLog($"AppDomain unhandled exception. IsTerminating={e.IsTerminating}.", [error])
    End Sub

    Private Sub TaskScheduler_UnobservedTaskException(sender As Object, e As UnobservedTaskExceptionEventArgs)
        WriteStartupLog("TaskScheduler unobserved exception.", e.Exception)
    End Sub

    Private Shared Sub WriteStartupLog(message As String, Optional ex As Exception = Nothing)
        Dim builder As New StringBuilder()
        builder.Append($"[{DateTime.Now:HH:mm:ss.fff}] {message}")

        If ex IsNot Nothing Then
            builder.AppendLine()
            builder.Append(ex.ToString())
        End If

        SyncLock GetType(Application)
            File.AppendAllText(StartupLogPath, builder.ToString() & Environment.NewLine, Encoding.UTF8)
        End SyncLock
    End Sub
End Class
