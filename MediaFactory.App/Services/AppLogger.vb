Imports System.IO
Imports System.Text
Imports System.Threading
Imports MediaFactory.Utils

Namespace Services
    Public Class AppLogger
        Private Shared ReadOnly SyncRoot As New Object()

        Private ReadOnly _logFilePath As String

        Public Sub New()
            Dim logsRoot = Path.Combine(AppPaths.LogsRoot, DateTime.Now.ToString("yyyyMMdd"))
            Directory.CreateDirectory(logsRoot)

            _logFilePath = Path.Combine(
                logsRoot,
                $"session_{DateTime.Now:HHmmss_fff}_{Environment.ProcessId}.log")

            WriteEntry("INFO", "日志会话已启动。")
        End Sub

        Public ReadOnly Property LogFilePath As String
            Get
                Return _logFilePath
            End Get
        End Property

        Public Function Info(message As String) As String
            Return WriteEntry("INFO", message)
        End Function

        Public Function Warn(message As String) As String
            Return WriteEntry("WARN", message)
        End Function

        Public Function [Error](message As String, Optional ex As Exception = Nothing) As String
            Return WriteEntry("ERROR", message, ex)
        End Function

        Private Function WriteEntry(level As String, message As String, Optional ex As Exception = Nothing) As String
            Dim builder As New StringBuilder()
            builder.Append($"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}")

            If ex IsNot Nothing Then
                builder.AppendLine()
                builder.Append(ex.ToString())
            End If

            Dim entry = builder.ToString()

            SyncLock SyncRoot
                File.AppendAllText(_logFilePath, entry & Environment.NewLine, Encoding.UTF8)
            End SyncLock

            Return entry
        End Function
    End Class
End Namespace
