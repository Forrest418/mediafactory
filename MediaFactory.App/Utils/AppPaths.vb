Imports System.IO

Namespace Utils
    Public NotInheritable Class AppPaths
        Private Shared ReadOnly _appDataRoot As String = EnsureDirectory(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MediaFactory"))
        Private Shared ReadOnly _dataRoot As String = EnsureDirectory(Path.Combine(_appDataRoot, "Data"))
        Private Shared ReadOnly _logsRoot As String = EnsureDirectory(Path.Combine(_appDataRoot, "Logs"))
        Private Shared ReadOnly _outputsRoot As String = EnsureDirectory(Path.Combine(_appDataRoot, "Outputs"))

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property AppDataRoot As String
            Get
                Return _appDataRoot
            End Get
        End Property

        Public Shared ReadOnly Property DataRoot As String
            Get
                Return _dataRoot
            End Get
        End Property

        Public Shared ReadOnly Property LogsRoot As String
            Get
                Return _logsRoot
            End Get
        End Property

        Public Shared ReadOnly Property OutputsRoot As String
            Get
                Return _outputsRoot
            End Get
        End Property

        Private Shared Function EnsureDirectory(path As String) As String
            Directory.CreateDirectory(path)
            Return path
        End Function
    End Class
End Namespace
