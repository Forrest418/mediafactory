Imports System.IO

Namespace Utils
    Public NotInheritable Class FileLocator
        Private Sub New()
        End Sub

        Public Shared Function FindUpwards(fileName As String) As String
            Dim candidates As New List(Of String) From {
                AppContext.BaseDirectory,
                Environment.CurrentDirectory
            }

            For Each startPath In candidates
                Dim directory As New DirectoryInfo(startPath)
                Dim safetyCounter As Integer = 0

                While directory IsNot Nothing AndAlso safetyCounter < 10
                    Dim fullPath = Path.Combine(directory.FullName, fileName)
                    If File.Exists(fullPath) Then
                        Return fullPath
                    End If

                    directory = directory.Parent
                    safetyCounter += 1
                End While
            Next

            Return String.Empty
        End Function
    End Class
End Namespace
