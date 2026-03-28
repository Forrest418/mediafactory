Imports System.Windows
Imports System.Windows.Controls

Namespace Views
    Public Class ProjectParameterPanel
        Inherits UserControl

        Private Async Sub GeneratePlanButton_Click(sender As Object, e As RoutedEventArgs)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.GeneratePlanForSelectedProjectAsync()
        End Sub

        Private Async Sub GenerateImagesButton_Click(sender As Object, e As RoutedEventArgs)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.GenerateImagesForSelectedProjectAsync()
        End Sub

        Private Sub OpenOutputDirectoryButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.OpenSelectedProjectOutputDirectory()
        End Sub
    End Class
End Namespace
