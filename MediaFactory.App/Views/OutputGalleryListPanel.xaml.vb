Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports MediaFactory.Models

Namespace Views
    Public Class OutputGalleryListPanel
        Inherits UserControl

        Private Async Sub RegenerateImageButton_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, GeneratedImageItem)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.RegenerateImageForSelectedProjectAsync(item)
        End Sub

        Private Sub DownloadImageButton_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, GeneratedImageItem)
            GetShell(Me)?.DownloadImageItem(item)
        End Sub

        Private Sub EditImageButton_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, GeneratedImageItem)
            GetShell(Me)?.EditImageItem(item)
        End Sub

        Private Sub PreviewImageMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, GeneratedImageItem)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            shell.PreviewGeneratedImage(item, GetHostWindow(Me))
        End Sub
    End Class
End Namespace
