Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports MediaFactory.Models

Namespace Views
    Public Class MediaLibraryView
        Inherits UserControl

        Private Sub AllMediaFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetAllMediaFilter()
        End Sub

        Private Sub SourceMediaFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetSourceMediaFilter()
        End Sub

        Private Sub GeneratedMediaFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetGeneratedMediaFilter()
        End Sub

        Private Sub FailedMediaFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetFailedMediaFilter()
        End Sub

        Private Sub OpenMediaProjectButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.OpenSelectedMediaLibraryProject()
        End Sub

        Private Sub OpenMediaFileButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.OpenSelectedMediaLibraryFile()
        End Sub

        Private Sub DownloadMediaButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.DownloadMediaLibraryItem(GetShell(Me)?.SelectedMediaLibraryItem)
        End Sub

        Private Sub PreviewMediaButton_Click(sender As Object, e As RoutedEventArgs)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            shell.PreviewSelectedMediaLibraryItem(GetHostWindow(Me))
        End Sub

        Private Sub PreviewMediaButtonInline_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, MediaLibraryItem)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            shell.PreviewMediaLibraryItem(item, GetHostWindow(Me))
        End Sub

        Private Sub OpenMediaFileButtonInline_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, MediaLibraryItem)
            GetShell(Me)?.OpenMediaLibraryFile(item)
        End Sub

        Private Sub DownloadMediaButtonInline_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, MediaLibraryItem)
            GetShell(Me)?.DownloadMediaLibraryItem(item)
        End Sub

        Private Sub OpenMediaProjectButtonInline_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, MediaLibraryItem)
            GetShell(Me)?.OpenMediaLibraryProject(item)
        End Sub

        Private Sub PreviewImageMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, MediaLibraryItem)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            shell.PreviewMediaLibraryItem(item, GetHostWindow(Me))
        End Sub
    End Class
End Namespace
