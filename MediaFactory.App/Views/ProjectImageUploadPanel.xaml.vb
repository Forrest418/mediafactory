Imports System.Windows
Imports System.Windows.Controls
Imports MediaFactory.Models

Namespace Views
    Public Class ProjectImageUploadPanel
        Inherits UserControl

        Private Sub ImportImagesButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.ImportImagesForSelectedProject()
        End Sub

        Private Sub RemoveUploadedImageButton_Click(sender As Object, e As RoutedEventArgs)
            Dim item = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProductImageItem)
            GetShell(Me)?.RemoveUploadedImageFromSelectedProject(item)
        End Sub
    End Class
End Namespace
