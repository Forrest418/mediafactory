Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports MediaFactory.ViewModels

Namespace Views
    Public Class ProjectHallView
        Inherits UserControl

        Private Sub OpenProjectButton_Click(sender As Object, e As RoutedEventArgs)
            Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
            GetShell(Me)?.OpenProject(project)
        End Sub

        Private Sub ProjectCard_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            Dim current = TryCast(e.OriginalSource, DependencyObject)
            While current IsNot Nothing
                If TypeOf current Is Button Then
                    Return
                End If
                current = VisualTreeHelper.GetParent(current)
            End While

            Dim project = TryCast(TryCast(sender, FrameworkElement)?.DataContext, ProjectSessionViewModel)
            GetShell(Me)?.OpenProject(project)
        End Sub

        Private Sub AllFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetAllFilter()
        End Sub

        Private Sub PlanningFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetPlanningFilter()
        End Sub

        Private Sub RenderingFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetRenderingFilter()
        End Sub

        Private Sub ReadyFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetReadyFilter()
        End Sub

        Private Sub FailedFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetFailedFilter()
        End Sub

        Private Sub CompletedFilterButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.SetCompletedFilter()
        End Sub

        Private Async Sub RetryFailedButton_Click(sender As Object, e As RoutedEventArgs)
            Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.RetryFailedImagesForProjectAsync(project)
        End Sub
    End Class
End Namespace
