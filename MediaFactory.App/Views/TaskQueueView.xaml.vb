Imports System.Windows
Imports System.Windows.Controls
Imports MediaFactory.ViewModels

Namespace Views
    Public Class TaskQueueView
        Inherits UserControl

        Private Async Sub TaskQueueActionButton_Click(sender As Object, e As RoutedEventArgs)
            Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.ExecuteQueueActionAsync(project)
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

        Private Async Sub RetryAllFailedButton_Click(sender As Object, e As RoutedEventArgs)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.RetryFailedImagesForFilteredProjectsAsync()
        End Sub

        Private Async Sub StartSelectedButton_Click(sender As Object, e As RoutedEventArgs)
            Dim shell = GetShell(Me)
            If shell Is Nothing Then Return
            Await shell.StartSelectedQueueProjectsAsync()
        End Sub

        Private Sub OpenProjectButton_Click(sender As Object, e As RoutedEventArgs)
            Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
            GetShell(Me)?.OpenProjectFromQueue(project)
        End Sub

        Private Sub CloseProjectButton_Click(sender As Object, e As RoutedEventArgs)
            Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
            GetShell(Me)?.CloseProject(project)
        End Sub
    End Class
End Namespace
