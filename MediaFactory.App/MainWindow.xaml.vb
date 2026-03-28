Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports MediaFactory.ViewModels

Class MainWindow
    Inherits Window

    Private ReadOnly _shell As StudioShellViewModel

    Public Sub New()
        InitializeComponent()
        _shell = New StudioShellViewModel()
        DataContext = _shell
    End Sub

    Private Sub ProjectHallNavButton_Click(sender As Object, e As RoutedEventArgs)
        _shell.NavigateToProjectHall()
    End Sub

    Private Sub TaskQueueNavButton_Click(sender As Object, e As RoutedEventArgs)
        _shell.NavigateToTaskQueue()
    End Sub

    Private Sub MediaLibraryNavButton_Click(sender As Object, e As RoutedEventArgs)
        _shell.NavigateToMediaLibrary()
    End Sub

    Private Sub SystemSettingsNavButton_Click(sender As Object, e As RoutedEventArgs)
        _shell.NavigateToSystemSettings()
    End Sub

    Private Sub NewProjectButton_Click(sender As Object, e As RoutedEventArgs)
        _shell.CreateNewProject()
    End Sub

    Private Sub WorkspaceTabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
        _shell.OpenProject(project)
    End Sub

    Private Sub WorkspaceTabsListBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim listBox = TryCast(sender, ListBox)
        If listBox Is Nothing OrElse listBox.SelectedItem Is Nothing Then Return
        listBox.ScrollIntoView(listBox.SelectedItem)
    End Sub

    Private Sub CloseWorkspaceTabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim project = TryCast(TryCast(sender, FrameworkElement)?.Tag, ProjectSessionViewModel)
        _shell.CloseWorkspaceTab(project)
        e.Handled = True
    End Sub

    Protected Overrides Sub OnClosed(e As EventArgs)
        MyBase.OnClosed(e)
        _shell.Dispose()
    End Sub
End Class
