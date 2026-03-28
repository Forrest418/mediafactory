Imports System.Windows
Imports System.Windows.Controls

Namespace Views
    Public Class ProjectWorkspaceView
        Inherits UserControl

        Private Sub BackToHallButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.NavigateToProjectHall()
        End Sub
    End Class
End Namespace
