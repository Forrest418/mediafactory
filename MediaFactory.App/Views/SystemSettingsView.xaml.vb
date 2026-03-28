Imports System.Windows
Imports System.Windows.Controls
Imports MediaFactory.Localization

Namespace Views
    Public Class SystemSettingsView
        Inherits UserControl

        Private Sub AddScenarioPresetButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.CreateScenarioPreset()
        End Sub

        Private Sub DeleteScenarioPresetButton_Click(sender As Object, e As RoutedEventArgs)
            GetShell(Me)?.DeleteSelectedScenarioPreset()
        End Sub

        Private Sub SaveScenarioPresetButton_Click(sender As Object, e As RoutedEventArgs)
            Try
                GetShell(Me)?.SaveSelectedScenarioPreset()
                MessageBox.Show(T("Dialog.SaveScenarioSuccess", "Scenario template saved successfully."),
                                T("Dialog.Settings", "System Settings"),
                                MessageBoxButton.OK,
                                MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Dialog.SaveScenarioFailed", "Failed to save scenario template"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub SaveModelSettingsButton_Click(sender As Object, e As RoutedEventArgs)
            Try
                GetShell(Me)?.SaveModelSettings()
                MessageBox.Show(T("Dialog.SaveModelSuccess", "Model settings saved successfully."),
                                T("Dialog.Settings", "System Settings"),
                                MessageBoxButton.OK,
                                MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Dialog.SaveModelFailed", "Failed to save model settings"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub
    End Class
End Namespace
