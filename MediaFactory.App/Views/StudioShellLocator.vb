Imports System.Windows
Imports System.Windows.Controls
Imports MediaFactory.ViewModels

Namespace Views
    Friend Module StudioShellLocator
        Friend Function GetShell(control As UserControl) As StudioShellViewModel
            Return TryCast(Window.GetWindow(control)?.DataContext, StudioShellViewModel)
        End Function

        Friend Function GetHostWindow(control As UserControl) As Window
            Return Window.GetWindow(control)
        End Function
    End Module
End Namespace
