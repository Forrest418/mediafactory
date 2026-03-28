Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Markup

Namespace Localization
    Public Module L10n
        Public Function T(key As String, Optional fallback As String = Nothing) As String
            Return LocalizationManager.Instance.GetString(key, fallback)
        End Function

        Public Function F(key As String, fallback As String, ParamArray args() As Object) As String
            Return String.Format(CultureInfo.CurrentCulture, T(key, fallback), args)
        End Function
    End Module

    <MarkupExtensionReturnType(GetType(Object))>
    Public Class LocExtension
        Inherits MarkupExtension

        Public Sub New()
        End Sub

        Public Sub New(key As String)
            Me.Key = key
        End Sub

        Public Property Key As String = String.Empty

        Public Overrides Function ProvideValue(serviceProvider As IServiceProvider) As Object
            Dim binding As New Binding($"[{Key}]") With {
                .Source = LocalizationManager.Instance,
                .Mode = BindingMode.OneWay,
                .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            }

            Return binding.ProvideValue(serviceProvider)
        End Function
    End Class
End Namespace
